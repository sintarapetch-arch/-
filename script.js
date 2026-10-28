const MAX_PHOTOS_PER_TASK = 100;
    // Increased attachment ceiling to 35 MB (maximum safe single payload for Google Apps Script)
    const MAX_ATTACHMENT_MB = 35;

    // Upload time scales directly with these two numbers. 900px keeps nameplates
    // and gauges readable in the lightbox while staying near 60-90 KB per photo;
    // raising the width to 1280 roughly doubles every save.
    const PHOTO_MAX_WIDTH = 900;
    const PHOTO_QUALITY = 0.7;
    const DEFAULT_LIVE_API_URL = "https://script.google.com/macros/s/AKfycbwQllvUEAPdbp_bV7PLTnyM9VaVog5ZYafdl1pnITi7wDyxZFHs8_l91g059FNRrhVj/exec"; // HARDCODED TASKS API

    /* -----------------------------------------------------------------------
     * SAFE HTML INTERPOLATION
     *
     * Every card, table row and modal is built with innerHTML from values that
     * staff type straight into the Google Sheet. Without escaping, a project
     * name containing a quote silently breaks the surrounding attribute, and
     * one containing a tag executes as markup. Anything interpolated from task
     * data goes through esc().
     * --------------------------------------------------------------------- */
    function esc(value) {
      return String(value === null || value === undefined ? '' : value)
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&#39;');
    }

    // Only http(s) and data: URLs may reach an src/href attribute - blocks javascript:
    function escUrl(value) {
      const raw = String(value || '').trim();
      if (!raw) return '';
      return /^(https?:|data:)/i.test(raw) ? esc(raw) : '';
    }

    /* -----------------------------------------------------------------------
     * API ENDPOINT RESOLUTION
     *
     * The saved URL used to win unconditionally, so any machine that had opened
     * the app once kept talking to the old spreadsheet no matter how many times
     * this file was updated - the built-in default could never take effect.
     *
     * We now remember WHICH shipped default a browser adopted. If the saved URL
     * is just that old default (nobody typed it), we follow the new one. A URL
     * the user actually entered in the settings dialog is always kept.
     * --------------------------------------------------------------------- */
    const API_URL_KEY = 'pts_api_url';
    const API_URL_ADOPTED_KEY = 'pts_api_url_adopted_default';

    const API_TIMEOUT_MS = 20000;
    const API_SAVE_TIMEOUT_MS = 120000;  // photo uploads legitimately take a while

    let apiUrlMigratedFrom = null;

    function normalizeApiUrl(value) {
      return String(value || '').trim().replace(/\s+/g, '');
    }

    function resolveApiUrl() {
      // EMBEDDED DIRECT LINK - ALWAYS USES DEFAULT_LIVE_API_URL ON BOTH MOBILE & PC
      try {
        localStorage.setItem(API_URL_KEY, DEFAULT_LIVE_API_URL);
        localStorage.setItem(API_URL_ADOPTED_KEY, DEFAULT_LIVE_API_URL);
      } catch (e) {}
      return DEFAULT_LIVE_API_URL;
    }

    /* -----------------------------------------------------------------------
     * FAILURE CLASSIFICATION
     *
     * A broken Apps Script connection almost always fails as *valid HTML*, not
     * as an HTTP error, so "ต่อไม่ติด" used to be all the user ever saw. These
     * codes turn each failure into one specific instruction.
     * --------------------------------------------------------------------- */
    const API_ERROR_TEXT = {
      'timeout':    'เซิร์ฟเวอร์ไม่ตอบภายในเวลาที่กำหนด — Apps Script อาจกำลังประมวลผลนาน หรือเน็ตช้า',
      'network':    'เรียก URL ไม่สำเร็จ (เน็ตขัดข้อง หรือ Google ปฏิเสธคำขอ) — มักเกิดเมื่อ Deploy ไม่ได้ตั้ง Who has access เป็น Anyone',
      'no-access':  'Google ขอให้ล็อกอินก่อน — ต้อง Deploy ใหม่โดยตั้ง "Who has access" เป็น Anyone',
      'not-found':  'Google ตอบว่า "ไม่พบเพจ" — ลิงก์ผิด หรือ deployment ถูกลบ/ยังไม่ได้กด Deploy',
      'not-json':   'ปลายทางตอบกลับมาไม่ใช่ JSON — URL อาจไม่ใช่ Web App หรือสคริปต์ error',
      'bad-url':    'รูปแบบ URL ไม่ถูกต้อง ต้องเป็น https://script.google.com/macros/s/.../exec'
    };

    function apiErrorText(code) {
      return API_ERROR_TEXT[code] || 'เชื่อมต่อไม่สำเร็จ';
    }

    function safeJsonParse(text) {
      try { return JSON.parse(text); } catch (err) { return null; }
    }

    function classifyHtmlResponse(text) {
      const body = String(text || '');
      if (/accounts\.google\.com|ServiceLogin|signin\/identifier/i.test(body)) return 'no-access';
      if (/ไม่พบเพจ|Page Not Found|Sorry, unable to open the file/i.test(body)) return 'not-found';
      return 'not-json';
    }

    // Catches the two mistakes that account for most "it just won't connect":
    // pasting the /dev test link, or pasting the script editor URL.
    function validateApiUrlShape(url) {
      const u = normalizeApiUrl(url);
      if (!/^https:\/\/script\.google\.com\/macros\/s\/[^/]+\/(exec|dev)$/.test(u)) {
        if (/\/edit|script\.google\.com\/home|\/d\//.test(u)) {
          return 'นี่คือลิงก์หน้าแก้ไขสคริปต์ ไม่ใช่ลิงก์ Web App — ต้องกด Deploy แล้วคัดลอก URL ที่ลงท้าย /exec';
        }
        return 'URL ต้องอยู่ในรูปแบบ https://script.google.com/macros/s/AKfycb.../exec';
      }
      if (u.endsWith('/dev')) {
        return 'ลิงก์ /dev ใช้ได้เฉพาะเจ้าของบัญชีที่ล็อกอินอยู่ — สำหรับใช้งานจริงต้องใช้ลิงก์ที่ลงท้าย /exec';
      }
      return '';
    }

    function delay(ms) {
      return new Promise(resolve => setTimeout(resolve, ms));
    }

    function fetchWithTimeout(url, options, timeoutMs) {
      const controller = new AbortController();
      const timer = setTimeout(() => controller.abort(), timeoutMs || API_TIMEOUT_MS);
      return fetch(url, { ...options, signal: controller.signal })
        .finally(() => clearTimeout(timer));
    }

    // HELPER: GENERATE INDUSTRIAL SVG SAMPLE PLACEHOLDER DATA URIS
    function createIndustrialSvg(title, color1, color2) {
      const svg = `<svg xmlns="http://www.w3.org/2000/svg" width="400" height="300" viewBox="0 0 400 300">
        <rect width="400" height="300" fill="${color1}"/>
        <circle cx="200" cy="130" r="60" fill="none" stroke="${color2}" stroke-width="12"/>
        <path d="M140 130 H260 M200 70 V190" stroke="${color2}" stroke-width="8"/>
        <rect x="50" y="220" width="300" height="40" rx="8" fill="${color2}"/>
        <text x="200" y="246" fill="#ffffff" font-family="sans-serif" font-size="16" font-weight="bold" text-anchor="middle">${title}</text>
      </svg>`;
      return 'data:image/svg+xml;utf8,' + encodeURIComponent(svg);
    }

    const SAMPLE_PHOTO_FIRE_PUMP = createIndustrialSvg('PTS FIRE PUMP INSPECTION', '#7f1d1d', '#b91c1c');
    const SAMPLE_PHOTO_FIRE_ALARM = createIndustrialSvg('FIRE ALARM CONTROL PANEL', '#0284c7', '#0369a1');
    const SAMPLE_PHOTO_VALVE = createIndustrialSvg('OS&Y GATE VALVE REPAIR', '#d97706', '#b45309');

    // INITIAL MOCK DATA TAILORED FOR PAKORN TECHNICAL SUPPLY LTD., PART.
    const INITIAL_SAMPLE_TASKS = [
      {
        Job_ID: "PTS-ENG-001",
        Project_Name: "โรงงานอุตสาหกรรม เคมีคอล มาบตาพุด (ระยอง)",
        Sub_Department: "PM Fire Pump",
        Technician_In_Charge: "ช่างปกรณ์ (หัวหน้าทีม 1)",
        Task_Detail: "ทดสอบประจำปี Diesel Fire Pump System 750 GPM & Electric Jockey Pump",
        Status: "เปิดงาน / มอบหมายงาน",
        JSA_Completed: "Yes",
        Priority: "High",
        Target_Date: "2026-08-05",
        PO_Approval_Date: "2026-07-01",
        Contract_Expiry_Date: "2026-08-06",
        Completion_Date: "",
        Site_Location: "นิคมมาบตาพุด อาคารปั๊มดับเพลิง Zone B",
        Site_Contact_Phone: "081-888-1234 (คุณสมศักดิ์)",
        Site_Map_Url: "https://maps.google.com/?q=Map+Ta+Phut+Industrial+Estate",
        Delivery_Doc: "[]",
        Notes_Issues: "[01/08/2026 09:00] ช่างปกรณ์: เข้าพื้นที่ติดตั้งเครื่องวัดแรงดันเรียบร้อย ปั๊มดีเซลทำงานปกติดี",
        Updated_At: "2026-08-01 09:00",
        Site_Photos: JSON.stringify([SAMPLE_PHOTO_FIRE_PUMP]),
        Document_Files: JSON.stringify([
          { name: "รายงานการตรวจเช็คปั๊มดับเพลิง_2026.pdf", size: "1.2 MB", type: "application/pdf", dataUrl: "data:application/pdf;base64,sample" },
          { name: "รายการเปลี่ยนอะไหล่ปั๊มดีเซล.xlsx", size: "45 KB", type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", dataUrl: "data:application/vnd.ms-excel;base64,sample" }
        ]),
        Video_Files: "[]"
      },
      {
        Job_ID: "PTS-ENG-002",
        Project_Name: "อาคารพาณิชย์ สีลม สแควร์ ทาวเวอร์",
        Sub_Department: "PM Fire Alarm",
        Technician_In_Charge: "ช่างสมชาย (หัวหน้าทีม 2)",
        Task_Detail: "ตรวจสอบ Smoke Detector ชั้น 15-30 และทดสอบตู้ FACP (Fire Alarm Control Panel)",
        Status: "ดำเนินการแก้ไข",
        JSA_Completed: "Yes",
        Priority: "Medium",
        Target_Date: "2026-08-07",
        PO_Approval_Date: "2026-07-10",
        Contract_Expiry_Date: "2026-08-15",
        Completion_Date: "",
        Site_Location: "สีลม กรุงเทพฯ",
        Site_Contact_Phone: "089-123-4567 (ช่างประจำอาคาร)",
        Site_Map_Url: "https://maps.google.com/?q=Silom+Square+Tower",
        Delivery_Doc: "[]",
        Notes_Issues: "[01/08/2026 10:15] ช่างสมชาย: กำลังแก้ไขวงจร Loop 2 ตู้คอนโทรล",
        Updated_At: "2026-08-01 10:15",
        Site_Photos: JSON.stringify([SAMPLE_PHOTO_FIRE_ALARM]),
        Document_Files: JSON.stringify([
          { name: "คู่มือตู้ควบคุม_FACP_Notifier.pdf", size: "2.8 MB", type: "application/pdf", dataUrl: "data:application/pdf;base64,sample" }
        ]),
        Video_Files: "[]"
      },
      {
        Job_ID: "PTS-ENG-003",
        Project_Name: "คลังสินค้า โลจิสติกส์ นิคมบางปู",
        Sub_Department: "งานโครงการ",
        Technician_In_Charge: "ช่างวิโรจน์ (หัวหน้าทีม 3)",
        Task_Detail: "เดินท่อเหล็กดับเพลิง HDG 4 นิ้ว และติดตั้งหัวกระจายน้ำดับเพลิง (Sprinkler System)",
        Status: "รอ Testing & Commissioning",
        JSA_Completed: "Yes",
        Priority: "High",
        Target_Date: "2026-08-03",
        PO_Approval_Date: "2026-06-15",
        Contract_Expiry_Date: "2026-08-04",
        Completion_Date: "",
        Site_Location: "นิคมอุตสาหกรรมบางปู สมุทรปราการ",
        Site_Contact_Phone: "086-555-7890 (คุณวิชัย)",
        Site_Map_Url: "https://maps.google.com/?q=Bangpoo+Industrial+Estate",
        Delivery_Doc: "[]",
        Notes_Issues: "[31/07/2026 16:30] ช่างวิโรจน์: ติดตั้งท่อเสร็จ 100% รอนัดวิศวกรผู้เชี่ยวชาญเข้าอัดแรงดัน Hydrostatic Test",
        Updated_At: "2026-07-31 16:30",
        Site_Photos: "[]",
        Document_Files: JSON.stringify([
          { name: "BOQ_งานติดตั้งท่อสปริงเกลอร์_นิคมบางปู.docx", size: "180 KB", type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document", dataUrl: "data:application/msword;base64,sample" }
        ]),
        Video_Files: "[]"
      },
      {
        Job_ID: "PTS-ENG-004",
        Project_Name: "โรงงานประกอบชิ้นส่วนยานยนต์ อมตะซิตี้",
        Sub_Department: "งานซ่อมบำรุง",
        Technician_In_Charge: "ช่างธีรเดช (หัวหน้าทีม 4)",
        Task_Detail: "ซ่อมบำรุงใหญ่เกทวาล์วขนาด 6 นิ้ว (OS&Y Gate Valve) และเปลี่ยนชุดปะเก็นกันรั่ว",
        Status: "รออะไหล่ชั่วคราว / ติดปัญหา",
        JSA_Completed: "No",
        Priority: "High",
        Target_Date: "2026-08-02",
        PO_Approval_Date: "2026-07-01",
        Contract_Expiry_Date: "2026-07-30",
        Completion_Date: "",
        Site_Location: "อมตะซิตี้ ชลบุรี",
        Site_Contact_Phone: "082-345-6789 (ฝ่ายซ่อมบำรุงโรงงาน)",
        Site_Map_Url: "https://maps.google.com/?q=Amata+City+Chonburi",
        Delivery_Doc: "[]",
        Notes_Issues: "[01/08/2026 08:30] ช่างธีรเดช: ตรวจพบวาล์วมีรอยแตกร้าวที่ตัวบอดี้ ต้องสั่งซื้อชุดบอดี้ใหม่จากโรงงาน รอจัดส่ง 2 วัน",
        Updated_At: "2026-08-01 08:30",
        Site_Photos: JSON.stringify([SAMPLE_PHOTO_VALVE]),
        Document_Files: "[]",
        Video_Files: "[]"
      },
      {
        Job_ID: "PTS-ENG-005",
        Project_Name: "โรงงานแปรรูปอาหาร สมุทรปราการ",
        Sub_Department: "PM Fire Pump",
        Technician_In_Charge: "ช่างปกรณ์ (หัวหน้าทีม 1)",
        Task_Detail: "ตรวจเช็คเปลี่ยนถ่ายน้ำมันเครื่อง และไส้กรองปั๊มดับเพลิงเครื่องยนต์ดีเซล",
        Status: "ค้างทำรายงาน",
        JSA_Completed: "Yes",
        Priority: "Medium",
        Target_Date: "2026-08-10",
        PO_Approval_Date: "2026-07-20",
        Contract_Expiry_Date: "2026-08-25",
        Completion_Date: "",
        Site_Location: "พระประแดง สมุทรปราการ",
        Site_Contact_Phone: "081-999-3333 (คุณประสิทธิ์)",
        Site_Map_Url: "https://maps.google.com/?q=Phra+Pradaeng+Samut+Prakan",
        Delivery_Doc: "[]",
        Notes_Issues: "[01/08/2026 11:00] ช่างปกรณ์: ตรวจเช็คเสร็จแล้ว อยู่ระหว่างจัดทำเล่มรายงานสรุปผลการทดสอบ",
        Updated_At: "2026-08-01 11:00",
        Site_Photos: "[]",
        Document_Files: "[]",
        Video_Files: "[]"
      },
      {
        Job_ID: "PTS-ENG-006",
        Project_Name: "ศูนย์ข้อมูล Data Center พระราม 9",
        Sub_Department: "งานโครงการ",
        Technician_In_Charge: "ช่างวิโรจน์ (หัวหน้าทีม 3)",
        Task_Detail: "ติดตั้งระบบดับเพลิงอัตโนมัติด้วยก๊าซ Novec 1230 ในห้อง Server Main Room",
        Status: "ส่งมอบงานแล้ว",
        JSA_Completed: "Yes",
        Priority: "Low",
        Target_Date: "2026-07-28",
        PO_Approval_Date: "2026-06-01",
        Contract_Expiry_Date: "2026-07-30",
        Completion_Date: "2026-07-28",
        Site_Location: "พระราม 9 กรุงเทพฯ",
        Site_Contact_Phone: "089-999-5678 (วิศวกรผู้ตรวจรับ)",
        Site_Map_Url: "https://maps.google.com/?q=Rama+9+Bangkok",
        Delivery_Doc: JSON.stringify([{
          name: "ใบส่งมอบงาน_และรับรองระบบดับเพลิง_พระราม9.pdf",
          size: "1.8 MB",
          type: "application/pdf",
          dataUrl: "data:application/pdf;base64,sample"
        }]),
        Notes_Issues: "[28/07/2026 17:00] ช่างวิโรจน์: ตรวจรับงานและส่งมอบเอกสาร As-Built Drawing เรียบร้อยแล้ว",
        Updated_At: "2026-07-28 17:00",
        Site_Photos: "[]",
        Document_Files: JSON.stringify([
          { name: "ใบรับรองการฉีดก๊าซดับเพลิง_Novec1230.pdf", size: "3.1 MB", type: "application/pdf", dataUrl: "data:application/pdf;base64,sample" }
        ]),
        Video_Files: "[]"
      },
      {
        Job_ID: "PTS-ENG-007",
        Project_Name: "โรงพยาบาลกรุงเทพ สุขุมวิท",
        Sub_Department: "PM Fire Alarm",
        Technician_In_Charge: "ช่างสมชาย (หัวหน้าทีม 2)",
        Task_Detail: "ตรวจเช็คประจำปีระบบแจ้งเหตุเพลิงไหม้อาคาร 1-4 และตู้กราฟิกแสดงผล",
        Status: "ปิดงาน",
        JSA_Completed: "Yes",
        Priority: "Low",
        Target_Date: "2026-07-15",
        PO_Approval_Date: "2026-06-10",
        Contract_Expiry_Date: "2026-07-20",
        Completion_Date: "2026-07-18",
        Site_Location: "สุขุมวิท กรุงเทพฯ",
        Site_Contact_Phone: "085-777-8888 (ฝ่ายอาคารสถานที่)",
        Site_Map_Url: "https://maps.google.com/?q=Sukhumvit+Bangkok",
        Delivery_Doc: JSON.stringify([{
          name: "ใบส่งมอบงานและตรวจรับงวดสุดท้าย_รพ.สุขุมวิท.pdf",
          size: "2.4 MB",
          type: "application/pdf",
          dataUrl: "data:application/pdf;base64,sample"
        }]),
        Notes_Issues: "[18/07/2026 15:00] ช่างสมชาย: ปิดงานและวางบิลเรียบร้อย สมบูรณ์ 100%",
        Updated_At: "2026-07-18 15:00",
        Site_Photos: "[]",
        Document_Files: "[]",
        Video_Files: "[]"
      }
    ];

    const ACTIVE_STATUSES = [
      "วางแผน / เตรียมอุปกรณ์",
      "ส่งหนังสือเข้างานแล้ว",
      "ระหว่างปฏิบัติงาน",
      "รอ Testing & Commissioning",
      "ติดปัญหาหน้างาน / รออะไหล่",
      "ค้างทำรายงาน",
      "ส่งมอบ / เอกสารเรียบร้อย"
    ];

    const ALL_STATUSES = [
      "วางแผน / เตรียมอุปกรณ์",
      "ส่งหนังสือเข้างานแล้ว",
      "ระหว่างปฏิบัติงาน",
      "รอ Testing & Commissioning",
      "ติดปัญหาหน้างาน / รออะไหล่",
      "ค้างทำรายงาน",
      "ส่งมอบ / เอกสารเรียบร้อย",
      "ปิดงาน"
    ];

    function normalizeStatus(status) {
      if (!status) return 'วางแผน / เตรียมอุปกรณ์';
      const s = String(status).trim();

      // Exact matches first
      if (ALL_STATUSES.includes(s)) return s;

      // 8. ปิดงาน
      if (s === 'ปิดงาน' || s.includes('ปิดงาน') || s.includes('สมบูรณ์/เก็บเข้าคลัง')) {
        return 'ปิดงาน';
      }

      // 5. ติดปัญหาหน้างาน / รออะไหล่ (Check BEFORE generic 'หน้างาน')
      if (s === 'ติดปัญหาหน้างาน / รออะไหล่' || s.includes('ติดปัญหา') || s.includes('รออะไหล่') || s.includes('ปัญหาหน้างาน')) {
        return 'ติดปัญหาหน้างาน / รออะไหล่';
      }

      // 7. ส่งมอบ / เอกสารเรียบร้อย
      if (s === 'ส่งมอบ / เอกสารเรียบร้อย' || s === 'ส่งมอบงานแล้ว' || s.includes('ส่งมอบ') || s.includes('As-Built')) {
        return 'ส่งมอบ / เอกสารเรียบร้อย';
      }

      // 4. รอ Testing & Commissioning
      if (s === 'รอ Testing & Commissioning' || s.includes('Testing') || s.includes('Commissioning') || s.includes('ทดสอบระบบ')) {
        return 'รอ Testing & Commissioning';
      }

      // 6. ค้างทำรายงาน
      if (s === 'ค้างทำรายงาน' || s.includes('ค้างทำรายงาน') || s.includes('รอทำรายงาน') || s.includes('จัดทำเล่มรายงาน')) {
        return 'ค้างทำรายงาน';
      }

      // 3. ระหว่างปฏิบัติงาน
      if (s === 'ระหว่างปฏิบัติงาน' || s.includes('ระหว่างปฏิบัติงาน') || s.includes('ติดต่อประสานงาน') || s === 'เข้าปฏิบัติงานหน้างาน' || s.includes('ปฏิบัติงาน') || s.includes('หน้างาน') || s.includes('ทีมช่างอยู่หน้างาน')) {
        return 'ระหว่างปฏิบัติงาน';
      }

      // 2. ส่งหนังสือเข้างานแล้ว
      if (s === 'ส่งหนังสือเข้างานแล้ว' || s.includes('ส่งหนังสือ') || s === 'แจ้งเข้าทำงาน' || s.includes('แจ้งเข้า') || s.includes('อนุมัติเข้าพื้นที่') || s.includes('ดำเนินการแก้ไข')) {
        return 'ส่งหนังสือเข้างานแล้ว';
      }

      // 1. วางแผน / เตรียมอุปกรณ์
      if (s === 'วางแผน / เตรียมอุปกรณ์' || s.includes('วางแผน') || s.includes('เตรียมอุปกรณ์') || s.includes('เปิดงาน') || s.includes('มอบหมายงาน')) {
        return 'วางแผน / เตรียมอุปกรณ์';
      }

      return s;
    }

    function isClosedStatus(status) {
      const norm = normalizeStatus(status);
      return norm === 'ปิดงาน';
    }

    // APPLICATION STATE WITH EMBEDDED DIRECT GOOGLE APPS SCRIPT URL
    let state = {
      tasks: [],
      currentView: 'kanban',
      dbMode: 'live',
      apiUrl: DEFAULT_LIVE_API_URL,
      lastError: null,
      currentEditingTask: null,
      currentDetailJobId: null,
      lastSyncedAt: null,
      lastDataSignature: null,
      serverRev: null,
      lightboxPhotos: [],
      lightboxIndex: 0,
      lightboxLabel: '',
      tempNewPhotos: [],
      tempNewDocs: [],
      tempNewVideos: [],
      tempNewDelivery: [],
      tempEditPhotos: [],
      tempEditDocs: [],
      tempEditVideos: [],
      tempEditDelivery: [],
      currentDeliveryJobId: null,
      deliveryPending: [],
      deliveryPendingJobId: null,
      detailPhotos: [],
      detailPhotoTask: null
    };

    // INITIALIZATION (EMBEDDED LIVE MODE DIRECT)
    window.addEventListener('DOMContentLoaded', () => {
      setTimeout(hideSplashScreen, 1500);
      loadState();
      startAutoRefresh();
      initPhotoDropZones();
      renderOptionLists();   // defaults immediately, server lists once fetched
      fetchOptions();
      fetchJobTasks();       // fills the bell badge and the due/overdue banner
      listenToServiceWorker();
      // Keep the device registration current (the owner's name can change)
      if (pushSupported()) setTimeout(() => { fetchPushStatus().then(refreshPushRegistration); }, 3000);
    });

    function loadState() {
      state.apiUrl = DEFAULT_LIVE_API_URL;
      state.dbMode = 'live';
      fetchTasks();
      updateStatusBadge();
    }

    function saveMockStorage() {
      localStorage.setItem('pts_mock_tasks', JSON.stringify(state.tasks));
    }

    function updateStatusBadge() {
      const badge = document.getElementById('connectionStatusBadge');
      const dot = document.getElementById('statusDot');
      const text = document.getElementById('statusText');
      badge.classList.remove('hidden');

      if (state.dbMode === 'mock') {
        dot.className = "w-2.5 h-2.5 rounded-full bg-amber-400 animate-pulse";
        text.textContent = "Demo Sandbox Mode";
        badge.title = '';
      } else if (state.lastError) {
        // A silent failed poll used to look identical to a healthy connection,
        // so people kept working against a board that had stopped updating.
        dot.className = "w-2.5 h-2.5 rounded-full bg-rose-500 animate-pulse";
        text.textContent = "ต่อ Google Sheets ไม่ได้ · กดเพื่อดูวิธีแก้";
        badge.title = apiErrorText(state.lastError);
      } else {
        dot.className = "w-2.5 h-2.5 rounded-full bg-emerald-400 animate-pulse";
        const t = state.lastSyncedAt;
        const pad = n => String(n).padStart(2, '0');
        text.textContent = t
          ? `Google Sheets Live · ซิงก์ ${pad(t.getHours())}:${pad(t.getMinutes())}:${pad(t.getSeconds())}`
          : "Google Sheets Live";
        badge.title = '';
      }

      badge.classList.toggle('cursor-pointer', state.dbMode !== 'mock' && !!state.lastError);
      badge.onclick = null;
    }

    // FETCH TASKS FROM API WITH JSONP & REDIRECT FOLLOW FAILSAFE
    // options.silent  - background refresh: no spinner, no sample-data fallback
    // options.fresh   - bypass the Apps Script read cache
    async function fetchTasks(options) {
      options = options || {};
      const icon = document.getElementById('refreshIcon');
      if (icon && !options.silent) icon.classList.add('fa-spin');

      if (state.dbMode === 'mock') {
        if (icon) icon.classList.remove('fa-spin');
        renderApp();
        return;
      }

      if (!state.apiUrl) {
        if (icon) icon.classList.remove('fa-spin');
        if (!options.silent) {
          alert("กรุณาระบุ Google Apps Script Web App URL ในเมนู 'ตั้งค่า API'");
          
        }
        return;
      }

      // Several fetches can be in flight at once (poll + save reconcile + manual
      // refresh). Without this counter a slow early response can overwrite the
      // board with data that is already out of date.
      const seq = ++fetchSeq;
      const { json, error } = await apiGetDetailed(options.fresh ? { fresh: '1' } : {});
      if (seq !== fetchSeq) return false;

      const ok = !!(json && json.status === 'success' && Array.isArray(json.data));

      if (ok) {
        state.lastError = null;
        if (typeof json.rev !== 'undefined') state.serverRev = json.rev;
        applyFetchedTasks(json.data, options);
        state.lastSyncedAt = new Date();
      } else {
        // Never quietly substitute demo rows for a live sheet - a technician
        // would edit a fake job and push it into the real database.
        state.lastError = error || (json && json.message) || 'not-json';
        if (!options.silent) {
          showSyncToast('เชื่อมต่อไม่สำเร็จ: ' + apiErrorText(state.lastError));
        }
      }

      updateStatusBadge();
      if (icon) icon.classList.remove('fa-spin');
      return ok;
    }

    let fetchSeq = 0;

    /* ---------------------------------------------------------------------
     * TRANSPORT: normal fetch first, JSONP fallback for Apps Script's
     * redirect/CORS quirks. Both return the parsed object (or null).
     * ------------------------------------------------------------------- */
    function buildApiUrl(params) {
      const qs = Object.keys(params || {})
        .map(k => encodeURIComponent(k) + '=' + encodeURIComponent(params[k]))
        .join('&');
      if (!qs) return state.apiUrl;
      return state.apiUrl + (state.apiUrl.includes('?') ? '&' : '?') + qs;
    }

    /**
     * Returns { json, error } - error is one of the API_ERROR_TEXT codes.
     *
     * Three things had to change to make this stable:
     *  - a timeout, so a hung request cannot leave the spinner turning forever;
     *  - one retry, because Apps Script returns a transient 500 often enough;
     *  - reading the body as TEXT first. Apps Script signals most real problems
     *    (needs login, deployment deleted) with an HTML page and HTTP 200, and
     *    resp.json() turned all of them into the same useless parse error.
     */
    async function apiGetDetailed(params) {
      if (!state.apiUrl) return { json: null, error: 'bad-url' };

      const url = buildApiUrl(params);
      let lastError = 'network';

      // 1. Try modern fast Fetch (Timeout 6s)
      try {
        const resp = await fetchWithTimeout(url, { 
          method: 'GET', 
          redirect: 'follow',
          cache: 'no-store'
        }, 6000);
        const text = await resp.text();
        const json = safeJsonParse(text);
        if (json && json.status) return { json: json, error: null };
        lastError = classifyHtmlResponse(text);
      } catch (err) {
        lastError = err.name === 'AbortError' ? 'timeout' : 'network';
      }

      // 2. High-speed JSONP fallback for Mobile Safari / iOS / CORS redirects
      try {
        const viaJsonp = await jsonpGet(params);
        if (viaJsonp && viaJsonp.status) return { json: viaJsonp, error: null };
      } catch (e) {
        // Continue to return classified error
      }

      return { json: null, error: lastError };
    }

    async function apiGet(params) {
      const result = await apiGetDetailed(params);
      return result.json;
    }

    function jsonpGet(params) {
      return new Promise((resolve) => {
        const cbName = 'pts_jsonp_' + Math.round(1e9 * Math.random());
        let settled = false;
        let timer = null;
        const script = document.createElement('script');

        const done = (value) => {
          if (settled) return;
          settled = true;
          clearTimeout(timer);
          try { delete window[cbName]; } catch(e) { window[cbName] = undefined; }
          if (script.parentNode) script.parentNode.removeChild(script);
          resolve(value);
        };

        window[cbName] = (json) => done(json || null);

        const paramsWithBuster = { ...(params || {}), callback: cbName, _ts: Date.now() };
        script.src = buildApiUrl(paramsWithBuster);
        script.async = true;
        script.onerror = () => done(null);

        const target = document.head || document.documentElement || document.body;
        target.appendChild(script);

        // Fail-safe timer for mobile networks (10s)
        timer = setTimeout(() => done(null), 10000);
      });
    }

    /**
     * Writes are NOT retried automatically on purpose: a CREATE whose response
     * was lost may well have reached the sheet, and a blind retry would file the
     * same job twice. The caller reports the failure and refreshes instead, so
     * the user can see what actually landed before deciding to press Save again.
     */
    async function apiPost(payload) {
      if (!state.apiUrl) return { json: null, error: 'bad-url' };

      try {
        const resp = await fetchWithTimeout(state.apiUrl, {
          method: 'POST',
          headers: { 'Content-Type': 'text/plain;charset=utf-8' },
          body: JSON.stringify(payload),
          redirect: 'follow'
        }, API_SAVE_TIMEOUT_MS);

        const text = await resp.text();
        const json = safeJsonParse(text);
        if (json && json.status) return { json: json, error: null };
        return { json: null, error: classifyHtmlResponse(text) };
      } catch (err) {
        return { json: null, error: err.name === 'AbortError' ? 'timeout' : 'network' };
      }
    }

    // Only re-render when the payload actually changed - avoids the whole board
    // flickering every time the background poller runs.
    function applyFetchedTasks(data, options) {
      options = options || {};
      // An empty live sheet means an empty board. Filling it with the demo rows
      // (the old behaviour) made a brand-new sheet look populated and let sample
      // jobs be saved back into the real database.
      const incoming = data;

      const signature = JSON.stringify(incoming);
      if (signature === state.lastDataSignature) return;

      const hadData = state.lastDataSignature !== null;
      const before = state.tasks.length;

      state.lastDataSignature = signature;
      state.tasks = incoming;
      renderOptionLists();   // dropdowns also list every value jobs actually use
      renderApp();
      refreshOpenDetailModal();

      // Tell the user when a change arrived from somewhere else
      if (hadData && options.silent) {
        const diff = incoming.length - before;
        if (diff > 0) showSyncToast(`เพิ่มใบงานใหม่ ${diff} รายการจากเครื่องอื่น`);
        else if (diff < 0) showSyncToast(`ลบใบงาน ${Math.abs(diff)} รายการจากเครื่องอื่น`);
        else showSyncToast('ข้อมูลถูกอัปเดตจากเครื่องอื่น');
      }
    }

    let syncToastTimer = null;
    function showSyncToast(message) {
      const toast = document.getElementById('syncToast');
      if (!toast) return;
      document.getElementById('syncToastText').textContent = message;
      toast.classList.remove('hidden', 'opacity-0');
      clearTimeout(syncToastTimer);
      syncToastTimer = setTimeout(() => {
        toast.classList.add('opacity-0');
        setTimeout(() => toast.classList.add('hidden'), 300);
      }, 3500);
    }

    // Keep an open preview modal in sync when new data arrives
    function refreshOpenDetailModal() {
      const jobId = state.currentDetailJobId;
      if (!jobId) return;
      if (document.getElementById('detailModal').classList.contains('hidden')) return;
      // Redrawing mid-task would throw away the photos the user is ticking, or
      // swap the list out from under a download that is already running.
      if (photoSelectMode || downloadToken) return;
      if (state.tasks.some(t => t.Job_ID === jobId)) openDetailModal(jobId);
    }

    /* ---------------------------------------------------------------------
     * REALTIME SYNC
     *
     * Apps Script cannot push, so the app polls a dedicated ?ping=1 endpoint
     * that returns nothing but a revision counter - it never opens the
     * spreadsheet, so it is cheap enough to run every few seconds. Full data is
     * downloaded ONLY when that number actually changes, which is what makes
     * adds/edits/deletes from another device show up almost immediately.
     *
     * The interval backs off while nothing is happening so a machine left open
     * all day does not burn through the Apps Script runtime quota.
     * ------------------------------------------------------------------- */
    const SYNC_FAST_MS = 4000;    // right after any activity
    const SYNC_IDLE_MS = 20000;   // after a couple of quiet minutes
    const SYNC_SAFETY_MS = 120000; // full refresh even if the counter looks unchanged

    let syncTimer = null;
    let quietPolls = 0;
    let lastFullFetch = 0;

    function currentSyncInterval() {
      if (quietPolls < 15) return SYNC_FAST_MS;        // ~1 min of fast polling
      if (quietPolls < 40) return SYNC_FAST_MS * 2;
      return SYNC_IDLE_MS;
    }

    // Any local action means the user is active - go back to fast polling
    function markActivity() {
      quietPolls = 0;
      scheduleSync(SYNC_FAST_MS);
    }

    function scheduleSync(delay) {
      clearTimeout(syncTimer);
      syncTimer = setTimeout(runSyncTick, delay);
    }

    async function runSyncTick() {
      if (state.dbMode === 'mock' || !state.apiUrl || document.hidden) {
        scheduleSync(currentSyncInterval());
        return;
      }

      try {
        const { json: ping, error } = await apiGetDetailed({ ping: '1' });
        const changed = ping && ping.status === 'success' &&
          typeof ping.rev !== 'undefined' && ping.rev !== state.serverRev;
        const overdue = Date.now() - lastFullFetch > SYNC_SAFETY_MS;

        if (changed || overdue || !ping) {
          quietPolls = 0;
          lastFullFetch = Date.now();
          await fetchTasks({ silent: true, fresh: changed });
          // The to-do list bumps the same revision counter, so it rides along
          // with the job refresh instead of needing a poll of its own.
          await fetchJobTasks({ silent: true });
        } else {
          quietPolls++;
          state.lastError = error || null;
          state.lastSyncedAt = new Date();
          updateStatusBadge();
        }
      } catch (err) {
        console.warn('sync tick failed', err);
      }

      scheduleSync(currentSyncInterval());
    }

    /* ---------------------------------------------------------------------
     * MOBILE UI HELPERS
     * ------------------------------------------------------------------- */
    function isMobileViewport() {
      return window.matchMedia('(max-width: 767px)').matches;
    }

    // Two search boxes (compact one on phones, wide one on desktop) share state.
    // getFilteredTasks always reads #searchInput, so mirror into it.
    function syncSearch(value) {
      const desktop = document.getElementById('searchInput');
      const mobile = document.getElementById('searchInputMobile');
      if (desktop && desktop.value !== value) desktop.value = value;
      if (mobile && mobile.value !== value) mobile.value = value;
      renderApp();
    }

    function toggleMobileFilters() {
      const group = document.getElementById('filtersGroup');
      group.classList.toggle('hidden');
      group.classList.toggle('flex');
    }

    function updateActiveFilterBadge() {
      const badge = document.getElementById('activeFilterBadge');
      if (!badge) return;
      const count = ['subDeptFilter', 'statusFilter', 'contractFilter', 'priorityFilter']
        .filter(id => {
          const el = document.getElementById(id);
          return el && el.value !== 'ALL';
        }).length;

      badge.textContent = count;
      badge.classList.toggle('hidden', count === 0);
    }

    // A phone stacks the kanban columns, so empty ones are just noise
    function collapseEmptyKanbanColumns(counts) {
      const mobile = isMobileViewport();
      ['st1', 'st2', 'st3', 'st4', 'st5', 'st6'].forEach((key, i) => {
        const body = document.getElementById('col-' + key);
        if (!body) return;
        const wrapper = body.closest('.kanban-col');
        if (!wrapper) return;
        wrapper.classList.toggle('hidden', mobile && counts[i] === 0);
      });
    }

    function startAutoRefresh() {
      lastFullFetch = Date.now();
      scheduleSync(SYNC_FAST_MS);

      // Re-evaluate the collapsed columns when the phone is rotated
      window.addEventListener('resize', () => {
        clearTimeout(window.__ptsResizeTimer);
        window.__ptsResizeTimer = setTimeout(renderApp, 200);
      });

      // Catch up the instant the user returns to the tab / window
      document.addEventListener('visibilitychange', () => {
        if (!document.hidden) markActivity();
      });
      window.addEventListener('focus', markActivity);
      window.addEventListener('online', markActivity);
    }

    const ATTACHMENT_FIELDS = ['Site_Photos', 'Document_Files', 'Video_Files', 'Delivery_Doc'];

    /**
     * Splits a task into the text fields and the attachment fields, and reports
     * whether any attachment still holds a Base64 payload that has to be pushed
     * to Drive.
     */
    function splitAttachments(task) {
      const core = { ...task };
      const attach = { Job_ID: task.Job_ID };
      let pendingFiles = 0;

      ATTACHMENT_FIELDS.forEach(key => {
        attach[key] = task[key];
        delete core[key];
        getTaskAttachments(task, key).forEach(item => {
          const src = typeof item === 'string' ? item : (item && (item.dataUrl || item.url)) || '';
          if (String(src).indexOf('data:') === 0) pendingFiles++;
        });
      });

      return { core, attach, pendingFiles };
    }

    /**
     * SAVE - TWO PHASES
     *
     * Uploading files to Drive is what makes a save feel broken: the request sits
     * open for the whole upload, so pressing Save with ten photos meant staring
     * at a frozen dialog for many seconds with no sign anything was happening.
     *
     * The text fields are therefore committed on their own first - that request
     * touches nothing but the spreadsheet and comes back in about a second, so
     * the dialog closes and the job is already on the board. The files are then
     * pushed by a second request that runs in the background.
     */
    async function saveTaskToDb(task, isNew = false) {
      if (state.dbMode === 'mock') {
        if (isNew) {
          state.tasks.unshift(task);
        } else {
          const idx = state.tasks.findIndex(t => t.Job_ID === task.Job_ID);
          if (idx !== -1) state.tasks[idx] = task;
        }
        saveMockStorage();
        renderApp();
        return true;
      }

      const split = splitAttachments(task);
      const deferFiles = split.pendingFiles > 0;

      const { json, error } = await apiPost({
        action: isNew ? 'CREATE' : 'UPDATE',
        data: deferFiles ? split.core : task
      });

      if (!json) {
        state.lastError = error;
        updateStatusBadge();
        alert("บันทึกลง Google Sheets ไม่สำเร็จ\n\n" + apiErrorText(error) +
          "\n\n⚠️ ข้อมูลอาจถูกบันทึกไปแล้วบางส่วน กด 🔄 รีเฟรชเพื่อตรวจสอบก่อนกดบันทึกซ้ำ " +
          "(กดซ้ำทันทีอาจได้ใบงานซ้ำสองใบ)");
        return false;
      }

      state.lastError = null;
      if (json.status !== 'success') {
        alert("บันทึกลง Google Sheets ไม่สำเร็จ: " + (json.message || "Unknown error"));
        return false;
      }

      // Show the change immediately instead of waiting for a full re-download
      const saved = { ...task };
      if (isNew && json.job_id) saved.Job_ID = json.job_id;
      const idx = state.tasks.findIndex(t => t.Job_ID === saved.Job_ID);
      if (idx !== -1) state.tasks[idx] = saved;
      else state.tasks.unshift(saved);
      state.lastDataSignature = null;
      renderApp();

      reportSaveWarnings(json.warnings);

      if (deferFiles) {
        split.attach.Job_ID = saved.Job_ID;   // CREATE may have re-issued the id
        uploadAttachmentsInBackground(split.attach, split.pendingFiles);
      } else {
        // Reconcile with the server (Drive URLs, timestamps) in the background
        fetchTasks({ silent: true, fresh: true }).then(markActivity);
      }
      return true;
    }

    function parseAttachmentValue(value) {
      if (Array.isArray(value)) return value;
      try {
        const parsed = JSON.parse(value || '[]');
        return Array.isArray(parsed) ? parsed : [];
      } catch (err) {
        return [];
      }
    }

    /**
     * Phase two - uploads the files IN BATCHES.
     *
     * Sending every attachment in one request is what made saving fail outright:
     * a handful of photos plus one clip is tens of megabytes once Base64-encoded,
     * and Apps Script answers an oversized POST with an HTML error page instead
     * of JSON ("ไม่พบเพจ"). Nothing about the message said "too big", so it looked
     * like a broken deployment.
     *
     * Each request now carries only a few megabytes: the files already stored
     * (as short Drive URLs, echoed back by the server) plus the next few new
     * ones. Only Job_ID and the attachment columns are sent, so the server keeps
     * every other column exactly as it found it.
     */
    const UPLOAD_BATCH_CHARS = 4000000;   // ~4 MB of Base64 per request
    // Drive takes roughly a second per file on the server side. Capping the
    // file count keeps every request comfortably inside API_SAVE_TIMEOUT_MS
    // even when someone stages 60 small photos at once.
    const UPLOAD_BATCH_FILES = 12;
    const UPLOAD_BATCH_RETRIES = 2;       // transient timeouts / network blips are retried silently

    function showUploadProgress(title, filename, percent, bytesText) {
      const modal = document.getElementById('uploadProgressModal');
      if (!modal) return;
      const tEl = document.getElementById('uploadProgressTitle');
      const fEl = document.getElementById('uploadProgressFilename');
      const bEl = document.getElementById('uploadProgressBar');
      const pEl = document.getElementById('uploadProgressPercent');
      const byEl = document.getElementById('uploadProgressBytes');

      if (tEl) tEl.textContent = title || 'กำลังอัปโหลดไฟล์...';
      if (fEl) fEl.textContent = filename || '';
      if (bEl) bEl.style.width = Math.min(100, Math.max(0, percent)) + '%';
      if (pEl) pEl.textContent = Math.round(percent) + '%';
      if (byEl) byEl.textContent = bytesText || '';
      modal.classList.remove('hidden');
    }

    function hideUploadProgress() {
      const modal = document.getElementById('uploadProgressModal');
      if (modal) modal.classList.add('hidden');
    }

    function promptAddDriveVideo(mode) {
      const url = prompt("วางลิงก์ไฟล์วิดีโอจาก Google Drive\n(รองรับไฟล์ขนาดใหญ่ 500MB - 1GB - 5GB+ ได้ไม่จำกัด):\n\nเช่น https://drive.google.com/file/d/.../view");
      if (!url || !url.trim()) return;
      const cleanUrl = url.trim();
      const videoEntry = {
        name: `Google Drive Video (${new Date().toLocaleTimeString('th-TH')})`,
        url: cleanUrl,
        type: 'video/drive',
        size: 'Drive Cloud'
      };

      if (mode === 'new') {
        state.tempNewVideos = state.tempNewVideos || [];
        state.tempNewVideos.push(videoEntry);
        renderVideosPreview('new');
      } else {
        state.tempEditVideos = state.tempEditVideos || [];
        state.tempEditVideos.push(videoEntry);
        renderVideosPreview('edit');
      }
      alert("✅ เพิ่มลิงก์วิดีโอจาก Google Drive เรียบร้อยแล้ว สามารถกดบันทึกใบงานได้เลยครับ");
    }

    async function uploadAttachmentsInBackground(attachFields, fileCount) {
      const jobId = attachFields.Job_ID;
      const settled = {};
      const queue = [];

      let totalBytesToUpload = 0;
      ATTACHMENT_FIELDS.forEach(field => {
        settled[field] = [];
        parseAttachmentValue(attachFields[field]).forEach(item => {
          const src = typeof item === 'string' ? item : (item && (item.dataUrl || item.url)) || '';
          if (String(src).indexOf('data:') === 0) {
            const charLen = String(src).length;
            const approxBytes = Math.round(charLen * 0.75);
            totalBytesToUpload += approxBytes;
            queue.push({ field, item, chars: charLen, approxBytes: approxBytes, name: (item && item.name) || 'ไฟล์แนบ' });
          } else {
            settled[field].push(item);
          }
        });
      });

      let uploadedBytes = 0;
      let uploadedCount = 0;

      if (queue.length > 0) {
        showUploadProgress('กำลังเริ่มต้นอัปโหลดไฟล์...', `รหัสงาน: ${jobId}`, 0, `0 B / ${formatBytes(totalBytesToUpload)}`);
      }

      while (queue.length > 0) {
        const batch = [];
        let chars = 0;
        let batchBytes = 0;
        // Always take at least one, even if that single file is over the budget
        while (queue.length > 0 && (batch.length === 0 ||
               (batch.length < UPLOAD_BATCH_FILES && chars + queue[0].chars <= UPLOAD_BATCH_CHARS))) {
          const next = queue.shift();
          chars += next.chars;
          batchBytes += next.approxBytes;
          batch.push(next);
        }

        const currentFileName = batch.length > 3
          ? `${batch.length} ไฟล์ (${batch[0].name} ... ${batch[batch.length - 1].name})`
          : batch.map(b => b.name).join(', ');
        const percent = Math.min(95, (uploadedBytes / Math.max(1, totalBytesToUpload)) * 100);
        showUploadProgress(`กำลังส่งข้อมูล (${uploadedCount + batch.length}/${fileCount} ไฟล์)...`, currentFileName, percent, `${formatBytes(uploadedBytes)} / ${formatBytes(totalBytesToUpload)}`);

        const data = { Job_ID: jobId };
        ATTACHMENT_FIELDS.forEach(field => {
          const additions = batch.filter(b => b.field === field).map(b => b.item);
          data[field] = JSON.stringify(settled[field].concat(additions));
        });

        // The payload always re-sends what is already settled, so re-posting the
        // same batch after a timeout can never duplicate a file on the sheet.
        let attempt = 0;
        let json = null;
        let error = null;
        while (attempt <= UPLOAD_BATCH_RETRIES) {
          if (attempt > 0) {
            showUploadProgress(`เชื่อมต่อขัดข้อง กำลังลองส่งใหม่ (ครั้งที่ ${attempt}/${UPLOAD_BATCH_RETRIES})...`, currentFileName, percent, `${formatBytes(uploadedBytes)} / ${formatBytes(totalBytesToUpload)}`);
            await delay(1500 * attempt);
          }
          const res = await apiPost({ action: 'UPDATE', data: data });
          json = res.json;
          error = res.error;
          const transient = !json && (error === 'timeout' || error === 'network');
          if (!transient) break;
          attempt++;
        }

        if (!json || json.status !== 'success') {
          hideUploadProgress();
          state.lastError = error || 'not-json';
          updateStatusBadge();
          alert("บันทึกใบงานสำเร็จแล้ว แต่อัปโหลดไฟล์แนบไม่สำเร็จ\n\n" +
            (json && json.message ? json.message : apiErrorText(error)) +
            `\n\nอัปโหลดสำเร็จแล้ว ${uploadedCount} จาก ${fileCount} ไฟล์ — ข้อมูลใบงานและไฟล์ที่ขึ้นไปแล้วไม่หาย` +
            "\nเปิดใบงานนี้แล้วแนบไฟล์ที่เหลือใหม่อีกครั้งได้เลย");
          fetchTasks({ silent: true, fresh: true }).then(markActivity);
          return;
        }

        reportSaveWarnings(json.warnings);
        uploadedBytes += batchBytes;
        uploadedCount += batch.length;

        const updatedPercent = Math.min(100, (uploadedBytes / Math.max(1, totalBytesToUpload)) * 100);
        showUploadProgress(`อัปโหลดแล้ว (${uploadedCount}/${fileCount} ไฟล์)...`, 'กำลังบันทึกลง Google Drive...', updatedPercent, `${formatBytes(uploadedBytes)} / ${formatBytes(totalBytesToUpload)}`);

        if (json.attachments) {
          ATTACHMENT_FIELDS.forEach(field => {
            if (typeof json.attachments[field] !== 'undefined') {
              settled[field] = parseAttachmentValue(json.attachments[field]);
            }
          });
        } else if (queue.length > 0) {
          hideUploadProgress();
          alert("อัปโหลดไฟล์แนบต่อไม่ได้\n\nสคริปต์ฝั่ง Google ยังเป็นเวอร์ชันเก่า " +
            "กรุณาวาง Code.gs ใหม่แล้ว Deploy > Manage deployments > New version\n\n" +
            `(อัปโหลดสำเร็จแล้ว ${uploadedCount} จาก ${fileCount} ไฟล์)`);
          fetchTasks({ silent: true, fresh: true }).then(markActivity);
          return;
        }
      }

      showUploadProgress('✅ อัปโหลดไฟล์แนบและวิดีโอสมบูรณ์!', 'บันทึกเข้า Google Drive เรียบร้อย', 100, `${formatBytes(totalBytesToUpload)} / ${formatBytes(totalBytesToUpload)}`);
      setTimeout(() => {
        hideUploadProgress();
        showSyncToast(`อัปโหลดไฟล์แนบและวิดีโอเรียบร้อย (${fileCount} ไฟล์)`);
      }, 1000);

      fetchTasks({ silent: true, fresh: true }).then(markActivity);
    }

    function reportSaveWarnings(warnings) {
      if (!Array.isArray(warnings) || warnings.length === 0) return;
      const unique = [...new Set(warnings.map(w => String(w).replace(/\s*\(https?:\/\/\S+\)/g, '')))];
      const shown = unique.slice(0, 3).join("\n\n");
      const more = unique.length > 3 ? `\n\n(และอีก ${unique.length - 3} รายการ)` : '';
      alert("⚠️ บันทึกข้อมูลสำเร็จ แต่ไฟล์แนบยังไม่ถูกอัปโหลด\n\n" + shown + more);
    }

    async function withBusy(buttonId, label, fn) {
      const btn = document.getElementById(buttonId);
      const original = btn ? btn.innerHTML : null;
      if (btn) {
        btn.disabled = true;
        btn.classList.add('opacity-70', 'cursor-wait');
        btn.innerHTML = `<i class="fa-solid fa-circle-notch fa-spin mr-1.5"></i>${esc(label)}`;
      }
      try {
        return await fn();
      } finally {
        if (btn) {
          btn.disabled = false;
          btn.classList.remove('opacity-70', 'cursor-wait');
          btn.innerHTML = original;
        }
      }
    }

    async function deleteTaskFromDb(jobId) {
      if (state.dbMode === 'mock') {
        state.tasks = state.tasks.filter(t => t.Job_ID !== jobId);
        saveMockStorage();
        renderApp();
        return true;
      }

      const { json, error } = await apiPost({ action: 'DELETE', job_id: jobId });

      if (!json) {
        state.lastError = error;
        updateStatusBadge();
        alert("ลบข้อมูลไม่สำเร็จ\n\n" + apiErrorText(error) + "\n\nกด 🔄 รีเฟรชเพื่อตรวจสอบสถานะล่าสุด");
        return false;
      }

      state.lastError = null;
      if (json.status === 'success') {
        state.tasks = state.tasks.filter(t => t.Job_ID !== jobId);
        state.lastDataSignature = null;
        renderApp();
        fetchTasks({ silent: true, fresh: true }).then(markActivity);
        return true;
      }

      alert("ลบข้อมูลไม่สำเร็จ: " + (json.message || "Unknown error"));
      return false;
    }

    function parseDateOnly(value) {
      if (!value) return null;
      if (value instanceof Date) return new Date(value.getFullYear(), value.getMonth(), value.getDate());

      const s = String(value).trim();
      if (!s) return null;

      // 1. ISO format: YYYY-MM-DD or YYYY/MM/DD
      let m = /^(\d{4})[-\/.](\d{1,2})[-\/.](\d{1,2})/.exec(s);
      if (m) {
        let y = +m[1];
        if (y > 2400) y -= 543; // Handle Thai Buddhist Era
        return new Date(y, +m[2] - 1, +m[3]);
      }

      // 2. Thai / UK standard: DD/MM/YYYY or DD-MM-YYYY
      m = /^(\d{1,2})[-\/.](\d{1,2})[-\/.](\d{4})/.exec(s);
      if (m) {
        let y = +m[3];
        if (y > 2400) y -= 543; // Handle Thai Buddhist Era
        return new Date(y, +m[2] - 1, +m[1]);
      }

      const d = new Date(s);
      if (isNaN(d.getTime())) return null;
      let y = d.getFullYear();
      if (y > 2400) y -= 543;
      return new Date(y, d.getMonth(), d.getDate());
    }

    function formatDisplayDate(value) {
      if (!value) return '-';
      const d = parseDateOnly(value);
      if (!d) return String(value);
      const pad = n => String(n).padStart(2, '0');
      const day = pad(d.getDate());
      const month = pad(d.getMonth() + 1);
      const year = d.getFullYear();
      return `${day}/${month}/${year}`;
    }

    function formatDateForInput(value) {
      if (!value) return '';
      const d = parseDateOnly(value);
      if (!d) return '';
      const pad = n => String(n).padStart(2, '0');
      const day = pad(d.getDate());
      const month = pad(d.getMonth() + 1);
      const year = d.getFullYear();
      return `${year}-${month}-${day}`;
    }

    function getDaysUntilExpiry(expiryDateStr) {
      const expDate = parseDateOnly(expiryDateStr);
      if (!expDate) return null;

      const today = new Date();
      today.setHours(0, 0, 0, 0);

      return Math.round((expDate - today) / (1000 * 60 * 60 * 24));
    }

    function getFilteredTasks() {
      const search = (document.getElementById('searchInput') ? document.getElementById('searchInput').value : '').toLowerCase().trim();
      const searchM = (document.getElementById('searchInputMobile') ? document.getElementById('searchInputMobile').value : '').toLowerCase().trim();
      const query = search || searchM;
      const subDept = document.getElementById('subDeptFilter') ? document.getElementById('subDeptFilter').value : 'ALL';
      const contractFilter = document.getElementById('contractFilter') ? document.getElementById('contractFilter').value : 'ALL';
      const priority = document.getElementById('priorityFilter') ? document.getElementById('priorityFilter').value : 'ALL';
      const statusFilterVal = state.currentActiveStatus || 'ACTIVE_ALL';

      const list = (state.tasks || []).filter(t => {
        const isClosed = isClosedStatus(t.Status);
        const norm = normalizeStatus(t.Status);

        // Separate closed vs active jobs / Contract Expiry filter:
        if (statusFilterVal === 'CONTRACT_EXP') {
          if (isClosed) return false;
          const days = getDaysUntilExpiry(t.Contract_Expiry_Date);
          if (days === null || days > 7) return false;
        } else if (statusFilterVal === 'ACTIVE_ALL' || statusFilterVal === 'ALL') {
          if (isClosed) return false; // Exclude closed tasks to show only active/unfinished tasks
        } else if (statusFilterVal === 'ปิดงาน') {
          if (!isClosed) return false; // Show only closed tasks
        } else {
          if (norm !== normalizeStatus(statusFilterVal)) return false;
        }

        const matchSearch = !query || 
          (t.Job_ID && t.Job_ID.toLowerCase().includes(query)) ||
          (t.Project_Name && t.Project_Name.toLowerCase().includes(query)) ||
          (t.Technician_In_Charge && t.Technician_In_Charge.toLowerCase().includes(query)) ||
          (t.Site_Location && t.Site_Location.toLowerCase().includes(query)) ||
          (t.Site_Contact_Phone && t.Site_Contact_Phone.toLowerCase().includes(query)) ||
          (t.Task_Detail && t.Task_Detail.toLowerCase().includes(query)) ||
          (t.Notes_Issues && t.Notes_Issues.toLowerCase().includes(query));

        const matchSubDept = subDept === 'ALL' || t.Sub_Department === subDept;
        const matchPriority = priority === 'ALL' || t.Priority === priority;

        let matchContract = true;
        const daysRem = getDaysUntilExpiry(t.Contract_Expiry_Date);
        if (contractFilter === 'EXPIRING_SOON') {
          matchContract = daysRem !== null && daysRem >= 0 && daysRem <= 7;
        } else if (contractFilter === 'EXPIRED') {
          matchContract = daysRem !== null && daysRem < 0;
        }

        return matchSearch && matchSubDept && matchPriority && matchContract;
      });

      const sortBy = document.getElementById('sortByFilter') ? document.getElementById('sortByFilter').value : 'ID_DESC';

      // SORTING LOGIC
      list.sort((a, b) => {
        switch (sortBy) {
          case 'ID_ASC': {
            const numA = parseInt((/(\d+)\s*$/.exec(String(a.Job_ID || '')) || [0, 0])[1], 10);
            const numB = parseInt((/(\d+)\s*$/.exec(String(b.Job_ID || '')) || [0, 0])[1], 10);
            return numA - numB;
          }
          case 'ID_DESC': {
            const numA = parseInt((/(\d+)\s*$/.exec(String(a.Job_ID || '')) || [0, 0])[1], 10);
            const numB = parseInt((/(\d+)\s*$/.exec(String(b.Job_ID || '')) || [0, 0])[1], 10);
            return numB - numA;
          }
          case 'EXPIRY_ASC': {
            const dateA = parseDateOnly(a.Contract_Expiry_Date);
            const dateB = parseDateOnly(b.Contract_Expiry_Date);
            if (!dateA && !dateB) return 0;
            if (!dateA) return 1;
            if (!dateB) return -1;
            return dateA - dateB;
          }
          case 'EXPIRY_DESC': {
            const dateA = parseDateOnly(a.Contract_Expiry_Date);
            const dateB = parseDateOnly(b.Contract_Expiry_Date);
            if (!dateA && !dateB) return 0;
            if (!dateA) return 1;
            if (!dateB) return -1;
            return dateB - dateA;
          }
          case 'PO_DESC': {
            const dateA = parseDateOnly(a.PO_Approval_Date);
            const dateB = parseDateOnly(b.PO_Approval_Date);
            if (!dateA && !dateB) return 0;
            if (!dateA) return 1;
            if (!dateB) return -1;
            return dateB - dateA;
          }
          case 'NAME_ASC': {
            return String(a.Project_Name || '').localeCompare(String(b.Project_Name || ''), 'th');
          }
          case 'NAME_DESC': {
            return String(b.Project_Name || '').localeCompare(String(a.Project_Name || ''), 'th');
          }
          case 'PRIORITY_DESC': {
            const prioWeight = { 'High': 3, 'Medium': 2, 'Low': 1 };
            const wA = prioWeight[a.Priority] || 0;
            const wB = prioWeight[b.Priority] || 0;
            return wB - wA;
          }
          default:
            return 0;
        }
      });

      return list;
    }

    function handleStatusFilterChange() {
      renderApp();
    }
    function hideSplashScreen() {
      const splash = document.getElementById('appSplashScreen');
      if (splash) {
        splash.classList.add('opacity-0', 'pointer-events-none');
        setTimeout(() => { try { splash.remove(); } catch(e){} }, 600);
      }
    }
    window.hideSplashScreen = hideSplashScreen;

    function renderApp() {
      hideSplashScreen();
      const filtered = getFilteredTasks();
      updateMetrics();
      updateActiveFilterBadge();

      // Update table counter badge
      const countEl = document.getElementById('tableCountNum');
      if (countEl) countEl.textContent = filtered.length;

      renderTable(filtered);
    }

    // UPDATE METRICS DASHBOARD CARDS & SHORTCUT COUNTS
    function updateMetrics() {
      const all = state.tasks || [];
      const activeTasks = all.filter(t => !isClosedStatus(t.Status));

      const planCount = all.filter(t => normalizeStatus(t.Status) === 'วางแผน / เตรียมอุปกรณ์').length;
      const clearedCount = all.filter(t => normalizeStatus(t.Status) === 'ส่งหนังสือเข้างานแล้ว').length;
      const onSiteCount = all.filter(t => normalizeStatus(t.Status) === 'ระหว่างปฏิบัติงาน').length;
      const testingCount = all.filter(t => normalizeStatus(t.Status) === 'รอ Testing & Commissioning').length;
      const blockedCount = all.filter(t => normalizeStatus(t.Status) === 'ติดปัญหาหน้างาน / รออะไหล่').length;
      const reportCount = all.filter(t => normalizeStatus(t.Status) === 'ค้างทำรายงาน').length;
      const deliveredCount = all.filter(t => normalizeStatus(t.Status) === 'ส่งมอบ / เอกสารเรียบร้อย').length;
      const closedCount = all.filter(t => normalizeStatus(t.Status) === 'ปิดงาน').length;

      const expiringCount = activeTasks.filter(t => {
        const days = getDaysUntilExpiry(t.Contract_Expiry_Date);
        return days !== null && days <= 7;
      }).length;

      const setVal = (id, val) => {
        const el = document.getElementById(id);
        if (el) el.textContent = val;
      };

      setVal('metricTotal', activeTasks.length);
      setVal('activeCountBadge', activeTasks.length);
      setVal('closedCountBadge', closedCount);

      setVal('metricPlan', planCount);
      setVal('metricCleared', clearedCount);
      setVal('metricOnSite', onSiteCount);
      setVal('metricTesting', testingCount);
      setVal('metricBlocked', blockedCount);
      setVal('metricReport', reportCount);
      setVal('metricExpiringContract', expiringCount);
      setVal('metricDelivered', deliveredCount);
      setVal('metricClosed', closedCount);
    }

    // QUICK STATUS FILTER & JUMP SHORTCUTS
    function filterByStatusQuick(statusVal) {
      state.currentActiveStatus = statusVal;

      // Reset contract filter to avoid sticking
      const cEl = document.getElementById('contractFilter');
      if (cEl) cEl.value = 'ALL';

      updateDashboardActiveCards(statusVal);

      renderApp();
      window.scrollTo({ top: 0, behavior: 'smooth' });
    }

    function filterByContractQuick(contractVal) {
      state.currentActiveStatus = 'CONTRACT_EXP';
      const cEl = document.getElementById('contractFilter');
      if (cEl) cEl.value = contractVal;

      updateDashboardActiveCards('CONTRACT_EXP');
      renderApp();
      window.scrollTo({ top: 0, behavior: 'smooth' });
    }

    const quickFilterStatus = filterByStatusQuick;
    const quickFilterContract = filterByContractQuick;

    function updateDashboardActiveCards(activeStatus) {
      const cardIds = [
        { id: 'card-metric-ALL', key: 'ACTIVE_ALL' },
        { id: 'card-metric-1', key: 'วางแผน / เตรียมอุปกรณ์' },
        { id: 'card-metric-2', key: 'ส่งหนังสือเข้างานแล้ว' },
        { id: 'card-metric-3', key: 'ระหว่างปฏิบัติงาน' },
        { id: 'card-metric-4', key: 'รอ Testing & Commissioning' },
        { id: 'card-metric-5', key: 'ติดปัญหาหน้างาน / รออะไหล่' },
        { id: 'card-metric-6', key: 'ค้างทำรายงาน' },
        { id: 'card-metric-7', key: 'ส่งมอบ / เอกสารเรียบร้อย' },
        { id: 'card-metric-8', key: 'ปิดงาน' },
        { id: 'card-metric-EXP', key: 'CONTRACT_EXP' }
      ];

      cardIds.forEach(item => {
        const el = document.getElementById(item.id);
        if (!el) return;
        const isActive = activeStatus === item.key || (item.key === 'ACTIVE_ALL' && (!activeStatus || activeStatus === 'ALL' || activeStatus === 'ACTIVE_ALL'));
        if (isActive) {
          el.classList.add('ring-2', 'ring-pts-800', 'bg-pts-50/70', 'shadow-md');
          el.classList.remove('bg-white');
        } else {
          el.classList.remove('ring-2', 'ring-pts-800', 'bg-pts-50/70', 'shadow-md');
          if (item.id !== 'card-metric-EXP') el.classList.add('bg-white');
        }
      });

      // Update active banner
      const banner = document.getElementById('activeStatusBanner');
      const title = document.getElementById('activeStatusBannerTitle');
      if (banner && title) {
        if (!activeStatus || activeStatus === 'ALL' || activeStatus === 'ACTIVE_ALL') {
          banner.classList.add('hidden');
        } else if (activeStatus === 'CONTRACT_EXP') {
          banner.classList.remove('hidden');
          title.textContent = '⚠️ สัญญาใกล้ครบกำหนด (≤ 7 วัน)';
        } else if (activeStatus === 'ปิดงาน') {
          banner.classList.remove('hidden');
          title.textContent = '8. ปิดงาน (Closed - เสร็จสมบูรณ์แล้ว)';
        } else {
          banner.classList.remove('hidden');
          title.textContent = activeStatus;
        }
      }
    }

    // SAFE PARSER FOR ATTACHMENT ARRAYS
    function getTaskAttachments(task, key) {
      if (!task || !task[key]) return [];
      try {
        if (typeof task[key] === 'string') {
          const parsed = JSON.parse(task[key]);
          return Array.isArray(parsed) ? parsed : [];
        }
        if (Array.isArray(task[key])) return task[key];
      } catch(e) {
        return [];
      }
      return [];
    }

    function getTaskDeliveryFiles(task) {
      return getTaskAttachments(task, 'Delivery_Doc');
    }

    // MEDIA SOURCE RESOLVER
    function photoSrc(entry) {
      const raw = typeof entry === 'string' ? entry : (entry && (entry.url || entry.dataUrl)) || '';
      if (!raw) return '';
      if (raw.indexOf('data:') === 0) return raw;

      const id = getDriveFileId(raw);
      if (id && raw.indexOf('drive.google.com/thumbnail') === -1) {
        return 'https://drive.google.com/thumbnail?id=' + id + '&sz=w1600';
      }
      return raw;
    }

    function fileSrc(item) {
      if (!item) return '';
      if (typeof item === 'string') return item;
      return item.url || item.dataUrl || '';
    }

    function getDriveFileId(url) {
      if (!url) return '';
      let m = /\/file\/d\/([a-zA-Z0-9_-]+)/.exec(url);
      if (m) return m[1];
      m = /[?&]id=([a-zA-Z0-9_-]+)/.exec(url);
      if (m) return m[1];
      return '';
    }

    function driveOpenUrl(url) {
      const id = getDriveFileId(url);
      return id ? 'https://drive.google.com/file/d/' + id + '/view' : '';
    }

    function drivePreviewUrl(url) {
      const id = getDriveFileId(url);
      return id ? 'https://drive.google.com/file/d/' + id + '/preview' : '';
    }

    function videoThumb(item) {
      if (!item) return '';
      if (item.thumb) return item.thumb;
      const id = getDriveFileId(fileSrc(item));
      return id ? 'https://drive.google.com/thumbnail?id=' + id + '&sz=w400' : '';
    }

    function isLocalFile(src) {
      return !!src && src.indexOf('data:') === 0;
    }

    function openVideoPreview(item) {
      const src = fileSrc(item);
      if (!src) return;
      const name = (item && item.name) || 'Video';

      if (isLocalFile(src)) {
        openLightboxModal('video', src, name);
        return;
      }

      const embed = drivePreviewUrl(src);
      if (embed) {
        openLightboxModal('drive', embed, name, driveOpenUrl(src));
      } else {
        window.open(src, '_blank', 'noopener');
      }
    }

    // KEYBOARD SHORTCUTS FOR MODALS
    window.addEventListener('keydown', (e) => {
      if (e.key === 'Escape') {
        closeCreateModal();
        closeEditModal();
        closeDetailModal();
        closeLightboxModal();
        closeConfigModal();
        closeDeliveryModal();
      }
    });

    // RENDER TABLE VIEW
    function renderTable(tasks) {
      const tbody = document.getElementById('tableBody');
      tbody.innerHTML = '';

      if (tasks.length === 0) {
        tbody.innerHTML = `
          <tr>
            <td colspan="10" class="text-center py-8 text-slate-400 text-sm"
              ไม่พบรายการข้อมูลตามเงื่อนไขที่เลือก
            </td>
          </tr>`;
        return;
      }

      tasks.forEach(task => {
        const tr = document.createElement('tr');
        tr.className = "hover:bg-slate-50/80 transition cursor-pointer";
        // Row click = read-only preview (edit is an explicit button in the last column)
        tr.onclick = (e) => {
          if (e.target.closest('button') || e.target.closest('[data-thumb]')) return;
          openDetailModal(task.Job_ID);
        };

        const isJsaOk = (task.JSA_Completed || "").toLowerCase() === 'yes';
        const photos = getTaskAttachments(task, 'Site_Photos');
        const docs = getTaskAttachments(task, 'Document_Files');
        const videos = getTaskAttachments(task, 'Video_Files');
        const deliveryFiles = getTaskDeliveryFiles(task);

        const isCompleted = task.Status === 'ส่งมอบงานแล้ว' || task.Status === 'ปิดงาน' || task.Status === 'ส่งมอบ / เอกสารเรียบร้อย';
        const daysRem = getDaysUntilExpiry(task.Contract_Expiry_Date);

        let contractBadge = '<span class="text-xs text-slate-400">-</span>';
        if (task.Contract_Expiry_Date) {
          const displayExp = formatDisplayDate(task.Contract_Expiry_Date);
          if (isCompleted) {
            contractBadge = `<span class="inline-block px-2.5 py-0.5 rounded-lg text-xs font-bold bg-slate-100 text-slate-700">${esc(displayExp)}</span>`;
          } else if (daysRem !== null && daysRem < 0) {
            contractBadge = `<span class="inline-block px-2.5 py-0.5 rounded-lg text-xs font-bold bg-rose-100 text-pts-900 border border-rose-300 animate-pulse" title="ครบกำหนด: ${esc(displayExp)}">🚨 เลย ${Math.abs(daysRem)} วัน</span>`;
          } else if (daysRem !== null && daysRem <= 7) {
            contractBadge = `<span class="inline-block px-2.5 py-0.5 rounded-lg text-xs font-bold bg-amber-100 text-amber-950 border border-amber-300 animate-pulse" title="ครบกำหนด: ${esc(displayExp)}">⚠️ ครบสัญญาใน ${daysRem} วัน</span>`;
          } else {
            contractBadge = `<span class="inline-block px-2.5 py-0.5 rounded-lg text-xs font-medium bg-slate-100 text-slate-700">${esc(displayExp)}</span>`;
          }
        }

        tr.innerHTML = `
          <td class="py-3 px-2 sm:px-4 font-mono text-xs sm:text-sm font-bold text-pts-800 whitespace-nowrap align-top">
            <span class="sm:hidden">${esc(String(task.Job_ID).replace('PTS-ENG-', '#'))}</span>
            <span class="hidden sm:inline">${esc(task.Job_ID)}</span>
          </td>
          <td class="py-3 px-2 sm:px-4 font-medium text-slate-900 align-top">
            <div class="cell-clamp sm:max-w-xs">
              <div class="font-bold text-slate-900">${esc(task.Project_Name)}</div>
              ${task.Site_Location ? `<p class="text-xs text-slate-500 mt-0.5"><i class="fa-solid fa-location-dot text-rose-500 mr-1"></i>${esc(task.Site_Location)}</p>` : ''}
              ${task.Site_Contact_Phone ? `<p class="text-xs text-slate-500 mt-0.5"><i class="fa-solid fa-phone text-emerald-600 mr-1"></i>${esc(task.Site_Contact_Phone)}</p>` : ''}
              <div class="sm:hidden mt-1.5 flex flex-wrap gap-1">
                <span class="inline-block px-2 py-0.5 rounded-lg text-[10px] font-bold leading-tight ${getStatusBadgeStyle(task.Status)}">${esc(task.Status)}</span>
                ${isCompleted && task.Completion_Date ? `<span class="inline-block px-2 py-0.5 rounded-lg text-[10px] font-bold bg-emerald-100 text-emerald-900">เสร็จ: ${esc(formatDisplayDate(task.Completion_Date))}</span>` : (task.Contract_Expiry_Date ? contractBadge : '')}
              </div>
            </div>
          </td>
          <td class="hidden md:table-cell py-3 px-4 text-xs font-semibold text-slate-600 align-top">${esc(task.Sub_Department || '-')}</td>
          <td class="py-3 px-2 sm:px-4 text-xs font-medium text-slate-700 align-top">
            <i class="fa-solid fa-user-gear text-slate-400 mr-1"></i>${esc(task.Technician_In_Charge || '-')}
          </td>
          <td class="hidden sm:table-cell py-3 px-4 text-xs font-mono text-slate-600 align-top">${esc(formatDisplayDate(task.PO_Approval_Date))}</td>
          <td class="py-3 px-2 sm:px-4 align-top">
            ${isCompleted ? `
              <span class="inline-flex items-center px-2.5 py-1 rounded-xl text-xs font-bold font-mono bg-emerald-50 text-emerald-950 border border-emerald-300">
                <i class="fa-solid fa-calendar-check text-emerald-600 mr-1"></i>${esc(formatDisplayDate(task.Completion_Date) || 'เสร็จสิ้น')}
              </span>
            ` : contractBadge}
          </td>
          <td class="py-3 px-2 sm:px-4 align-top">
            <div class="flex items-center gap-1.5">
              ${photos.length > 0 ? `
                <img data-thumb data-idx="0" src="${escUrl(photoSrc(photos[0]))}" loading="lazy"
                     class="w-9 h-9 rounded-lg object-cover border border-slate-200 bg-slate-100 cursor-zoom-in hover:scale-110 hover:border-pts-500 transition"
                     title="คลิกเพื่อดูรูปเต็ม" alt="thumbnail">
              ` : ''}
              <div class="flex flex-wrap items-center gap-1 text-xs font-bold text-slate-600">
                ${photos.length > 0 ? `<span class="px-1.5 py-0.5 rounded bg-slate-100 text-pts-800"><i class="fa-solid fa-camera mr-0.5"></i>${photos.length}</span>` : ''}
                ${docs.length > 0 ? `<span class="px-1.5 py-0.5 rounded bg-blue-50 text-blue-700"><i class="fa-solid fa-file-pdf mr-0.5"></i>${docs.length}</span>` : ''}
                ${deliveryFiles.length > 0 ? `<button type="button" onclick="event.stopPropagation(); openDeliveryDocModal('${esc(task.Job_ID)}')" class="px-1.5 py-0.5 rounded bg-blue-100 text-blue-900 font-bold border border-blue-300 hover:bg-blue-200 transition cursor-pointer" title="คลิกดูใบส่งมอบงาน (${deliveryFiles.length} ไฟล์)"><i class="fa-solid fa-file-signature mr-0.5 text-blue-700"></i>${deliveryFiles.length}</button>` : ''}
                ${videos.length > 0 ? `<span class="px-1.5 py-0.5 rounded bg-purple-50 text-purple-700"><i class="fa-solid fa-video mr-0.5"></i>${videos.length}</span>` : ''}
                ${photos.length === 0 && docs.length === 0 && deliveryFiles.length === 0 && videos.length === 0 ? '<span class="text-slate-400 font-normal">-</span>' : ''}
              </div>
            </div>
          </td>
          <td class="py-3 px-2 sm:px-4 align-top">${jobTaskCellHtml(task.Job_ID)}</td>
          <td class="hidden sm:table-cell py-3 px-4 align-top">
            <span class="inline-block px-2.5 py-1 rounded-xl text-xs font-bold ${getStatusBadgeStyle(task.Status)}">
              ${esc(task.Status)}
            </span>
          </td>
          <td class="py-3 px-2 sm:px-4 text-center align-top">
            <div class="flex items-center justify-center gap-1.5">
              <button type="button" data-action="preview" title="ดูรายละเอียด / พรีวิวรูป" class="w-9 h-9 sm:w-8 sm:h-8 rounded-lg bg-slate-100 hover:bg-slate-200 active:bg-slate-300 text-slate-600 transition">
                <i class="fa-regular fa-eye"></i>
              </button>
              <button type="button" data-action="edit" title="แก้ไขใบงาน" class="w-9 h-9 sm:w-8 sm:h-8 rounded-lg bg-pts-800 hover:bg-pts-700 active:bg-pts-900 text-white transition">
                <i class="fa-solid fa-pen-to-square"></i>
              </button>
            </div>
          </td>
        `;

        tr.querySelectorAll('[data-thumb]').forEach(el => {
          el.addEventListener('click', (e) => {
            e.stopPropagation();
            openPhotoLightbox(photos, parseInt(el.dataset.idx, 10) || 0, task);
          });
        });
        tr.querySelector('[data-action="preview"]').onclick = (e) => { e.stopPropagation(); openDetailModal(task.Job_ID); };
        tr.querySelector('[data-action="edit"]').onclick = (e) => { e.stopPropagation(); openEditModal(task.Job_ID); };
        const taskCell = tr.querySelector('[data-job-tasks]');
        // A job with tasks opens the filtered list; one without goes straight
        // to a new task already tied to it.
        if (taskCell) taskCell.onclick = (e) => {
          e.stopPropagation();
          if (jobOpenTasks(task.Job_ID).length > 0) openTasksForJobModal(task.Job_ID);
          else addTaskForJob(task.Job_ID);
        };

        tbody.appendChild(tr);
      });
    }

    function getStatusBadgeStyle(status) {
      const norm = normalizeStatus(status);
      switch(norm) {
        case 'วางแผน / เตรียมอุปกรณ์':
          return 'bg-slate-100 text-slate-800 border border-slate-300';
        case 'แจ้งเข้าทำงาน':
          return 'bg-cyan-100 text-cyan-900 border border-cyan-300 font-semibold';
        case 'เข้าปฏิบัติงานหน้างาน':
          return 'bg-amber-100 text-amber-900 border border-amber-300 font-semibold';
        case 'รอ Testing & Commissioning':
          return 'bg-purple-100 text-purple-900 border border-purple-300 font-semibold';
        case 'ติดปัญหาหน้างาน / รออะไหล่':
          return 'bg-rose-100 text-rose-900 border border-rose-300 font-bold';
        case 'ค้างทำรายงาน':
          return 'bg-indigo-100 text-indigo-900 border border-indigo-300 font-bold';
        case 'ส่งมอบ / เอกสารเรียบร้อย':
          return 'bg-teal-100 text-teal-900 border border-teal-300 font-bold';
        case 'ปิดงาน':
          return 'bg-emerald-100 text-emerald-950 border border-emerald-300 font-bold';
        default:
          return 'bg-slate-100 text-slate-700 border border-slate-200';
      }
    }

    // MULTI-FORMAT FILE ATTACHMENT HANDLER (PHOTOS, DOCS, VIDEOS)
    async function handleFileSelect(event, mode, fileCategory) {
      const input = event.target;
      const files = Array.from(input.files || []);
      input.value = '';                       // read the list first, then reset the picker
      if (files.length === 0) return;

      const tooBig = files.filter(f => f.size > MAX_ATTACHMENT_MB * 1024 * 1024);
      const usable = files.filter(f => f.size <= MAX_ATTACHMENT_MB * 1024 * 1024);
      if (tooBig.length > 0) {
        const fileNames = tooBig.map(f => `• ${f.name} (${formatBytes(f.size)})`).join('\n');
        if (fileCategory === 'video') {
          const askDrive = confirm(`ไฟล์วิดีโอต่อไปนี้มีขนาดเกินขีดจำกัดการส่งตรง (${MAX_ATTACHMENT_MB} MB):\n\n${fileNames}\n\n💡 แนะนำ: สำหรับคลิปขนาดใหญ่ (50MB - 1GB+) ให้อัปโหลดเข้า Google Drive แล้วนำลิงก์มาวางในระบบเพื่อเปิดดูได้ไม่จำกัดขนาด\n\nต้องการวางลิงก์ Google Drive ของไฟล์นี้ตอนนี้เลยหรือไม่?`);
          if (askDrive) {
            promptAddDriveVideo(mode);
          }
        } else {
          alert(`ไฟล์ต่อไปนี้มีขนาดเกิน ${MAX_ATTACHMENT_MB} MB:\n\n${fileNames}\n\n📌 ระบบรองรับการอัปโหลดไฟล์ผ่านแอปได้สูงสุดถึง ${MAX_ATTACHMENT_MB} MB ต่อไฟล์\nสำหรับไฟล์ขนาดใหญ่กว่านี้ แนะนำอัปโหลดขึ้น Google Drive แล้วนำลิงก์มาวางในช่อง "รายละเอียดงาน / หมายเหตุ" ได้ครับ`);
        }
      }
      if (usable.length === 0) return;

      if (fileCategory === 'photo') {
        await addPhotosToStaging(mode, usable);
        return;
      }

      if (fileCategory === 'delivery') {
        const results = await Promise.all(usable.map(file => readFileAsDataUrl(file).then(dataUrl => (
          dataUrl ? {
            name: file.name,
            size: formatBytes(file.size),
            type: file.type || getFileExt(file.name),
            dataUrl: dataUrl
          } : null
        ))));

        const target = mode === 'new' ? (state.tempNewDelivery = state.tempNewDelivery || []) : (state.tempEditDelivery = state.tempEditDelivery || []);
        results.forEach(item => { if (item) target.push(item); });
        renderDeliveryPreview(mode);

        const failed = results.filter(r => !r).length;
        if (failed > 0) alert(`อ่านไฟล์ไม่สำเร็จ ${failed} ไฟล์`);
        return;
      }

      const isVideo = fileCategory === 'video';
      const results = await Promise.all(usable.map(file => readFileAsDataUrl(file).then(dataUrl => (
        dataUrl ? {
          name: file.name,
          size: formatBytes(file.size),
          type: file.type || (isVideo ? 'video/mp4' : getFileExt(file.name)),
          dataUrl: dataUrl
        } : null
      ))));

      const target = isVideo
        ? (mode === 'new' ? state.tempNewVideos : state.tempEditVideos)
        : (mode === 'new' ? state.tempNewDocs : state.tempEditDocs);
      results.forEach(item => { if (item) target.push(item); });

      if (isVideo) renderVideosPreview(mode);
      else renderDocsPreview(mode);

      const failed = results.filter(r => !r).length;
      if (failed > 0) alert(`อ่านไฟล์ไม่สำเร็จ ${failed} ไฟล์`);
    }

    /* -----------------------------------------------------------------------
     * BULK PHOTO STAGING
     *
     * Photos are never uploaded one by one: every picker / drop / paste adds to
     * the same staging list, and ONE press of Save sends the whole list.
     *
     * Decoding is done a few files at a time. Decoding 30 full-size phone photos
     * at once (the old Promise.all) could exhaust memory on a phone - Safari
     * then reloads the page and every staged photo is gone, which is exactly
     * what pushed users into adding a single photo per save. Small chunks keep
     * memory flat, preserve the order the user picked, and let the gallery fill
     * in as it goes.
     * --------------------------------------------------------------------- */
    const PHOTO_PREP_CHUNK = 3;
    const photoPrepInProgress = { new: 0, edit: 0 };

    function photoPrepBusy(mode) {
      return (photoPrepInProgress[mode] || 0) > 0;
    }

    function setPhotoPrepStatus(mode, text) {
      const el = document.getElementById(mode === 'new' ? 'newPhotoPrepStatus' : 'editPhotoPrepStatus');
      if (!el) return;
      if (!text) {
        el.classList.add('hidden');
        el.textContent = '';
        return;
      }
      el.innerHTML = `<i class="fa-solid fa-circle-notch fa-spin mr-1"></i>${esc(text)}`;
      el.classList.remove('hidden');
    }

    // The save button is locked while photos are still decoding - saving early
    // would silently drop everything not yet in the list.
    function lockSaveWhilePreparing(mode, locked) {
      const btn = document.getElementById(mode === 'new' ? 'createSubmitBtn' : 'editSaveBtn');
      if (!btn) return;
      if (locked) {
        if (!btn.dataset.prepLabel) btn.dataset.prepLabel = btn.innerHTML;
        btn.disabled = true;
        btn.classList.add('opacity-60', 'cursor-wait');
        btn.innerHTML = '<i class="fa-solid fa-circle-notch fa-spin mr-1.5"></i>กำลังเตรียมรูป...';
      } else {
        btn.disabled = false;
        btn.classList.remove('opacity-60', 'cursor-wait');
        if (btn.dataset.prepLabel) btn.innerHTML = btn.dataset.prepLabel;
        delete btn.dataset.prepLabel;
      }
    }

    async function addPhotosToStaging(mode, files) {
      const images = (files || []).filter(f => f && (/^image\//i.test(f.type || '') || /\.(jpe?g|png|gif|webp|bmp|heic|heif)$/i.test(f.name || '')));
      if (images.length === 0) return;

      const target = mode === 'new' ? state.tempNewPhotos : state.tempEditPhotos;
      const availableSlots = MAX_PHOTOS_PER_TASK - target.length - (photoPrepInProgress[mode] || 0);

      if (availableSlots <= 0) {
        alert(`ใบงานนี้แนบรูปภาพครบโควต้าสูงสุดแล้ว (${MAX_PHOTOS_PER_TASK} รูป)`);
        return;
      }

      let queue = images;
      if (queue.length > availableSlots) {
        alert(`คุณเลือกรูปภาพเกินโควต้า ระบบจะเพิ่มรูปให้เพียง ${availableSlots} รูปแรกเท่านั้น`);
        queue = queue.slice(0, availableSlots);
      }

      const total = queue.length;
      photoPrepInProgress[mode] = (photoPrepInProgress[mode] || 0) + total;
      lockSaveWhilePreparing(mode, true);

      let done = 0;
      let failed = 0;
      try {
        for (let i = 0; i < queue.length; i += PHOTO_PREP_CHUNK) {
          const chunk = queue.slice(i, i + PHOTO_PREP_CHUNK);
          setPhotoPrepStatus(mode, `กำลังเตรียมรูป ${Math.min(done + chunk.length, total)} / ${total} รูป... (ยังเพิ่มรูปต่อได้ แล้วค่อยกดบันทึกครั้งเดียว)`);

          // Promise.all inside a chunk keeps the picked order; chunks are sequential.
          const results = await Promise.all(chunk.map(f => compressImage(f, PHOTO_MAX_WIDTH, PHOTO_QUALITY)));
          results.forEach(dataUrl => {
            if (dataUrl) target.push(dataUrl);
            else failed++;
          });
          done += chunk.length;
          photoPrepInProgress[mode] -= chunk.length;
          renderPhotosPreview(mode);
        }
      } finally {
        photoPrepInProgress[mode] = Math.max(0, photoPrepInProgress[mode] - (total - done));
        if (!photoPrepBusy(mode)) {
          setPhotoPrepStatus(mode, '');
          lockSaveWhilePreparing(mode, false);
        }
      }

      if (failed > 0) {
        alert(`อ่านไฟล์รูปภาพไม่สำเร็จ ${failed} ไฟล์ (ไฟล์อาจเสียหายหรือเป็นชนิดที่เบราว์เซอร์ไม่รองรับ เช่น HEIC)\n\nเพิ่มสำเร็จ ${total - failed} รูป`);
      }
    }

    // Which photo staging list a drop / paste should land in
    function activePhotoMode() {
      const createOpen = document.getElementById('createModal') && !document.getElementById('createModal').classList.contains('hidden');
      if (createOpen) return 'new';
      const editOpen = document.getElementById('editModal') && !document.getElementById('editModal').classList.contains('hidden');
      if (editOpen) return 'edit';
      return null;
    }

    function initPhotoDropZones() {
      document.querySelectorAll('[data-photo-drop]').forEach(zone => {
        const mode = zone.dataset.photoDrop;
        const highlight = on => {
          zone.classList.toggle('ring-2', on);
          zone.classList.toggle('ring-pts-800', on);
          zone.classList.toggle('bg-pts-50', on);
        };
        zone.addEventListener('dragover', e => { e.preventDefault(); highlight(true); });
        zone.addEventListener('dragleave', () => highlight(false));
        zone.addEventListener('drop', e => {
          e.preventDefault();
          highlight(false);
          const files = Array.from((e.dataTransfer && e.dataTransfer.files) || []);
          if (files.length) addPhotosToStaging(mode, files);
        });
      });

      // Ctrl+V a screenshot or a copied image straight into the open dialog
      document.addEventListener('paste', e => {
        const mode = activePhotoMode();
        if (!mode) return;
        const items = Array.from((e.clipboardData && e.clipboardData.items) || []);
        const files = items.filter(it => it.kind === 'file' && /^image\//i.test(it.type)).map(it => it.getAsFile()).filter(Boolean);
        if (files.length === 0) return;
        e.preventDefault();
        addPhotosToStaging(mode, files);
      });
    }

    function readFileAsDataUrl(file) {
      return new Promise(resolve => {
        const reader = new FileReader();
        reader.onload = e => resolve(e.target.result);
        reader.onerror = () => resolve(null);
        reader.readAsDataURL(file);
      });
    }

    function formatBytes(bytes) {
      if (!bytes) return '0 B';
      const k = 1024;
      const sizes = ['B', 'KB', 'MB', 'GB'];
      const i = Math.floor(Math.log(bytes) / Math.log(k));
      return parseFloat((bytes / Math.pow(k, i)).toFixed(1)) + ' ' + sizes[i];
    }

    function getFileExt(fileName) {
      return String(fileName || '').split('.').pop().toLowerCase();
    }

    // Resolves to a JPEG data URL, or null if the browser cannot decode the file.
    // The old version had no error path at all, so an unreadable photo simply
    // vanished with no message.
    function compressImage(file, maxWidth, quality) {
      return new Promise(resolve => {
        const reader = new FileReader();
        reader.onerror = () => resolve(null);
        reader.onload = function (e) {
          const img = new Image();
          img.onerror = () => resolve(null);
          img.onload = function () {
            try {
              let width = img.naturalWidth || img.width;
              let height = img.naturalHeight || img.height;
              if (!width || !height) return resolve(null);

              if (width > maxWidth) {
                height = Math.round((height * maxWidth) / width);
                width = maxWidth;
              }

              const canvas = document.createElement('canvas');
              canvas.width = width;
              canvas.height = height;
              const ctx = canvas.getContext('2d');
              ctx.drawImage(img, 0, 0, width, height);
              resolve(canvas.toDataURL('image/jpeg', quality));
            } catch (err) {
              resolve(null);
            }
          };
          img.src = e.target.result;
        };
        reader.readAsDataURL(file);
      });
    }

    // RENDER PHOTOS PREVIEW
    function renderPhotosPreview(mode) {
      const isNew = mode === 'new';
      const container = document.getElementById(isNew ? 'newPhotosGalleryContainer' : 'editPhotosGalleryContainer');
      const badge = document.getElementById(isNew ? 'newPhotoCountBadge' : 'editPhotoCountBadge');
      if (!container || !badge) return;

      const photos = isNew ? state.tempNewPhotos : state.tempEditPhotos;
      container.innerHTML = '';
      const count = photos.length;
      badge.textContent = `(${count} / ${MAX_PHOTOS_PER_TASK} รูป)`;

      if (count === 0) {
        container.innerHTML = `<p class="text-xs text-slate-400 p-1 italic">ยังไม่มีรูปภาพแนบ</p>`;
        return;
      }

      photos.forEach((imgUrl, idx) => {
        const div = document.createElement('div');
        div.className = "relative group w-16 h-16 rounded-xl overflow-hidden border border-slate-300 shadow-sm cursor-pointer shrink-0";
        if (isNew) {
          div.onclick = () => openPhotoLightbox(state.tempNewPhotos, idx, { Job_ID: 'NEW', Project_Name: 'ใบงานใหม่' });
        } else {
          div.onclick = () => openPhotoLightbox(state.tempEditPhotos, idx, state.currentEditingTask);
        }
        div.innerHTML = `
          <img src="${escUrl(photoSrc(imgUrl))}" class="w-full h-full object-cover hover:scale-110 transition duration-200" alt="Site Photo">
          <button type="button" onclick="event.stopPropagation(); removeAttachment('${mode}', 'photo', ${idx})" class="absolute top-0.5 right-0.5 bg-rose-600 text-white rounded-full w-6 h-6 sm:w-4 sm:h-4 text-[11px] sm:text-[9px] flex items-center justify-center shadow hover:bg-rose-700 active:scale-90 transition">
            <i class="fa-solid fa-xmark"></i>
          </button>
        `;
        container.appendChild(div);
      });
    }

    // RENDER DOCUMENTS PREVIEW (PDF, WORD, EXCEL)
    function renderDocsPreview(mode) {
      const isNew = mode === 'new';
      const docs = isNew ? state.tempNewDocs : state.tempEditDocs;
      const container = document.getElementById(isNew ? 'newDocsGalleryContainer' : 'editDocsGalleryContainer');
      const label = document.getElementById(isNew ? 'newDocCountBadge' : 'editDocCountBadge');
      
      if (!container || !label) return;
      container.innerHTML = '';
      label.textContent = `(${docs.length} ไฟล์)`;

      if (docs.length === 0) {
        container.innerHTML = `<p class="text-xs text-slate-400 p-1 italic">ยังไม่มีไฟล์เอกสารแนบ</p>`;
        return;
      }

      docs.forEach((doc, idx) => {
        const ext = getFileExt(doc.name || '');
        let iconClass = "fa-file-lines text-slate-600";
        let bgClass = "bg-slate-100 border-slate-300";

        if (ext.includes('pdf')) {
          iconClass = "fa-file-pdf text-rose-600";
          bgClass = "bg-rose-50 border-rose-200";
        } else if (ext.includes('doc')) {
          iconClass = "fa-file-word text-blue-600";
          bgClass = "bg-blue-50 border-blue-200";
        } else if (ext.includes('xls')) {
          iconClass = "fa-file-excel text-emerald-600";
          bgClass = "bg-emerald-50 border-emerald-200";
        }

        const div = document.createElement('div');
        div.className = `flex items-center justify-between p-2 rounded-xl border ${bgClass} shadow-sm text-xs space-x-2 max-w-xs`;
        div.innerHTML = `
          <div class="flex items-center space-x-2 truncate">
            <i class="fa-solid ${iconClass} text-base shrink-0"></i>
            <div class="truncate">
              <p class="font-bold text-slate-800 truncate" title="${esc(doc.name)}">${esc(doc.name)}</p>
              <p class="text-[10px] text-slate-500">${esc(doc.size || '')}</p>
            </div>
          </div>
          <div class="flex items-center space-x-1 shrink-0">
            <a href="${escUrl(fileSrc(doc)) || '#'}" ${fileSrc(doc).indexOf('data:') === 0 ? `download="${esc(doc.name)}"` : 'rel="noopener"'} target="_blank" class="p-1 text-slate-600 hover:text-slate-900 transition" title="เปิด / ดาวน์โหลดไฟล์">
              <i class="fa-solid ${fileSrc(doc).indexOf('data:') === 0 ? 'fa-download' : 'fa-up-right-from-square'}"></i>
            </a>
            <button type="button" onclick="removeAttachment('${mode}', 'doc', ${idx})" class="p-1 text-rose-600 hover:text-rose-800 transition" title="ลบไฟล์">
              <i class="fa-solid fa-xmark"></i>
            </button>
          </div>
        `;
        container.appendChild(div);
      });
    }

    // RENDER VIDEOS PREVIEW (MP4, WEBM, MOV)
    function renderVideosPreview(mode) {
      const isNew = mode === 'new';
      const videos = isNew ? state.tempNewVideos : state.tempEditVideos;
      const container = document.getElementById(isNew ? 'newVideosGalleryContainer' : 'editVideosGalleryContainer');
      const label = document.getElementById(isNew ? 'newVideoCountBadge' : 'editVideoCountBadge');

      if (!container || !label) return;
      container.innerHTML = '';
      label.textContent = `(${videos.length} วิดีโอ)`;

      if (videos.length === 0) {
        container.innerHTML = `<p class="text-xs text-slate-400 p-1 italic">ยังไม่มีวิดีโอแนบ</p>`;
        return;
      }

      videos.forEach((vid, idx) => {
        const div = document.createElement('div');
        div.className = "relative group w-24 h-16 rounded-xl overflow-hidden border border-purple-300 bg-purple-950 text-white flex items-center justify-center cursor-pointer shrink-0 shadow-sm";
        div.onclick = () => openVideoPreview(vid);
        const poster = videoThumb(vid);
        div.innerHTML = `
          ${poster ? `<img src="${escUrl(poster)}" loading="lazy" alt="" class="absolute inset-0 w-full h-full object-cover opacity-70" onerror="this.remove()">` : ''}
          <div class="relative text-center p-1">
            <i class="fa-solid fa-circle-play text-xl text-white drop-shadow group-hover:scale-125 transition"></i>
            <p class="text-[9px] truncate max-w-[80px] mt-0.5 text-purple-100">${esc(vid.name)}</p>
          </div>
          <button type="button" onclick="event.stopPropagation(); removeAttachment('${mode}', 'video', ${idx})" class="absolute top-0.5 right-0.5 bg-rose-600 text-white rounded-full w-6 h-6 sm:w-4 sm:h-4 text-[11px] sm:text-[9px] flex items-center justify-center shadow hover:bg-rose-700 active:scale-90 transition z-10">
            <i class="fa-solid fa-xmark"></i>
          </button>
        `;
        container.appendChild(div);
      });
    }

    function removeAttachment(mode, category, index) {
      if (mode === 'new') {
        if (category === 'photo') { state.tempNewPhotos.splice(index, 1); renderPhotosPreview('new'); }
        if (category === 'doc') { state.tempNewDocs.splice(index, 1); renderDocsPreview('new'); }
        if (category === 'delivery') { (state.tempNewDelivery || []).splice(index, 1); renderDeliveryPreview('new'); }
        if (category === 'video') { state.tempNewVideos.splice(index, 1); renderVideosPreview('new'); }
      } else if (mode === 'edit') {
        if (category === 'photo') { state.tempEditPhotos.splice(index, 1); renderPhotosPreview('edit'); }
        if (category === 'doc') { state.tempEditDocs.splice(index, 1); renderDocsPreview('edit'); }
        if (category === 'delivery') { (state.tempEditDelivery || []).splice(index, 1); renderDeliveryPreview('edit'); }
        if (category === 'video') { state.tempEditVideos.splice(index, 1); renderVideosPreview('edit'); }
      }
    }

    // RENDER DELIVERY DOCS PREVIEW (FOR CREATE & EDIT MODALS)
    function renderDeliveryPreview(mode) {
      const isNew = mode === 'new';
      const docs = isNew ? (state.tempNewDelivery || []) : (state.tempEditDelivery || []);
      const container = document.getElementById(isNew ? 'newDeliveryGalleryContainer' : 'editDeliveryGalleryContainer');
      const label = document.getElementById(isNew ? 'newDeliveryCountBadge' : 'editDeliveryCountBadge');
      
      if (!container || !label) return;
      container.innerHTML = '';
      label.textContent = `(${docs.length} ไฟล์)`;

      if (docs.length === 0) {
        container.innerHTML = `<p class="text-xs text-slate-400 p-1 italic">ยังไม่มีไฟล์ใบส่งมอบงานแนบ</p>`;
        return;
      }

      docs.forEach((doc, idx) => {
        const ext = getFileExt(doc.name || '');
        let iconClass = "fa-file-lines text-slate-600";
        let bgClass = "bg-white border-emerald-200";

        if (ext.includes('pdf')) {
          iconClass = "fa-file-pdf text-rose-600";
          bgClass = "bg-rose-50 border-rose-200";
        } else if (ext.includes('doc')) {
          iconClass = "fa-file-word text-blue-600";
          bgClass = "bg-blue-50 border-blue-200";
        } else if (ext.includes('jpg') || ext.includes('jpeg') || ext.includes('png')) {
          iconClass = "fa-file-image text-emerald-600";
          bgClass = "bg-emerald-50 border-emerald-200";
        }

        const div = document.createElement('div');
        div.className = `flex items-center justify-between p-2 rounded-xl border ${bgClass} shadow-sm text-xs space-x-2 max-w-xs`;
        div.innerHTML = `
          <div class="flex items-center space-x-2 truncate">
            <i class="fa-solid ${iconClass} text-base shrink-0"></i>
            <div class="truncate">
              <p class="font-bold text-slate-800 truncate" title="${esc(doc.name)}">${esc(doc.name)}</p>
              <p class="text-[10px] text-slate-500">${esc(doc.size || '')}</p>
            </div>
          </div>
          <div class="flex items-center space-x-1 shrink-0">
            <a href="${escUrl(fileSrc(doc)) || '#'}" ${fileSrc(doc).indexOf('data:') === 0 ? `download="${esc(doc.name)}"` : 'rel="noopener"'} target="_blank" class="p-1 text-emerald-700 hover:text-emerald-900 transition" title="เปิด / ดาวน์โหลด">
              <i class="fa-solid ${fileSrc(doc).indexOf('data:') === 0 ? 'fa-download' : 'fa-up-right-from-square'}"></i>
            </a>
            <button type="button" onclick="removeAttachment('${mode}', 'delivery', ${idx})" class="p-1 text-rose-600 hover:text-rose-800 transition" title="ลบไฟล์">
              <i class="fa-solid fa-xmark"></i>
            </button>
          </div>
        `;
        container.appendChild(div);
      });
    }

    // LIGHTBOX PREVIEW (IMAGE & VIDEO)
    function openLightboxModal(type, src, caption, openUrl) {
      const img = document.getElementById('lightboxImg');
      const video = document.getElementById('lightboxVideo');
      const frame = document.getElementById('lightboxFrame');
      const cap = document.getElementById('lightboxCaption');
      const link = document.getElementById('lightboxOpenLink');

      cap.textContent = caption || 'Media Preview';

      // Reset every player before showing the requested one
      video.pause();
      video.classList.add('hidden');
      img.classList.add('hidden');
      frame.classList.add('hidden');
      frame.src = 'about:blank';
      const dlBtn = document.getElementById('lightboxDownloadBtn');
      if (dlBtn) dlBtn.classList.add('hidden');   // re-shown by renderLightboxPhoto

      if (type === 'image') {
        img.onerror = () => {
          img.onerror = null;
          cap.textContent = (caption || '') + ' — โหลดรูปไม่สำเร็จ (ตรวจสอบสิทธิ์แชร์ไฟล์ใน Google Drive)';
        };
        img.src = src;
        img.classList.remove('hidden');
      } else if (type === 'video') {
        video.src = src;
        video.classList.remove('hidden');
      } else if (type === 'drive') {
        frame.src = src;
        frame.classList.remove('hidden');
      }

      const fallbackUrl = openUrl || driveOpenUrl(src);
      if (fallbackUrl) {
        link.href = fallbackUrl;
        link.classList.remove('hidden');
      } else {
        link.classList.add('hidden');
      }

      document.getElementById('lightboxModal').classList.remove('hidden');
    }

    // GALLERY LIGHTBOX: preview any photo straight from a card/row/detail modal
    function openPhotoLightbox(photos, startIndex, task) {
      const list = (photos || []).map(photoSrc).filter(Boolean);
      if (list.length === 0) return;

      state.lightboxPhotos = list;
      state.lightboxIndex = Math.min(Math.max(startIndex || 0, 0), list.length - 1);
      state.lightboxLabel = task ? `${task.Job_ID} · ${task.Project_Name || ''}` : '';
      renderLightboxPhoto();
    }

    function renderLightboxPhoto() {
      const list = state.lightboxPhotos;
      const idx = state.lightboxIndex;
      const showNav = list.length > 1;

      document.getElementById('lightboxPrevBtn').classList.toggle('hidden', !showNav);
      document.getElementById('lightboxNextBtn').classList.toggle('hidden', !showNav);

      const caption = `${state.lightboxLabel ? state.lightboxLabel + ' · ' : ''}รูปที่ ${idx + 1} / ${list.length}`;
      openLightboxModal('image', list[idx], caption);

      // Only a photo gallery can be saved this way - videos and PDFs keep the
      // "open in Drive" link instead.
      const dl = document.getElementById('lightboxDownloadBtn');
      if (dl) dl.classList.remove('hidden');
    }

    function stepLightbox(delta) {
      const list = state.lightboxPhotos;
      if (!list || list.length === 0) return;
      state.lightboxIndex = (state.lightboxIndex + delta + list.length) % list.length;
      renderLightboxPhoto();
    }

    function closeLightboxModal() {
      const video = document.getElementById('lightboxVideo');
      if (video) { video.pause(); video.removeAttribute('src'); }
      const frame = document.getElementById('lightboxFrame');
      if (frame) { frame.src = 'about:blank'; frame.classList.add('hidden'); }
      state.lightboxPhotos = [];
      document.getElementById('lightboxPrevBtn').classList.add('hidden');
      document.getElementById('lightboxNextBtn').classList.add('hidden');
      document.getElementById('lightboxModal').classList.add('hidden');
    }

    // KEYBOARD NAVIGATION FOR THE PREVIEW LIGHTBOX
    document.addEventListener('keydown', (e) => {
      const lightboxOpen = !document.getElementById('lightboxModal').classList.contains('hidden');
      if (lightboxOpen) {
        if (e.key === 'Escape') closeLightboxModal();
        if (e.key === 'ArrowLeft') stepLightbox(-1);
        if (e.key === 'ArrowRight') stepLightbox(1);
        return;
      }
      if (e.key === 'Escape' && !document.getElementById('detailModal').classList.contains('hidden')) {
        closeDetailModal();
      }
    });

    /* ---------------------------------------------------------------------
     * MODAL 2.5: READ-ONLY DETAIL VIEW (PREVIEW WITHOUT ENTERING EDIT MODE)
     * ------------------------------------------------------------------- */
    function openDetailModal(jobId) {
      const task = state.tasks.find(t => t.Job_ID === jobId);
      if (!task) return;

      state.currentDetailJobId = jobId;

      const photos = getTaskAttachments(task, 'Site_Photos');
      const docs = getTaskAttachments(task, 'Document_Files');
      const videos = getTaskAttachments(task, 'Video_Files');

      document.getElementById('detailJobIdBadge').textContent = task.Job_ID;
      document.getElementById('detailProjectTitle').textContent = task.Project_Name || '-';
      document.getElementById('detailSubtitle').textContent =
        `${task.Sub_Department || '-'} · ${task.Technician_In_Charge || 'ไม่ระบุผู้รับผิดชอบ'}`;
      document.getElementById('detailUpdatedAt').textContent = `อัปเดตล่าสุด: ${task.Updated_At || '-'}`;
      document.getElementById('detailNotesDisplay').textContent = task.Notes_Issues || 'ยังไม่มีบันทึกเพิ่มเติม';

      renderDetailSummary(task);
      renderDetailInfoGrid(task);
      renderDetailTasks(task.Job_ID);
      renderDetailPhotos(task, photos);
      renderDetailDocs(docs);
      renderDetailVideos(videos);

      const viewBtn = document.getElementById('detailViewTasksBtn');
      if (viewBtn) viewBtn.onclick = () => openTasksForJobModal(task.Job_ID);
      const addBtn = document.getElementById('detailAddTaskBtn');
      if (addBtn) addBtn.onclick = () => addTaskForJob(task.Job_ID);

      document.getElementById('detailModal').classList.remove('hidden');
    }

    function closeDetailModal() {
      state.currentDetailJobId = null;
      // Leaving select mode armed would block the next auto-refresh of this
      // dialog and carry a stale tick list into the next job opened.
      setPhotoSelectMode(false);
      document.getElementById('detailModal').classList.add('hidden');
    }

    function editFromDetail() {
      const jobId = state.currentDetailJobId;
      closeDetailModal();
      if (jobId) openEditModal(jobId);
    }

    function renderDetailSummary(task) {
      const isJsaOk = (task.JSA_Completed || '').toLowerCase() === 'yes';
      const isCompleted = task.Status === 'ส่งมอบงานแล้ว' || task.Status === 'ปิดงาน' || task.Status === 'ส่งมอบ / เอกสารเรียบร้อย';
      const daysRem = getDaysUntilExpiry(task.Contract_Expiry_Date);
      const deliveryFiles = getTaskDeliveryFiles(task);

      let contractLine = '<span class="text-slate-400">ไม่ได้ระบุวันครบกำหนดสัญญา</span>';
      if (isCompleted) {
        contractLine = `<span class="text-emerald-700 font-bold"><i class="fa-solid fa-flag-checkered mr-1"></i>ส่งมอบงานเรียบร้อย (${esc(formatDisplayDate(task.Completion_Date))})</span>`;
      } else if (daysRem !== null && daysRem < 0) {
        contractLine = `<span class="text-rose-700 font-bold"><i class="fa-solid fa-triangle-exclamation mr-1"></i>เลยกำหนดสัญญาแล้ว ${Math.abs(daysRem)} วัน (${esc(formatDisplayDate(task.Contract_Expiry_Date))})</span>`;
      } else if (daysRem !== null && daysRem <= 7) {
        contractLine = `<span class="text-amber-800 font-bold"><i class="fa-solid fa-clock mr-1"></i>ครบกำหนดสัญญาในอีก ${daysRem} วัน (${esc(formatDisplayDate(task.Contract_Expiry_Date))})</span>`;
      } else if (daysRem !== null) {
        contractLine = `<span class="text-slate-600"><i class="fa-regular fa-calendar-check mr-1"></i>ครบกำหนดสัญญา: ${esc(formatDisplayDate(task.Contract_Expiry_Date))}</span>`;
      }

      document.getElementById('detailSummary').innerHTML = `
        <div class="flex flex-wrap items-center gap-2">
          <span class="inline-block px-3 py-1 rounded-xl text-xs font-bold ${getStatusBadgeStyle(task.Status)}">
            ${esc(task.Status || '-')}
          </span>
          <span class="inline-flex items-center px-2.5 py-1 rounded-xl text-xs font-bold ${isJsaOk ? 'bg-emerald-100 text-emerald-800 border border-emerald-300' : 'bg-amber-100 text-amber-800 border border-amber-300'}">
            <i class="fa-solid ${isJsaOk ? 'fa-circle-check text-emerald-600' : 'fa-clock text-amber-600'} mr-1"></i>
            JSA ${isJsaOk ? 'Approved' : 'Pending'}
          </span>
          <span class="inline-block px-2.5 py-1 rounded-xl text-xs font-bold ${task.Priority === 'High' ? 'bg-rose-50 text-rose-700 border border-rose-200' : task.Priority === 'Medium' ? 'bg-amber-50 text-amber-700 border border-amber-200' : 'bg-slate-100 text-slate-700 border border-slate-200'}">
            Priority: ${esc(task.Priority || 'Medium')}
          </span>
          <!-- TOP ROW WORK HANDOVER BUTTON (คลิกดูเอกสารใบส่งมอบงาน) -->
          <button type="button" onclick="openDeliveryDocModal('${esc(task.Job_ID)}')" class="inline-flex items-center px-3 py-1 rounded-xl text-xs font-bold ${deliveryFiles.length > 0 ? 'bg-blue-600 hover:bg-blue-700 text-white shadow-md shadow-blue-900/20 border border-blue-500' : 'bg-white hover:bg-blue-50 text-blue-900 border border-blue-300'} transition active:scale-95 cursor-pointer">
            <i class="fa-solid ${deliveryFiles.length > 0 ? 'fa-file-circle-check text-blue-200' : 'fa-file-signature text-blue-600'} mr-1.5"></i>
            ใบส่งมอบงาน ${deliveryFiles.length > 0 ? `(${deliveryFiles.length} ไฟล์)` : ''}
          </button>
        </div>
        <div class="p-3 rounded-2xl bg-slate-50 border border-slate-200 text-xs">${contractLine}</div>
      `;
    }

    function renderDetailInfoGrid(task) {
      const cell = (label, value, icon, actionHtml = '') => `
        <div class="p-3 rounded-2xl bg-slate-50 border border-slate-200">
          <p class="text-[10px] font-bold text-slate-500 uppercase tracking-wider mb-0.5">
            <i class="fa-solid ${icon} mr-1 text-slate-400"></i>${label}
          </p>
          <div class="flex items-center justify-between gap-2">
            <p class="text-xs font-semibold text-slate-800 break-words">${esc(value || '-')}</p>
            ${actionHtml}
          </div>
        </div>`;

      const phoneAction = task.Site_Contact_Phone ? `
        <a href="tel:${esc(task.Site_Contact_Phone.replace(/[^\d+]/g, ''))}" class="px-2.5 py-1 bg-emerald-600 hover:bg-emerald-700 text-white rounded-xl text-[11px] font-bold shrink-0 flex items-center gap-1 shadow-sm transition" title="โทรออก">
          <i class="fa-solid fa-phone"></i> โทร
        </a>` : '';

      const mapUrl = task.Site_Map_Url || (task.Site_Location ? ('https://www.google.com/maps/search/?api=1&query=' + encodeURIComponent(task.Site_Location)) : '');
      const mapAction = mapUrl ? `
        <a href="${escUrl(mapUrl)}" target="_blank" rel="noopener" class="px-2.5 py-1 bg-rose-600 hover:bg-rose-700 text-white rounded-xl text-[11px] font-bold shrink-0 flex items-center gap-1 shadow-sm transition" title="เปิดดูแผนที่">
          <i class="fa-solid fa-map-location-dot"></i> ดูแผนที่
        </a>` : '';

      document.getElementById('detailInfoGrid').innerHTML = [
        cell('หัวหน้าช่างผู้รับผิดชอบ', task.Technician_In_Charge, 'fa-user-gear'),
        cell('เบอร์ติดต่อหน้างาน', task.Site_Contact_Phone, 'fa-phone-volume text-emerald-600', phoneAction),
        cell('สถานที่ / พิกัดหน้างาน', task.Site_Location, 'fa-location-dot text-rose-500', mapAction),
        cell('แผนที่โลเคชั่นงาน (Google Maps Link)', task.Site_Map_Url ? 'เปิดดูแผนที่' : (task.Site_Location ? 'ค้นหาตามพิกัด' : '-'), 'fa-map-pin text-rose-600', mapAction),
        cell('วันอนุมัติ PO', formatDisplayDate(task.PO_Approval_Date), 'fa-file-signature text-amber-600'),
        cell('วันปฏิบัติงาน (Target)', formatDisplayDate(task.Target_Date), 'fa-calendar-day'),
        cell('วันครบกำหนดสัญญา', formatDisplayDate(task.Contract_Expiry_Date), 'fa-calendar-xmark text-pts-800'),
        cell('วันที่งานเสร็จจริง', formatDisplayDate(task.Completion_Date), 'fa-flag-checkered text-emerald-700'),
        `<div class="sm:col-span-2 p-3 rounded-2xl bg-slate-50 border border-slate-200">
           <p class="text-[10px] font-bold text-slate-500 uppercase tracking-wider mb-0.5">
             <i class="fa-solid fa-clipboard-list mr-1 text-slate-400"></i>รายละเอียดงาน
           </p>
           <p class="text-xs text-slate-800 whitespace-pre-wrap">${esc(task.Task_Detail || '-')}</p>
         </div>`
      ].join('');
    }

    function renderDetailPhotos(task, photos) {
      document.getElementById('detailPhotoCount').textContent = `(${photos.length} รูป)`;

      // Remembered so the download buttons know exactly what is on screen
      state.detailPhotos = photos || [];
      state.detailPhotoTask = task;

      const tools = document.getElementById('detailPhotoTools');
      if (tools) {
        tools.classList.toggle('hidden', photos.length === 0);
        tools.classList.toggle('flex', photos.length > 0);
      }

      // Reopening the dialog always starts from a clean selection
      setPhotoSelectMode(false);
    }

    // Split out from renderDetailPhotos so toggling a selection redraws the
    // thumbnails without re-reading the task.
    function renderDetailPhotoThumbs() {
      const container = document.getElementById('detailPhotosGallery');
      if (!container) return;
      const photos = state.detailPhotos || [];
      const task = state.detailPhotoTask;
      container.innerHTML = '';

      if (photos.length === 0) {
        container.innerHTML = `<p class="col-span-full text-xs text-slate-400 italic p-1">ยังไม่มีรูปภาพแนบในใบงานนี้</p>`;
        return;
      }

      photos.forEach((p, idx) => {
        const src = photoSrc(p);
        const picked = selectedPhotoIdx.has(idx);
        const div = document.createElement('div');
        div.className = "relative aspect-square rounded-xl overflow-hidden border bg-slate-100 group transition " +
          (photoSelectMode
            ? (picked ? 'border-pts-800 ring-2 ring-pts-800 cursor-pointer' : 'border-slate-200 cursor-pointer hover:border-pts-400')
            : 'border-slate-200 cursor-zoom-in');
        div.onclick = () => {
          if (photoSelectMode) togglePhotoSelected(idx);
          else openPhotoLightbox(photos, idx, task);
        };
        div.innerHTML = `
          <img src="${escUrl(src)}" loading="lazy" alt="Site photo ${idx + 1}" class="w-full h-full object-cover ${photoSelectMode ? (picked ? 'opacity-100' : 'opacity-60') : 'group-hover:scale-110'} transition duration-200">
          ${photoSelectMode ? `
            <span class="absolute top-1 left-1 w-6 h-6 rounded-full flex items-center justify-center text-[11px] shadow ${picked ? 'bg-pts-800 text-white' : 'bg-white/90 text-slate-400 border border-slate-300'}">
              <i class="fa-solid ${picked ? 'fa-check' : 'fa-circle'} ${picked ? '' : 'text-[7px]'}"></i>
            </span>` : ''}
          <span class="absolute bottom-1 right-1 text-[9px] font-bold text-white bg-slate-900/70 px-1.5 py-0.5 rounded-md">${idx + 1}</span>
        `;
        container.appendChild(div);
      });
    }

    function renderDetailDocs(docs) {
      const container = document.getElementById('detailDocsList');
      document.getElementById('detailDocCount').textContent = `(${docs.length} ไฟล์)`;
      container.innerHTML = '';

      if (docs.length === 0) {
        container.innerHTML = `<p class="text-xs text-slate-400 italic p-1">ยังไม่มีไฟล์เอกสารแนบ</p>`;
        return;
      }

      docs.forEach(doc => {
        const ext = getFileExt(doc.name || '');
        let iconClass = "fa-file-lines text-slate-600";
        let bgClass = "bg-white border-slate-300";
        if (ext.includes('pdf')) { iconClass = "fa-file-pdf text-rose-600"; bgClass = "bg-rose-50 border-rose-200"; }
        else if (ext.includes('doc')) { iconClass = "fa-file-word text-blue-600"; bgClass = "bg-blue-50 border-blue-200"; }
        else if (ext.includes('xls')) { iconClass = "fa-file-excel text-emerald-600"; bgClass = "bg-emerald-50 border-emerald-200"; }

        const src = fileSrc(doc);
        const isDrive = !isLocalFile(src);
        const isPdfOnDrive = isDrive && ext.includes('pdf') && drivePreviewUrl(src);

        const a = document.createElement('a');
        a.href = src || '#';
        a.target = '_blank';
        a.rel = 'noopener';
        if (!isDrive && doc.name) a.download = doc.name;

        // PDFs stored on Drive render inline instead of opening a new tab
        if (isPdfOnDrive) {
          a.onclick = (e) => {
            e.preventDefault();
            openLightboxModal('drive', drivePreviewUrl(src), doc.name || 'PDF', driveOpenUrl(src));
          };
        }
        a.className = `flex items-center gap-2 p-2 rounded-xl border ${bgClass} shadow-sm text-xs max-w-xs hover:shadow-md transition`;
        a.innerHTML = `
          <i class="fa-solid ${iconClass} text-base shrink-0"></i>
          <div class="truncate">
            <p class="font-bold text-slate-800 truncate" title="${esc(doc.name || '')}">${esc(doc.name || 'ไฟล์แนบ')}</p>
            <p class="text-[10px] text-slate-500">${esc(doc.size || '')} ${isPdfOnDrive ? '· ดูในแอป' : isDrive ? '· เปิดใน Drive' : '· ดาวน์โหลด'}</p>
          </div>
          <i class="fa-solid ${isPdfOnDrive ? 'fa-eye' : isDrive ? 'fa-up-right-from-square' : 'fa-download'} text-slate-400 ml-auto shrink-0"></i>
        `;
        container.appendChild(a);
      });
    }

    function renderDetailVideos(videos) {
      const container = document.getElementById('detailVideosList');
      document.getElementById('detailVideoCount').textContent = `(${videos.length} คลิป)`;
      container.innerHTML = '';

      if (videos.length === 0) {
        container.innerHTML = `<p class="text-xs text-slate-400 italic p-1">ยังไม่มีวิดีโอแนบ</p>`;
        return;
      }

      videos.forEach(vid => {
        const poster = videoThumb(vid);
        const div = document.createElement('div');
        div.className = "relative w-28 h-20 rounded-xl overflow-hidden border border-purple-300 bg-purple-950 text-white flex items-center justify-center cursor-pointer shrink-0 shadow-sm group";
        div.onclick = () => openVideoPreview(vid);
        div.innerHTML = `
          ${poster ? `<img src="${escUrl(poster)}" loading="lazy" alt="" class="absolute inset-0 w-full h-full object-cover opacity-70 group-hover:opacity-90 transition" onerror="this.remove()">` : ''}
          <div class="relative text-center p-1">
            <i class="fa-solid fa-circle-play text-2xl text-white drop-shadow-lg group-hover:scale-125 transition"></i>
          </div>
          <p class="absolute bottom-0 inset-x-0 text-[9px] truncate px-1 py-0.5 bg-slate-900/80 text-purple-100">${esc(vid.name || 'video')}</p>
        `;
        container.appendChild(div);
      });
    }

    // MODAL 1: CREATE NEW TASK
    // Highest existing number + 1. Counting rows instead (the old approach)
    // repeats an ID as soon as any job has been deleted, and a duplicate Job_ID
    // makes every later edit or delete hit the wrong row in the sheet.
    function suggestNextJobId() {
      let max = 0;
      state.tasks.forEach(t => {
        const m = /(\d+)\s*$/.exec(String(t.Job_ID || '').trim());
        if (m) max = Math.max(max, parseInt(m[1], 10));
      });
      return 'PTS-ENG-' + String(max + 1).padStart(3, '0');
    }

    function openCreateModal() {
      const form = document.getElementById('createJobForm');
      if (form) form.reset();
      
      const setVal = (id, val) => {
        const el = document.getElementById(id);
        if (el) el.value = val || '';
      };

      setVal('newJobId', suggestNextJobId());

      const todayStr = new Date().toISOString().split('T')[0];
      setVal('newPoApprovalDate', todayStr);
      setVal('newTargetDate', todayStr);
      
      const futureDate = new Date();
      futureDate.setDate(futureDate.getDate() + 30);
      setVal('newContractExpiryDate', futureDate.toISOString().split('T')[0]);

      state.tempNewPhotos = [];
      state.tempNewDocs = [];
      state.tempNewVideos = [];
      state.tempNewDelivery = [];

      setVal('newSiteContactPhone', '');
      setVal('newSiteMapUrl', '');
      renderOptionLists();
      showAutofillHint('new', '');

      renderPhotosPreview('new');
      renderDocsPreview('new');
      renderVideosPreview('new');
      renderDeliveryPreview('new');

      const modal = document.getElementById('createModal');
      if (modal) modal.classList.remove('hidden');
    }

    function closeCreateModal() {
      const modal = document.getElementById('createModal');
      if (modal) modal.classList.add('hidden');
    }

    async function handleCreateJob(e) {
      e.preventDefault();

      const getVal = (id, def = '') => {
        const el = document.getElementById(id);
        return el ? el.value : def;
      };

      if (photoPrepBusy('new')) {
        alert("ระบบกำลังเตรียมรูปภาพที่เลือกไว้ กรุณารอสักครู่แล้วกดบันทึกอีกครั้ง (รูปทั้งหมดจะถูกบันทึกพร้อมกันในครั้งเดียว)");
        return;
      }
      if (getVal('newSubDept') === ADD_NEW_VALUE) {
        alert("กรุณาเลือกฝ่ายย่อยก่อนบันทึก");
        return;
      }
      // The text boxes behind the dropdowns may be hidden, so the browser's own
      // "required" check cannot reach them - validate here instead.
      if (!comboValueOrWarn('newProjectName', 'ชื่อไซต์งาน / ลูกค้า')) return;
      if (!comboValueOrWarn('newTechnician', 'หัวหน้าช่างผู้รับผิดชอบ')) return;

      const jobId = getVal('newJobId').trim();
      if (!jobId) {
        alert("กรุณาระบุรหัสใบงาน");
        return;
      }
      if (state.tasks.some(t => String(t.Job_ID).trim() === jobId)) {
        alert(`รหัสใบงาน ${jobId} ถูกใช้ไปแล้ว กรุณาเปลี่ยนรหัสใหม่\n(รหัสว่างถัดไป: ${suggestNextJobId()})`);
        return;
      }

      const jsaEl = document.getElementById('newJsaCompleted');
      const isJsaChecked = jsaEl ? jsaEl.checked : false;

      const newTask = {
        Job_ID: jobId,
        Project_Name: getVal('newProjectName'),
        Sub_Department: getVal('newSubDept', 'งานโครงการ'),
        Technician_In_Charge: getVal('newTechnician'),
        Task_Detail: getVal('newTaskDetail'),
        Status: getVal('newStatus', 'วางแผน / เตรียมอุปกรณ์'),
        JSA_Completed: isJsaChecked ? 'Yes' : 'No',
        Priority: getVal('newPriority', 'Medium'),
        Target_Date: getVal('newTargetDate'),
        PO_Approval_Date: getVal('newPoApprovalDate'),
        Contract_Expiry_Date: getVal('newContractExpiryDate'),
        Completion_Date: '',
        Site_Location: getVal('newSiteLocation'),
        Site_Contact_Phone: getVal('newSiteContactPhone'),
        Site_Map_Url: getVal('newSiteMapUrl'),
        Notes_Issues: `[${getNowFormatted()}] เปิดใบงานใหม่ในระบบ`,
        Updated_At: getNowFormatted(),
        Site_Photos: JSON.stringify(state.tempNewPhotos || []),
        Document_Files: JSON.stringify(state.tempNewDocs || []),
        Video_Files: JSON.stringify(state.tempNewVideos || []),
        Delivery_Doc: JSON.stringify(state.tempNewDelivery || [])
      };

      const success = await withBusy('createSubmitBtn', 'กำลังบันทึก...',
        () => saveTaskToDb(newTask, true));
      if (success) {
        closeCreateModal();
      }
    }

    // MODAL 2: EDIT & DETAILS MODAL
    function openEditModal(jobId) {
      const task = state.tasks.find(t => String(t.Job_ID).trim() === String(jobId).trim());
      if (!task) return;

      state.currentEditingTask = { ...task };
      state.tempEditPhotos = getTaskAttachments(task, 'Site_Photos');
      state.tempEditDocs = getTaskAttachments(task, 'Document_Files');
      state.tempEditVideos = getTaskAttachments(task, 'Video_Files');
      state.tempEditDelivery = getTaskAttachments(task, 'Delivery_Doc');

      const setVal = (id, val) => {
        const el = document.getElementById(id);
        if (el) el.value = val || '';
      };
      const setText = (id, val) => {
        const el = document.getElementById(id);
        if (el) el.textContent = val || '';
      };

      setText('editJobIdBadge', task.Job_ID);
      setText('editProjectTitle', task.Project_Name);
      setText('editSubDeptSubtitle', task.Sub_Department);

      renderOptionLists();
      showAutofillHint('edit', '');
      setVal('editProjectName', task.Project_Name);
      setVal('editTechnician', task.Technician_In_Charge);
      syncCombo('editProjectName');
      syncCombo('editTechnician');
      setVal('editSubDept', task.Sub_Department || 'งานโครงการ');
      setVal('editPriority', task.Priority || 'Medium');
      setVal('editSiteLocation', task.Site_Location);
      setVal('editSiteContactPhone', task.Site_Contact_Phone);
      setVal('editSiteMapUrl', task.Site_Map_Url);
      setVal('editTaskDetail', task.Task_Detail);
      
      setVal('editPoApprovalDate', formatDateForInput(task.PO_Approval_Date));
      setVal('editContractExpiryDate', formatDateForInput(task.Contract_Expiry_Date));
      setVal('editCompletionDate', formatDateForInput(task.Completion_Date));

      const statusSelect = document.getElementById('editStatusSelect');
      if (statusSelect) statusSelect.value = normalizeStatus(task.Status);

      setText('editUpdatedAt', `อัปเดตล่าสุด: ${task.Updated_At || '-'}`);
      setText('editNotesDisplay', task.Notes_Issues || 'ยังไม่มีบันทึกเพิ่มเติม');
      setVal('newAppendNote', '');

      renderStatusStepper(task.Status);
      renderJsaBanner(task.JSA_Completed);
      renderContractTimelineBanner(task);

      renderPhotosPreview('edit');
      renderDocsPreview('edit');
      renderVideosPreview('edit');
      renderDeliveryPreview('edit');

      document.getElementById('editModal').classList.remove('hidden');
    }

    function closeEditModal() {
      state.currentEditingTask = null;
      document.getElementById('editModal').classList.add('hidden');
    }

    function renderContractTimelineBanner(task) {
      const banner = document.getElementById('editContractTimelineBanner');
      if (!banner) return;
      const isCompleted = task.Status === 'ส่งมอบงานแล้ว' || task.Status === 'ปิดงาน' || task.Status === 'ส่งมอบ / เอกสารเรียบร้อย';
      const daysRem = getDaysUntilExpiry(task.Contract_Expiry_Date);

      if (isCompleted) {
        banner.className = "p-3.5 rounded-2xl border border-emerald-200 bg-emerald-50 flex items-center justify-between";
        banner.innerHTML = `
          <div class="flex items-center space-x-3">
            <div class="w-9 h-9 rounded-xl bg-emerald-200 text-emerald-800 flex items-center justify-center text-base">
              <i class="fa-solid fa-flag-checkered"></i>
            </div>
            <div>
              <h4 class="text-xs font-bold text-emerald-950">สถานะสัญญา: ส่งมอบงานเสร็จสมบูรณ์แล้ว</h4>
              <p class="text-[11px] text-emerald-700">วันที่อนุมัติ PO: ${formatDisplayDate(task.PO_Approval_Date)} | ครบสัญญา: ${formatDisplayDate(task.Contract_Expiry_Date)} | งานเสร็จจริง: ${formatDisplayDate(task.Completion_Date)}</p>
            </div>
          </div>
        `;
      } else if (daysRem !== null && daysRem < 0) {
        banner.className = "p-3.5 rounded-2xl border border-rose-300 bg-rose-50 flex items-center justify-between animate-pulse";
        banner.innerHTML = `
          <div class="flex items-center space-x-3">
            <div class="w-9 h-9 rounded-xl bg-rose-200 text-pts-900 flex items-center justify-center text-base">
              <i class="fa-solid fa-bell text-rose-600"></i>
            </div>
            <div>
              <h4 class="text-xs font-bold text-pts-950">🚨 แจ้งเตือน: สัญญาเลยกำหนดเวลาแล้ว ${Math.abs(daysRem)} วัน!</h4>
              <p class="text-[11px] text-rose-700">วันครบกำหนดสัญญา: ${formatDisplayDate(task.Contract_Expiry_Date)} | โปรดเร่งรัดดำเนินการส่งมอบงาน</p>
            </div>
          </div>
        `;
      } else if (daysRem !== null && daysRem <= 7) {
        banner.className = "p-3.5 rounded-2xl border border-amber-300 bg-amber-50 flex items-center justify-between animate-pulse";
        banner.innerHTML = `
          <div class="flex items-center space-x-3">
            <div class="w-9 h-9 rounded-xl bg-amber-200 text-amber-900 flex items-center justify-center text-base">
              <i class="fa-solid fa-clock-rotate-left text-amber-700"></i>
            </div>
            <div>
              <h4 class="text-xs font-bold text-amber-950">⚠️ แจ้งเตือนสัญญา: จะครบกำหนดภายในอีก ${daysRem} วัน</h4>
              <p class="text-[11px] text-amber-800">วันครบกำหนดสัญญา: ${formatDisplayDate(task.Contract_Expiry_Date)} (วันอนุมัติ PO: ${formatDisplayDate(task.PO_Approval_Date)})</p>
            </div>
          </div>
        `;
      } else {
        banner.className = "p-3.5 rounded-2xl border border-slate-200 bg-slate-50 flex items-center justify-between";
        banner.innerHTML = `
          <div class="flex items-center space-x-3">
            <div class="w-9 h-9 rounded-xl bg-slate-200 text-slate-700 flex items-center justify-center text-base">
              <i class="fa-regular fa-calendar-check"></i>
            </div>
            <div>
              <h4 class="text-xs font-bold text-slate-800">กำหนดเวลาตามสัญญา (Contract Timeline)</h4>
              <p class="text-[11px] text-slate-500">วันอนุมัติ PO: ${formatDisplayDate(task.PO_Approval_Date)} | ครบกำหนดสัญญา: ${formatDisplayDate(task.Contract_Expiry_Date)}</p>
            </div>
          </div>
        `;
      }
    }

    function renderStatusStepper(currentStatus) {
      const container = document.getElementById('statusStepperContainer');
      if (!container) return;
      container.innerHTML = '';

      ALL_STATUSES.forEach((st, idx) => {
        const isCurrent = normalizeStatus(st) === normalizeStatus(currentStatus);
        const btn = document.createElement('button');
        btn.type = 'button';
        btn.onclick = () => changeEditingStatus(st);

        let btnClass = "p-2 rounded-xl text-[11px] font-bold text-center border transition flex flex-col items-center justify-center min-h-[48px] cursor-pointer";
        if (isCurrent) {
          btnClass += " bg-pts-800 text-white border-pts-900 shadow-md ring-2 ring-pts-800/40";
        } else {
          btnClass += " bg-white text-slate-600 border-slate-200 hover:bg-slate-100 hover:text-slate-900";
        }
        btn.className = btnClass;

        btn.innerHTML = `
          <span class="text-[10px] opacity-75">${idx + 1}.</span>
          <span class="leading-tight line-clamp-2">${st}</span>
        `;
        container.appendChild(btn);
      });
    }

    function changeEditingStatus(newStatus) {
      if (!state.currentEditingTask) return;
      const oldStatus = state.currentEditingTask.Status;
      state.currentEditingTask.Status = newStatus;

      const select = document.getElementById('editStatusSelect');
      if (select && select.value !== newStatus) {
        select.value = newStatus;
      }

      const nowStr = getNowFormatted();

      if (newStatus === 'ส่งมอบงานแล้ว' || newStatus === 'ปิดงาน' || newStatus === 'ส่งมอบ / เอกสารเรียบร้อย') {
        const todayIso = new Date().toISOString().split('T')[0];
        state.currentEditingTask.Completion_Date = todayIso;
        const compEl = document.getElementById('editCompletionDate');
        if (compEl) compEl.value = todayIso;
      }

      if (normalizeStatus(oldStatus) !== normalizeStatus(newStatus)) {
        const logEntry = `\n[${nowStr}] เปลี่ยนสถานะ: ${oldStatus} ➔ ${newStatus}`;
        state.currentEditingTask.Notes_Issues = (state.currentEditingTask.Notes_Issues || '') + logEntry;
        const notesDisplay = document.getElementById('editNotesDisplay');
        if (notesDisplay) notesDisplay.textContent = state.currentEditingTask.Notes_Issues;
      }

      renderStatusStepper(newStatus);
      renderContractTimelineBanner(state.currentEditingTask);
    }

    function renderJsaBanner(jsaStatus) {
      const isOk = (jsaStatus || "").toLowerCase() === 'yes';
      const banner = document.getElementById('editJsaBanner');
      const iconBox = document.getElementById('editJsaIconBox');
      const title = document.getElementById('editJsaTitle');
      const btn = document.getElementById('editJsaToggleBtn');
      if (!banner || !iconBox || !title || !btn) return;

      if (isOk) {
        banner.className = "p-4 rounded-2xl border border-emerald-200 bg-emerald-50 flex items-center justify-between";
        iconBox.className = "w-10 h-10 rounded-xl bg-emerald-200 text-emerald-800 flex items-center justify-center text-lg";
        title.className = "text-sm font-bold text-emerald-950";
        title.textContent = "JSA Clearance: APPROVED (อนุมัติความปลอดภัยเรียบร้อย)";
        btn.className = "px-4 py-2 rounded-xl text-xs font-bold bg-white text-emerald-800 border border-emerald-300 hover:bg-emerald-100 transition shadow-sm";
        btn.innerHTML = `<i class="fa-solid fa-check text-emerald-600"></i><span>อนุมัติแล้ว (กดเพื่อยกเลิก)</span>`;
      } else {
        banner.className = "p-4 rounded-2xl border border-amber-200 bg-amber-50 flex items-center justify-between";
        iconBox.className = "w-10 h-10 rounded-xl bg-amber-200 text-amber-800 flex items-center justify-center text-lg";
        title.className = "text-sm font-bold text-amber-950";
        title.textContent = "JSA Clearance: PENDING (รออนุมัติความปลอดภัย)";
        btn.className = "px-4 py-2 rounded-xl text-xs font-bold bg-emerald-600 text-white hover:bg-emerald-700 transition shadow-md";
        btn.innerHTML = `<i class="fa-solid fa-shield-check"></i><span>อนุมัติ JSA เดี๋ยวนี้</span>`;
      }
    }

    function toggleEditJsa() {
      if (!state.currentEditingTask) return;
      const current = (state.currentEditingTask.JSA_Completed || "").toLowerCase() === 'yes';
      const newJsa = current ? 'No' : 'Yes';
      state.currentEditingTask.JSA_Completed = newJsa;

      const logMsg = newJsa === 'Yes' ? 'อนุมัติการตรวจสอบความปลอดภัย JSA' : 'ยกเลิกสถานะอนุมัติ JSA';
      state.currentEditingTask.Notes_Issues = (state.currentEditingTask.Notes_Issues || '') + `\n[${getNowFormatted()}] ${logMsg}`;

      const notesDisplay = document.getElementById('editNotesDisplay');
      if (notesDisplay) notesDisplay.textContent = state.currentEditingTask.Notes_Issues;
      renderJsaBanner(newJsa);
    }

    function appendSiteNote() {
      const input = document.getElementById('newAppendNote');
      if (!input) return;
      const text = input.value.trim();
      if (!text || !state.currentEditingTask) return;

      const logEntry = `\n[${getNowFormatted()}] หมายเหตุ: ${text}`;
      state.currentEditingTask.Notes_Issues = (state.currentEditingTask.Notes_Issues || '') + logEntry;
      const notesDisplay = document.getElementById('editNotesDisplay');
      if (notesDisplay) notesDisplay.textContent = state.currentEditingTask.Notes_Issues;
      input.value = '';
    }

    async function saveTaskEdits() {
      if (!state.currentEditingTask) return;
      if (photoPrepBusy('edit')) {
        alert("ระบบกำลังเตรียมรูปภาพที่เลือกไว้ กรุณารอสักครู่แล้วกดบันทึกอีกครั้ง (รูปทั้งหมดจะถูกบันทึกพร้อมกันในครั้งเดียว)");
        return;
      }

      const getVal = (id, def = '') => {
        const el = document.getElementById(id);
        return el ? el.value : def;
      };
      if (getVal('editSubDept') === ADD_NEW_VALUE) {
        alert("กรุณาเลือกฝ่ายย่อยก่อนบันทึก");
        return;
      }

      state.currentEditingTask.Project_Name = getVal('editProjectName');
      state.currentEditingTask.Technician_In_Charge = getVal('editTechnician');
      state.currentEditingTask.Sub_Department = getVal('editSubDept', 'งานโครงการ');
      state.currentEditingTask.Priority = getVal('editPriority', 'Medium');
      state.currentEditingTask.Site_Location = getVal('editSiteLocation');
      state.currentEditingTask.Site_Contact_Phone = getVal('editSiteContactPhone');
      state.currentEditingTask.Site_Map_Url = getVal('editSiteMapUrl');
      state.currentEditingTask.Task_Detail = getVal('editTaskDetail');
      
      state.currentEditingTask.PO_Approval_Date = getVal('editPoApprovalDate');
      state.currentEditingTask.Contract_Expiry_Date = getVal('editContractExpiryDate');
      state.currentEditingTask.Completion_Date = getVal('editCompletionDate');

      state.currentEditingTask.Updated_At = getNowFormatted();
      state.currentEditingTask.Site_Photos = JSON.stringify(state.tempEditPhotos || []);
      state.currentEditingTask.Document_Files = JSON.stringify(state.tempEditDocs || []);
      state.currentEditingTask.Video_Files = JSON.stringify(state.tempEditVideos || []);
      state.currentEditingTask.Delivery_Doc = JSON.stringify(state.tempEditDelivery || []);

      const success = await withBusy('editSaveBtn', 'กำลังบันทึก...',
        () => saveTaskToDb(state.currentEditingTask, false));
      if (success) {
        closeEditModal();
      }
    }

    async function deleteCurrentTask() {
      if (!state.currentEditingTask) return;
      const jobId = state.currentEditingTask.Job_ID;
      if (confirm(`คุณแน่ใจหรือไม่ที่จะลบใบงานรหัส ${jobId} ?`)) {
        const success = await deleteTaskFromDb(jobId);
        if (success) {
          closeEditModal();
        }
      }
    }

    // HELPER UTILS
    function getNowFormatted() {
      const d = new Date();
      const pad = n => String(n).padStart(2, '0');
      return `${pad(d.getDate())}/${pad(d.getMonth()+1)}/${d.getFullYear()} ${pad(d.getHours())}:${pad(d.getMinutes())}`;
    }
    function getNowFormatted() {
      const d = new Date();
      const pad = n => String(n).padStart(2, '0');
      return `${pad(d.getDate())}/${pad(d.getMonth()+1)}/${d.getFullYear()} ${pad(d.getHours())}:${pad(d.getMinutes())}`;
    }
  


    /* ---------------------------------------------------------------------
     * MODAL 5: DELIVERY DOCUMENT & WORK HANDOVER MODAL
     * ------------------------------------------------------------------- */
    let currentDeliveryTab = 'form';

    function getTaskDeliveryFiles(task) {
      if (!task) return [];
      const direct = getTaskAttachments(task, 'Delivery_Doc');
      if (direct.length > 0) return direct;
      // Fallback: check Document_Files for handover / delivery keywords
      const docs = getTaskAttachments(task, 'Document_Files');
      return docs.filter(d => {
        const name = (d.name || '').toLowerCase();
        return name.includes('ส่งมอบ') || name.includes('ตรวจรับ') || name.includes('delivery') || name.includes('handover') || name.includes('รับรอง');
      });
    }

    function openDeliveryDocModal(jobId) {
      const task = state.tasks.find(t => t.Job_ID === jobId);
      if (!task) return;

      state.currentDeliveryJobId = jobId;
      // Files staged for a different job never carry over
      if (state.deliveryPendingJobId !== jobId) {
        state.deliveryPending = [];
        state.deliveryPendingJobId = jobId;
      }

      document.getElementById('deliveryJobIdBadge').textContent = task.Job_ID;
      document.getElementById('deliverySubtitle').textContent = `${task.Project_Name || '-'} · ${task.Sub_Department || '-'}`;

      // Form Sheet Data
      document.getElementById('formSheetJobId').textContent = task.Job_ID;
      document.getElementById('formSheetDate').textContent = `วันที่ส่งมอบ: ${formatDisplayDate(task.Completion_Date || task.Target_Date || getNowFormatted().split(' ')[0])}`;
      document.getElementById('formSheetProject').textContent = task.Project_Name || '-';
      document.getElementById('formSheetSubDept').textContent = task.Sub_Department || '-';
      document.getElementById('formSheetLocation').textContent = task.Site_Location || '-';
      document.getElementById('formSheetContact').textContent = task.Site_Contact_Phone || '-';
      document.getElementById('formSheetPoDate').textContent = formatDisplayDate(task.PO_Approval_Date);
      document.getElementById('formSheetExpiryDate').textContent = formatDisplayDate(task.Contract_Expiry_Date);
      document.getElementById('formSheetTaskDetail').textContent = task.Task_Detail || 'ดำเนินการติดตั้ง / ตรวจสอบระบบตามสัญญาเสร็จสิ้นสมบูรณ์';
      
      const isJsaOk = (task.JSA_Completed || '').toLowerCase() === 'yes';
      document.getElementById('formSheetJsa').innerHTML = isJsaOk
        ? '<i class="fa-solid fa-shield-check mr-1 text-emerald-600"></i>ผ่านการประเมินความปลอดภัยเรียบร้อย (JSA Approved)'
        : '<i class="fa-solid fa-triangle-exclamation mr-1 text-amber-600"></i>ยังไม่ได้รับการรับรอง JSA (Pending)';
      
      document.getElementById('formSheetCompletionDate').textContent = formatDisplayDate(task.Completion_Date || task.Target_Date);
      document.getElementById('formSheetTechName').textContent = `(${task.Technician_In_Charge || 'หัวหน้าช่างผู้รับผิดชอบ'})`;

      // Attached Delivery Files
      const deliveryFiles = getTaskDeliveryFiles(task);
      const attachedContainer = document.getElementById('deliveryAttachedFilesList');
      const badge = document.getElementById('deliveryFilesBadge');
      badge.textContent = `(${deliveryFiles.length})`;
      attachedContainer.innerHTML = '';

      if (deliveryFiles.length === 0) {
        attachedContainer.innerHTML = `
          <div class="col-span-full py-8 text-center text-slate-400 w-full bg-slate-50 rounded-xl border border-dashed border-slate-300">
            <i class="fa-solid fa-folder-open text-3xl text-slate-300 mb-2 block"></i>
            <p class="text-xs font-bold text-slate-700">ยังไม่มีไฟล์แนบใบส่งมอบงานที่บันทึกไว้ในระบบ</p>
            <p class="text-[11px] text-slate-500 mt-1">กดปุ่ม <span class="font-semibold text-blue-700">+ เลือกไฟล์ (หลายไฟล์)</span> ด้านบนเพื่อเลือก PDF, สแกน หรือรูปภาพได้หลายไฟล์พร้อมกัน แล้วกด <span class="font-semibold text-emerald-700">บันทึกทั้งหมด</span> ครั้งเดียว</p>
          </div>
        `;
      } else {
        deliveryFiles.forEach((doc, idx) => {
          const ext = getFileExt(doc.name || '');
          let iconClass = "fa-file-lines text-slate-600";
          let bgClass = "bg-white border-slate-300";
          if (ext.includes('pdf')) { iconClass = "fa-file-pdf text-rose-600"; bgClass = "bg-rose-50 border-rose-200"; }
          else if (ext.includes('doc')) { iconClass = "fa-file-word text-blue-600"; bgClass = "bg-blue-50 border-blue-200"; }
          else if (ext.includes('jpg') || ext.includes('jpeg') || ext.includes('png')) { iconClass = "fa-file-image text-emerald-600"; bgClass = "bg-emerald-50 border-emerald-200"; }

          const src = fileSrc(doc);
          const isDrive = !isLocalFile(src);
          const isPdfOnDrive = isDrive && ext.includes('pdf') && drivePreviewUrl(src);

          const itemDiv = document.createElement('div');
          itemDiv.className = `flex items-center justify-between gap-2.5 p-3 rounded-xl border ${bgClass} shadow-sm text-xs w-full sm:max-w-sm hover:shadow-md transition bg-white`;
          
          let previewAction = ``;
          if (isPdfOnDrive) {
            previewAction = `openLightboxModal('drive', '${escUrl(drivePreviewUrl(src))}', '${esc(doc.name || 'PDF')}', '${escUrl(driveOpenUrl(src))}');`;
          } else if (ext.includes('jpg') || ext.includes('jpeg') || ext.includes('png')) {
            previewAction = `openLightboxModal('photo', '${escUrl(src)}', '${esc(doc.name || 'ใบส่งมอบงาน')}');`;
          } else {
            previewAction = `window.open('${escUrl(src)}', '_blank');`;
          }

          itemDiv.innerHTML = `
            <div class="flex items-center gap-2.5 min-w-0 cursor-pointer flex-1" onclick="${previewAction}">
              <i class="fa-solid ${iconClass} text-xl shrink-0"></i>
              <div class="truncate">
                <p class="font-bold text-slate-800 truncate hover:text-blue-700" title="${esc(doc.name || '')}">${esc(doc.name || 'ใบส่งมอบงาน')}</p>
                <p class="text-[10px] text-slate-500">${esc(doc.size || '')} ${isPdfOnDrive ? '· ดูในแอป' : isDrive ? '· ใน Google Drive' : '· บันทึกในระบบแล้ว'}</p>
              </div>
            </div>
            <div class="flex items-center gap-1 shrink-0">
              <button type="button" onclick="${previewAction}" class="p-1.5 rounded-lg text-blue-600 hover:bg-blue-50 transition" title="เปิดดูเอกสาร">
                <i class="fa-solid ${isPdfOnDrive ? 'fa-eye' : isDrive ? 'fa-up-right-from-square' : 'fa-download'}"></i>
              </button>
              <button type="button" onclick="removeDeliveryFileFromModal(${idx})" class="p-1.5 rounded-lg text-rose-600 hover:bg-rose-50 transition" title="ลบไฟล์ออกจากระบบ">
                <i class="fa-solid fa-trash-can"></i>
              </button>
            </div>
          `;
          attachedContainer.appendChild(itemDiv);
        });
      }

      renderDeliveryPending();

      const hasPending = (state.deliveryPending || []).length > 0;
      if (deliveryFiles.length > 0 || hasPending) {
        switchDeliveryTab('files');
      } else {
        switchDeliveryTab('form');
      }

      document.getElementById('deliveryModal').classList.remove('hidden');
    }

    function closeDeliveryModal() {
      const pending = (state.deliveryPending || []).length;
      if (pending > 0 && !confirm(`มีไฟล์ที่เลือกไว้แต่ยังไม่ได้บันทึก ${pending} ไฟล์\n\nปิดหน้าต่างโดยไม่บันทึกใช่หรือไม่?`)) {
        return;
      }
      state.deliveryPending = [];
      state.deliveryPendingJobId = null;
      state.currentDeliveryJobId = null;
      document.getElementById('deliveryModal').classList.add('hidden');
    }

    function switchDeliveryTab(tab) {
      currentDeliveryTab = tab;
      const formView = document.getElementById('deliveryFormView');
      const filesView = document.getElementById('deliveryFilesView');
      const tabFormBtn = document.getElementById('tabDeliveryFormBtn');
      const tabFilesBtn = document.getElementById('tabDeliveryFilesBtn');
      
      const headerPrintBtn = document.getElementById('deliveryHeaderPrintBtn');
      const headerUploadBtn = document.getElementById('deliveryHeaderUploadBtn');
      const footerPrintBtn = document.getElementById('deliveryFooterPrintBtn');
      const footerUploadBtn = document.getElementById('deliveryFooterUploadBtn');

      if (tab === 'form') {
        formView.classList.remove('hidden');
        filesView.classList.add('hidden');
        tabFormBtn.className = "px-3.5 py-1.5 rounded-xl text-xs font-bold bg-white text-blue-900 shadow-sm border border-slate-200 transition";
        tabFilesBtn.className = "px-3.5 py-1.5 rounded-xl text-xs font-bold bg-slate-200 text-slate-700 hover:bg-white transition";
        
        if (headerPrintBtn) headerPrintBtn.classList.remove('hidden');
        if (headerUploadBtn) headerUploadBtn.classList.add('hidden');
        if (footerPrintBtn) footerPrintBtn.classList.remove('hidden');
        if (footerUploadBtn) footerUploadBtn.classList.add('hidden');
      } else {
        formView.classList.add('hidden');
        filesView.classList.remove('hidden');
        tabFilesBtn.className = "px-3.5 py-1.5 rounded-xl text-xs font-bold bg-white text-blue-900 shadow-sm border border-slate-200 transition";
        tabFormBtn.className = "px-3.5 py-1.5 rounded-xl text-xs font-bold bg-slate-200 text-slate-700 hover:bg-white transition";

        if (headerPrintBtn) headerPrintBtn.classList.add('hidden');
        if (headerUploadBtn) headerUploadBtn.classList.remove('hidden');
        if (footerPrintBtn) footerPrintBtn.classList.add('hidden');
        if (footerUploadBtn) footerUploadBtn.classList.remove('hidden');
      }
    }

    async function removeDeliveryFileFromModal(index) {
      const jobId = state.currentDeliveryJobId;
      if (!jobId) return;
      const task = state.tasks.find(t => t.Job_ID === jobId);
      if (!task) return;

      if (!confirm('คุณต้องการลบไฟล์ใบส่งมอบงานนี้ออกจากระบบใช่หรือไม่?')) return;

      const current = getTaskAttachments(task, 'Delivery_Doc');
      if (index >= 0 && index < current.length) {
        current.splice(index, 1);
        task.Delivery_Doc = JSON.stringify(current);
        await saveTaskToDb(task, false);
        openDeliveryDocModal(jobId);
        showSyncToast('ลบและอัปเดตไฟล์ในระบบเรียบร้อย');
      }
    }

    function printDeliverySlip() {
      const sheet = document.getElementById('printableHandoverSheet');
      if (!sheet) return;

      const printWin = window.open('', '', 'width=900,height=750');
      printWin.document.write(`
        <html>
          <head>
            <title>ใบส่งมอบงาน - ${document.getElementById('formSheetJobId').textContent}</title>
            <script src="https://cdn.tailwindcss.com"></script>
            <link href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css" rel="stylesheet">
            <link href="https://fonts.googleapis.com/css2?family=Prompt:wght@300;400;500;600;700&display=swap" rel="stylesheet">
            <style>
              body { font-family: 'Prompt', sans-serif; padding: 20px; }
              @media print {
                body { padding: 0; }
                button { display: none !important; }
              }
            </style>
          </head>
          <body class="bg-white">
            <div class="max-w-3xl mx-auto">
              ${sheet.outerHTML}
            </div>
            <script>
              window.onload = function() {
                setTimeout(function() {
                  window.print();
                }, 400);
              };
            </script>
          </body>
        </html>
      `);
      printWin.document.close();
    }

    function editFromDeliveryModal() {
      const jobId = state.currentDeliveryJobId;
      closeDeliveryModal();
      // The user may have chosen to keep the dialog open to save pending files
      if (state.currentDeliveryJobId) return;
      if (jobId) openEditModal(jobId);
    }

    /* -----------------------------------------------------------------------
     * DELIVERY DOCS - STAGE, THEN SAVE ONCE
     *
     * Picking a file here used to fire a full save straight away, so attaching
     * five scans meant five round-trips and five progress dialogs. Files now
     * pile up in a pending list and one press of "บันทึกทั้งหมด" sends them all.
     * --------------------------------------------------------------------- */
    async function handleDeliveryModalUpload(event) {
      const jobId = state.currentDeliveryJobId;
      if (!jobId) return;

      const input = event.target;
      const files = Array.from(input.files || []);
      input.value = '';
      if (files.length === 0) return;

      const tooBig = files.filter(f => f.size > MAX_ATTACHMENT_MB * 1024 * 1024);
      if (tooBig.length > 0) {
        alert(`ไฟล์ต่อไปนี้มีขนาดเกิน ${MAX_ATTACHMENT_MB} MB จึงไม่ถูกเพิ่ม:\n\n` + tooBig.map(f => `• ${f.name} (${formatBytes(f.size)})`).join('\n'));
      }
      const usableFiles = files.filter(f => f.size <= MAX_ATTACHMENT_MB * 1024 * 1024);
      if (usableFiles.length === 0) return;

      const results = await Promise.all(usableFiles.map(file => readFileAsDataUrl(file).then(dataUrl => (
        dataUrl ? {
          name: file.name,
          size: formatBytes(file.size),
          type: file.type || getFileExt(file.name),
          dataUrl: dataUrl
        } : null
      ))));

      const usable = results.filter(Boolean);
      if (usable.length === 0) return;

      if (state.deliveryPendingJobId !== jobId) {
        state.deliveryPending = [];
        state.deliveryPendingJobId = jobId;
      }
      usable.forEach(item => state.deliveryPending.push(item));
      renderDeliveryPending();
      switchDeliveryTab('files');
    }

    function renderDeliveryPending() {
      const section = document.getElementById('deliveryPendingSection');
      const list = document.getElementById('deliveryPendingList');
      const count = document.getElementById('deliveryPendingCount');
      const footerBtn = document.getElementById('deliveryFooterSaveBtn');
      const footerCount = document.getElementById('deliveryFooterSaveCount');
      if (!section || !list) return;

      const pending = (state.deliveryPendingJobId === state.currentDeliveryJobId) ? (state.deliveryPending || []) : [];
      if (count) count.textContent = `(${pending.length})`;
      if (footerCount) footerCount.textContent = `(${pending.length})`;

      if (pending.length === 0) {
        section.classList.add('hidden');
        if (footerBtn) { footerBtn.classList.add('hidden'); footerBtn.classList.remove('flex'); }
        list.innerHTML = '';
        return;
      }

      section.classList.remove('hidden');
      if (footerBtn) { footerBtn.classList.remove('hidden'); footerBtn.classList.add('flex'); }

      list.innerHTML = '';
      pending.forEach((doc, idx) => {
        const ext = getFileExt(doc.name || '');
        let iconClass = "fa-file-lines text-slate-600";
        if (ext.includes('pdf')) iconClass = "fa-file-pdf text-rose-600";
        else if (ext.includes('doc')) iconClass = "fa-file-word text-blue-600";
        else if (ext.includes('xls')) iconClass = "fa-file-excel text-emerald-600";
        else if (/^(jpe?g|png|gif|webp)$/.test(ext)) iconClass = "fa-file-image text-emerald-600";

        const row = document.createElement('div');
        row.className = "flex items-center justify-between gap-2 p-2.5 rounded-xl border border-amber-200 bg-white text-xs";
        row.innerHTML = `
          <div class="flex items-center gap-2 min-w-0">
            <i class="fa-solid ${iconClass} text-lg shrink-0"></i>
            <div class="truncate">
              <p class="font-bold text-slate-800 truncate" title="${esc(doc.name || '')}">${esc(doc.name || 'ไฟล์')}</p>
              <p class="text-[10px] text-amber-700">${esc(doc.size || '')} · รอบันทึก</p>
            </div>
          </div>
          <button type="button" onclick="removeDeliveryPending(${idx})" class="p-1.5 rounded-lg text-rose-600 hover:bg-rose-50 transition shrink-0" title="เอาออกจากรายการ">
            <i class="fa-solid fa-xmark"></i>
          </button>
        `;
        list.appendChild(row);
      });
    }

    function removeDeliveryPending(index) {
      if (!state.deliveryPending) return;
      state.deliveryPending.splice(index, 1);
      renderDeliveryPending();
    }

    function clearDeliveryPending() {
      state.deliveryPending = [];
      renderDeliveryPending();
    }

    async function saveDeliveryPendingFiles() {
      const jobId = state.currentDeliveryJobId;
      if (!jobId) return;
      const task = state.tasks.find(t => t.Job_ID === jobId);
      if (!task) return;

      const pending = (state.deliveryPendingJobId === jobId) ? (state.deliveryPending || []) : [];
      if (pending.length === 0) return;

      // Work on a copy: if the save fails the task in memory must stay as it
      // was, so pressing Save again cannot append the same files twice.
      const current = getTaskAttachments(task, 'Delivery_Doc');
      const draft = { ...task, Delivery_Doc: JSON.stringify(current.concat(pending)) };

      const success = await withBusy('deliverySavePendingBtn', 'กำลังบันทึก...',
        () => saveTaskToDb(draft, false));
      if (!success) return;

      state.deliveryPending = [];
      openDeliveryDocModal(jobId);
      showSyncToast(`บันทึกเอกสารใบส่งมอบงาน ${pending.length} ไฟล์เรียบร้อย`);
    }

    /* =======================================================================
     * MANAGED OPTION LISTS  (ฝ่ายย่อย / รายชื่อช่าง / ลูกค้า)
     *
     * These used to be hard-coded <option> tags, so every new department meant
     * a code change. They are now stored in the "App_Settings" tab of the same
     * spreadsheet (via Code.gs SAVE_SETTINGS) and edited from the ⚙ dialog.
     *
     * Every dropdown shows:  configured list  ∪  values already used by jobs.
     * The union means a value can never "disappear" from a job that uses it,
     * and a freshly typed value shows up on every device once the job is saved.
     *
     * If the deployed Code.gs is older than this file the settings endpoint
     * does not exist; the lists then fall back to localStorage on this device
     * and the dialog says so.
     * ===================================================================== */
    const DEFAULT_SUB_DEPTS = ['งานโครงการ', 'PM Fire Pump', 'PM Fire Alarm', 'PM Generator', 'PM Fire Suppression', 'งานซ่อมบำรุง'];
    const OPTION_GROUPS = {
      sub_departments: { label: 'ฝ่ายย่อย (Sub-Department)', field: 'Sub_Department', icon: 'fa-sitemap', placeholder: 'เช่น PM Sprinkler, งานติดตั้ง', hint: 'แสดงเป็นตัวเลือกในฟอร์มเปิดงาน/แก้ไข และตัวกรองหน้าหลัก' },
      technicians:     { label: 'หัวหน้าช่าง / ผู้รับผิดชอบ', field: 'Technician_In_Charge', icon: 'fa-helmet-safety', placeholder: 'เช่น ช่างสมชาย', hint: 'พิมพ์ในช่องหัวหน้าช่างแล้วระบบจะแนะนำชื่อให้' },
      customers:       { label: 'ลูกค้า / ไซต์งาน', field: 'Project_Name', icon: 'fa-building', placeholder: 'เช่น บริษัท ... จำกัด (นิคม ...)', hint: 'พิมพ์ในช่องชื่อไซต์แล้วระบบจะแนะนำ และดึงสถานที่/เบอร์จากใบงานล่าสุดให้' },
      // `from` says which list the "values already in use" come from: most
      // options are filled in on jobs, task types only ever appear on tasks.
      task_types:      { label: 'ประเภท Task', field: 'Type', from: 'jobTasks', icon: 'fa-tag', placeholder: 'เช่น ติดต่อลูกค้า, เข้าหน้างาน', hint: 'ตัวเลือกประเภทในฟอร์มเพิ่ม Task (ชุดเดียวกับแอพงานขาย)' }
    };
    const LOCAL_OPTIONS_KEY = 'pts_custom_options';
    const ADD_NEW_VALUE = '__add_new__';

    state.options = { sub_departments: [], technicians: [], customers: [], task_types: [] };
    state.optionsSource = 'unknown';   // 'server' | 'local' | 'unknown'

    function loadLocalOptions() {
      try {
        const parsed = JSON.parse(localStorage.getItem(LOCAL_OPTIONS_KEY) || '{}');
        const out = {};
        Object.keys(OPTION_GROUPS).forEach(k => { out[k] = Array.isArray(parsed[k]) ? parsed[k] : []; });
        return out;
      } catch (err) {
        const out = {};
        Object.keys(OPTION_GROUPS).forEach(k => { out[k] = []; });
        return out;
      }
    }

    function saveLocalOptions() {
      try { localStorage.setItem(LOCAL_OPTIONS_KEY, JSON.stringify(state.options)); } catch (err) {}
    }

    function cleanOptionValue(v) {
      return String(v || '').replace(/\s+/g, ' ').trim().slice(0, 120);
    }

    async function fetchOptions() {
      if (state.dbMode === 'mock' || !state.apiUrl) {
        state.options = loadLocalOptions();
        state.optionsSource = 'local';
        renderOptionLists();
        return;
      }
      const { json } = await apiGetDetailed({ settings: '1' });
      if (json && json.status === 'success' && json.settings && typeof json.settings === 'object') {
        Object.keys(OPTION_GROUPS).forEach(k => {
          state.options[k] = Array.isArray(json.settings[k]) ? json.settings[k].map(cleanOptionValue).filter(Boolean) : [];
        });
        state.optionsSource = 'server';
      } else {
        // Old Code.gs answers the plain task list instead - keep working locally
        state.options = loadLocalOptions();
        state.optionsSource = 'local';
      }
      renderOptionLists();
    }

    // Distinct values a field currently holds across all jobs, with usage counts
    function valuesInUse(field, from) {
      const counts = {};
      const rows = from === 'jobTasks' ? (state.jobTasks || []) : (state.tasks || []);
      rows.forEach(t => {
        const v = cleanOptionValue(t[field]);
        if (v) counts[v] = (counts[v] || 0) + 1;
      });
      return counts;
    }

    // Configured list first (in the order the team chose), then anything jobs
    // still use that is not on the list, alphabetically.
    function defaultOptionList(key) {
      if (key === 'sub_departments') return DEFAULT_SUB_DEPTS.slice();
      if (key === 'task_types') return DEFAULT_TASK_TYPES.slice();
      return [];
    }

    function getOptionList(key) {
      const group = OPTION_GROUPS[key];
      const configured = (state.options[key] || []).slice();
      const base = configured.length > 0 ? configured : defaultOptionList(key);
      const seen = new Set(base);
      const extras = Object.keys(valuesInUse(group.field, group.from)).filter(v => !seen.has(v)).sort((a, b) => a.localeCompare(b, 'th'));
      return base.concat(extras);
    }

    function fillSelect(selectEl, values, keepValue, options) {
      if (!selectEl) return;
      options = options || {};
      const current = keepValue !== undefined ? keepValue : selectEl.value;
      selectEl.innerHTML = '';
      if (options.allLabel) {
        const all = document.createElement('option');
        all.value = 'ALL';
        all.textContent = options.allLabel;
        selectEl.appendChild(all);
      }
      values.forEach((v, i) => {
        const opt = document.createElement('option');
        opt.value = v;
        opt.textContent = options.numbered ? (i + 1) + '. ' + v : v;
        selectEl.appendChild(opt);
      });
      // A job may carry a value nobody configured (typed on another device
      // before this list synced) - keep it selectable rather than silently
      // switching the job to the first option.
      const keepable = current && current !== ADD_NEW_VALUE && current !== 'ALL';
      if (keepable && !values.includes(current)) {
        const opt = document.createElement('option');
        opt.value = current;
        opt.textContent = current;
        selectEl.appendChild(opt);
      }
      if (options.addNew) {
        const add = document.createElement('option');
        add.value = ADD_NEW_VALUE;
        add.textContent = '➕ เพิ่มฝ่ายย่อยใหม่...';
        selectEl.appendChild(add);
      }
      if (keepable || current === 'ALL') selectEl.value = current;
      if (!selectEl.value) selectEl.value = options.allLabel ? 'ALL' : (values[0] || '');
    }

    function fillDatalist(id, values) {
      const dl = document.getElementById(id);
      if (!dl) return;
      dl.innerHTML = '';
      values.forEach(v => {
        const opt = document.createElement('option');
        opt.value = v;
        dl.appendChild(opt);
      });
    }

    function renderOptionLists() {
      const subDepts = getOptionList('sub_departments');
      const filterEl = document.getElementById('subDeptFilter');
      const filterBefore = filterEl ? filterEl.value : 'ALL';
      fillSelect(filterEl, subDepts, filterBefore, { allLabel: 'ทุกฝ่ายย่อย (All Sub-Depts)', numbered: true });
      if (filterEl && filterEl.value !== filterBefore) renderApp();

      fillSelect(document.getElementById('newSubDept'), subDepts, undefined, { numbered: true, addNew: true });
      fillSelect(document.getElementById('editSubDept'), subDepts, undefined, { addNew: true });

      fillDatalist('technicianOptions', getOptionList('technicians'));
      fillDatalist('customerOptions', getOptionList('customers'));

      Object.keys(COMBOS).forEach(renderCombo);
    }

    /* -----------------------------------------------------------------------
     * COMBO FIELDS - a tidy numbered dropdown that still accepts a new name
     *
     * Customer and technician are picked from a <select> like the
     * sub-department. The last choice, "พิมพ์ชื่อใหม่...", reveals the text
     * box for a name the system has not seen yet. The text box is always the
     * field the save code reads, so the rest of the app is unchanged.
     * --------------------------------------------------------------------- */
    const CUSTOM_VALUE = '__custom__';
    const COMBOS = {
      newProjectName:  { select: 'newProjectNameSelect',  key: 'customers',   pick: 'เลือกลูกค้า / ไซต์งานจากรายการ...', custom: '✏️ พิมพ์ชื่อไซต์งาน / ลูกค้าใหม่ (ไม่มีในรายการ)...', onPick: v => autofillFromCustomer('new', v) },
      editProjectName: { select: 'editProjectNameSelect', key: 'customers',   pick: 'เลือกลูกค้า / ไซต์งานจากรายการ...', custom: '✏️ พิมพ์ชื่อไซต์งาน / ลูกค้าใหม่ (ไม่มีในรายการ)...', onPick: v => autofillFromCustomer('edit', v) },
      newTechnician:   { select: 'newTechnicianSelect',   key: 'technicians', pick: 'เลือกหัวหน้าช่างจากรายชื่อ...',       custom: '✏️ พิมพ์ชื่อช่างใหม่ (ไม่มีในรายชื่อ)...' },
      editTechnician:  { select: 'editTechnicianSelect',  key: 'technicians', pick: 'เลือกหัวหน้าช่างจากรายชื่อ...',       custom: '✏️ พิมพ์ชื่อช่างใหม่ (ไม่มีในรายชื่อ)...' }
    };

    function renderCombo(inputId) {
      const cfg = COMBOS[inputId];
      const select = document.getElementById(cfg.select);
      if (!select) return;
      const values = getOptionList(cfg.key);

      select.innerHTML = '';
      const ph = document.createElement('option');
      ph.value = '';
      ph.textContent = cfg.pick;
      select.appendChild(ph);
      values.forEach((v, i) => {
        const opt = document.createElement('option');
        opt.value = v;
        opt.textContent = (i + 1) + '. ' + v;
        select.appendChild(opt);
      });
      const custom = document.createElement('option');
      custom.value = CUSTOM_VALUE;
      custom.textContent = cfg.custom;
      select.appendChild(custom);

      syncCombo(inputId);
    }

    // Makes the dropdown reflect whatever the text box currently holds
    function syncCombo(inputId) {
      const cfg = COMBOS[inputId];
      const select = document.getElementById(cfg.select);
      const input = document.getElementById(inputId);
      if (!select || !input) return;
      const value = cleanOptionValue(input.value);

      if (value && Array.from(select.options).some(o => o.value === value)) {
        select.value = value;
        input.classList.add('hidden');
      } else if (value) {
        select.value = CUSTOM_VALUE;     // a name not on the list - keep it visible for editing
        input.classList.remove('hidden');
      } else {
        select.value = '';
        input.classList.add('hidden');
      }
    }

    function handleComboChange(inputId) {
      const cfg = COMBOS[inputId];
      const select = document.getElementById(cfg.select);
      const input = document.getElementById(inputId);
      if (!select || !input) return;

      if (select.value === CUSTOM_VALUE) {
        input.value = '';
        input.classList.remove('hidden');
        input.focus();
        return;
      }

      input.value = select.value;
      input.classList.add('hidden');
      if (select.value && cfg.onPick) cfg.onPick(select.value);
    }

    // The value the save code should use, with a friendly message when missing
    function comboValueOrWarn(inputId, label) {
      const input = document.getElementById(inputId);
      const value = cleanOptionValue(input ? input.value : '');
      if (value) return value;
      const cfg = COMBOS[inputId];
      const select = document.getElementById(cfg.select);
      alert('กรุณาเลือกหรือพิมพ์' + label);
      if (select && select.value === CUSTOM_VALUE && input) input.focus();
      else if (select) select.focus();
      return '';
    }

    // "➕ เพิ่มฝ่ายย่อยใหม่..." picked straight from the form
    async function handleSubDeptChange(selectEl) {
      if (!selectEl || selectEl.value !== ADD_NEW_VALUE) return;
      const prev = selectEl.dataset.prev || getOptionList('sub_departments')[0] || '';
      const typed = prompt('ชื่อฝ่ายย่อยใหม่ (จะถูกเพิ่มเป็นตัวเลือกให้ทุกเครื่อง):');
      const name = cleanOptionValue(typed);
      if (!name) { selectEl.value = prev; return; }

      const ok = await addOptionValue('sub_departments', name);
      renderOptionLists();
      selectEl.value = name;
      if (ok) showSyncToast('เพิ่มฝ่ายย่อย "' + name + '" แล้ว');
    }

    async function addOptionValue(key, value, quiet) {
      const name = cleanOptionValue(value);
      if (!name) return false;
      // Seed from the effective list so the first custom entry does not wipe the defaults
      const list = getOptionList(key);
      if (list.includes(name)) return true;
      const next = {};
      next[key] = list.concat([name]);
      return persistOptions(next, quiet);
    }

    /**
     * Saves the given lists ({key: [...]}) - to the sheet when Code.gs supports
     * it, otherwise to this browser. Returns true on success.
     */
    /**
     * `quiet` is for best-effort background updates - adding a newly typed
     * technician while saving a task, say. The task itself already saved, so
     * failing to also file the name must not raise an alarm about a failure
     * the user cannot act on.
     */
    async function persistOptions(partial, quiet) {
      const keys = Object.keys(partial).filter(k => OPTION_GROUPS[k]);
      if (keys.length === 0) return false;

      if (state.optionsSource === 'server') {
        const settings = {};
        keys.forEach(k => { settings[k] = partial[k].map(cleanOptionValue).filter(Boolean); });
        const { json, error } = await apiPost({ action: 'SAVE_SETTINGS', settings: settings });
        if (json && json.status === 'success' && json.settings) {
          Object.keys(OPTION_GROUPS).forEach(k => {
            state.options[k] = Array.isArray(json.settings[k]) ? json.settings[k] : [];
          });
          return true;
        }
        if (!quiet) alert('บันทึกตัวเลือกลง Google Sheet ไม่สำเร็จ\n\n' + (json && json.message ? json.message : apiErrorText(error)));
        return false;
      }

      keys.forEach(k => { state.options[k] = partial[k].map(cleanOptionValue).filter(Boolean); });
      saveLocalOptions();
      return true;
    }

    /* ---------------- ⚙ dialog ---------------- */
    let optionsDraft = null;

    function openOptionsModal(focusKey) {
      optionsDraft = {};
      Object.keys(OPTION_GROUPS).forEach(k => { optionsDraft[k] = getOptionList(k); });

      const note = document.getElementById('optionsSourceNote');
      if (note) {
        if (state.optionsSource === 'server') {
          note.className = 'text-[11px] rounded-xl px-3 py-2 border bg-emerald-50 border-emerald-200 text-emerald-800';
          note.innerHTML = '<i class="fa-solid fa-cloud-check mr-1"></i>ตัวเลือกเก็บใน Google Sheet (แท็บ <b>App_Settings</b>) — บันทึกแล้วทุกเครื่องเห็นเหมือนกัน';
        } else {
          note.className = 'text-[11px] rounded-xl px-3 py-2 border bg-amber-50 border-amber-300 text-amber-900';
          note.innerHTML = '<i class="fa-solid fa-triangle-exclamation mr-1"></i>ตอนนี้ตัวเลือกถูกจำไว้ใน<b>เครื่องนี้เท่านั้น</b> (สคริปต์ฝั่ง Google ยังเป็นเวอร์ชันเก่า) — ค่าที่ใช้ในใบงานที่บันทึกแล้วจะยังแสดงให้ทุกเครื่อง ถ้าต้องการซิงก์รายการทั้งหมด ให้วาง Code.gs ใหม่แล้ว Deploy เวอร์ชันใหม่';
        }
      }

      renderOptionsSections();
      document.getElementById('optionsModal').classList.remove('hidden');

      if (focusKey) {
        setTimeout(() => {
          const input = document.getElementById('optionInput_' + focusKey);
          if (input) { input.scrollIntoView({ block: 'center', behavior: 'smooth' }); input.focus(); }
        }, 80);
      }
    }

    function closeOptionsModal() {
      optionsDraft = null;
      document.getElementById('optionsModal').classList.add('hidden');
    }

    function renderOptionsSections() {
      const host = document.getElementById('optionsSections');
      if (!host || !optionsDraft) return;
      host.innerHTML = '';

      Object.keys(OPTION_GROUPS).forEach(key => {
        const group = OPTION_GROUPS[key];
        const inUse = valuesInUse(group.field, group.from);
        const list = optionsDraft[key];

        const section = document.createElement('div');
        section.className = 'rounded-2xl border border-slate-200 bg-white shadow-sm overflow-hidden';
        section.innerHTML =
          '<div class="px-4 py-3 bg-slate-50 border-b border-slate-200 flex items-center justify-between gap-2">' +
            '<div class="min-w-0">' +
              '<h4 class="text-sm font-bold text-slate-800 flex items-center gap-2"><i class="fa-solid ' + group.icon + ' text-pts-800"></i>' + esc(group.label) +
              ' <span class="text-[11px] font-bold text-slate-500 bg-white border border-slate-200 rounded-md px-1.5 py-0.5">' + list.length + '</span></h4>' +
              '<p class="text-[11px] text-slate-500">' + esc(group.hint) + '</p>' +
            '</div>' +
          '</div>' +
          '<div class="p-3 sm:p-4 space-y-3">' +
            '<div class="flex items-center gap-2">' +
              '<input type="text" id="optionInput_' + key + '" placeholder="' + esc(group.placeholder) + '" class="flex-1 px-3.5 py-2.5 bg-slate-50 border border-slate-300 rounded-xl text-sm focus:outline-none focus:ring-2 focus:ring-pts-800">' +
              '<button type="button" data-add="' + key + '" class="px-4 py-2.5 rounded-xl bg-pts-800 hover:bg-pts-700 text-white text-xs font-bold transition active:scale-95 shrink-0"><i class="fa-solid fa-plus mr-1"></i>เพิ่ม</button>' +
            '</div>' +
            '<div class="flex flex-wrap gap-2" data-chips="' + key + '"></div>' +
          '</div>';

        const chips = section.querySelector('[data-chips]');
        if (list.length === 0) {
          chips.innerHTML = '<p class="text-xs text-slate-400 italic">ยังไม่มีรายการ — พิมพ์ชื่อด้านบนแล้วกดเพิ่ม</p>';
        }
        list.forEach((value, idx) => {
          const count = inUse[value] || 0;
          const chip = document.createElement('span');
          chip.className = 'inline-flex items-center gap-1.5 pl-3 pr-1 py-1 rounded-full bg-slate-100 border border-slate-200 text-xs font-semibold text-slate-800 max-w-full';
          chip.innerHTML =
            '<span class="truncate max-w-[220px]" title="' + esc(value) + '">' + esc(value) + '</span>' +
            (count > 0 ? '<span class="text-[10px] font-bold text-slate-500 bg-white border border-slate-200 rounded-full px-1.5" title="มีใบงานใช้ค่านี้อยู่ ' + count + ' ใบ">' + count + '</span>' : '') +
            '<button type="button" data-remove="' + idx + '" class="w-6 h-6 rounded-full hover:bg-rose-100 text-rose-600 flex items-center justify-center transition" title="เอาออกจากรายการ"><i class="fa-solid fa-xmark text-[11px]"></i></button>';
          chip.querySelector('[data-remove]').onclick = () => {
            if (count > 0 && !confirm('"' + value + '" มีใบงานใช้อยู่ ' + count + ' ใบ\n\nถ้าเอาออก ค่านี้จะยังแสดงในตัวเลือกจนกว่าใบงานเหล่านั้นจะเปลี่ยนค่า แต่จะไม่แนะนำสำหรับงานใหม่\n\nเอาออกใช่หรือไม่?')) return;
            optionsDraft[key].splice(idx, 1);
            renderOptionsSections();
          };
          chips.appendChild(chip);
        });

        const input = section.querySelector('#optionInput_' + key);
        const addBtn = section.querySelector('[data-add]');
        const addFromInput = () => {
          const name = cleanOptionValue(input.value);
          if (!name) return;
          if (optionsDraft[key].some(v => v.toLowerCase() === name.toLowerCase())) {
            showSyncToast('"' + name + '" มีอยู่ในรายการแล้ว');
            input.value = '';
            return;
          }
          optionsDraft[key].push(name);
          input.value = '';
          renderOptionsSections();
          setTimeout(() => { const el = document.getElementById('optionInput_' + key); if (el) el.focus(); }, 0);
        };
        addBtn.onclick = addFromInput;
        input.addEventListener('keydown', e => { if (e.key === 'Enter') { e.preventDefault(); addFromInput(); } });

        host.appendChild(section);
      });
    }

    async function saveOptionsModal() {
      if (!optionsDraft) return;
      const payload = {};
      Object.keys(OPTION_GROUPS).forEach(k => { payload[k] = optionsDraft[k].slice(); });

      const ok = await withBusy('optionsSaveBtn', 'กำลังบันทึก...', () => persistOptions(payload));
      if (!ok) return;

      renderOptionLists();
      closeOptionsModal();
      showSyncToast(state.optionsSource === 'server' ? 'บันทึกตัวเลือกลง Google Sheet แล้ว (ทุกเครื่องเห็นเหมือนกัน)' : 'บันทึกตัวเลือกในเครื่องนี้แล้ว');
    }

    /* =======================================================================
     * CUSTOMER RECALL - reuse what a previous job already recorded
     *
     * The jobs on the sheet ARE the customer directory. Picking a customer
     * (from the address-book dialog, or by choosing a suggestion in the site
     * name box) copies the site location, contact phone, map link, technician
     * and sub-department from that customer's most recent job. Only EMPTY
     * fields are filled, so nothing the user already typed is overwritten.
     * ===================================================================== */
    const AUTOFILL_FIELDS = [
      { field: 'Site_Location',        idNew: 'newSiteLocation',     idEdit: 'editSiteLocation',     label: 'สถานที่' },
      { field: 'Site_Contact_Phone',   idNew: 'newSiteContactPhone', idEdit: 'editSiteContactPhone', label: 'เบอร์ติดต่อ' },
      { field: 'Site_Map_Url',         idNew: 'newSiteMapUrl',       idEdit: 'editSiteMapUrl',       label: 'แผนที่' },
      { field: 'Technician_In_Charge', idNew: 'newTechnician',       idEdit: 'editTechnician',       label: 'ช่าง' },
      { field: 'Sub_Department',       idNew: 'newSubDept',          idEdit: 'editSubDept',          label: 'ฝ่ายย่อย', isSelect: true }
    ];

    function jobSortKey(task) {
      const m = /(\d+)\s*$/.exec(String(task.Job_ID || '').trim());
      return m ? parseInt(m[1], 10) : 0;
    }

    // One entry per customer name: details come from the newest job
    function buildCustomerDirectory() {
      const byName = {};
      (state.tasks || []).forEach(t => {
        const name = cleanOptionValue(t.Project_Name);
        if (!name) return;
        const k = name.toLowerCase();
        if (!byName[k]) byName[k] = { name: name, jobs: 0, latest: null };
        byName[k].jobs++;
        if (!byName[k].latest || jobSortKey(t) > jobSortKey(byName[k].latest)) byName[k].latest = t;
      });
      // Names configured in ⚙ but without a job yet still appear (no details)
      (state.options.customers || []).forEach(name => {
        const k = name.toLowerCase();
        if (!byName[k]) byName[k] = { name: name, jobs: 0, latest: null };
      });
      return Object.values(byName).sort((a, b) => {
        if (b.jobs !== a.jobs) return b.jobs - a.jobs;
        return a.name.localeCompare(b.name, 'th');
      });
    }

    function findCustomer(name) {
      const k = cleanOptionValue(name).toLowerCase();
      if (!k) return null;
      return buildCustomerDirectory().find(c => c.name.toLowerCase() === k) || null;
    }

    // Copies details from `source` into the empty fields of the form. Returns
    // the labels that were filled so the caller can tell the user.
    function applyCustomerToForm(mode, source) {
      if (!source) return [];
      const filled = [];
      AUTOFILL_FIELDS.forEach(f => {
        const el = document.getElementById(mode === 'new' ? f.idNew : f.idEdit);
        if (!el) return;
        const value = cleanOptionValue(source[f.field]);
        if (!value) return;
        if (f.isSelect) {
          // Only for a NEW job, and only while the dropdown still sits on its
          // default first choice - never re-file an existing job's department
          if (mode !== 'new') return;
          const untouched = el.selectedIndex <= 0;
          if (!untouched || el.value === value) return;
          if (!Array.from(el.options).some(o => o.value === value)) {
            const opt = document.createElement('option');
            opt.value = value; opt.textContent = value;
            const addOpt = Array.from(el.options).find(o => o.value === ADD_NEW_VALUE) || null;
            el.insertBefore(opt, addOpt);
          }
          el.value = value;
          filled.push(f.label);
          return;
        }
        if (cleanOptionValue(el.value)) return;   // never overwrite what was typed
        el.value = value;
        if (COMBOS[el.id]) syncCombo(el.id);
        filled.push(f.label);
      });
      return filled;
    }

    function showAutofillHint(mode, text) {
      const el = document.getElementById(mode === 'new' ? 'newAutofillHint' : 'editAutofillHint');
      if (!el) return;
      if (!text) { el.classList.add('hidden'); el.textContent = ''; return; }
      el.innerHTML = '<i class="fa-solid fa-wand-magic-sparkles mr-1"></i>' + esc(text);
      el.classList.remove('hidden');
    }

    // Fired when the site-name box changes (typing a known name or picking a suggestion)
    function autofillFromCustomer(mode, name) {
      const customer = findCustomer(name);
      if (!customer || !customer.latest) { showAutofillHint(mode, ''); return; }
      const filled = applyCustomerToForm(mode, customer.latest);
      if (filled.length === 0) return;
      showAutofillHint(mode, 'เติม ' + filled.join(' · ') + ' จากใบงาน ' + customer.latest.Job_ID + ' ให้แล้ว (แก้ไขได้ตามต้องการ)');
    }

    /* ---------------- address-book dialog ---------------- */
    let customerPickerMode = 'new';

    function openCustomerPicker(mode) {
      customerPickerMode = mode;
      const search = document.getElementById('customerPickerSearch');
      if (search) search.value = '';
      renderCustomerPicker();
      document.getElementById('customerPickerModal').classList.remove('hidden');
      setTimeout(() => { if (search && !isMobileViewport()) search.focus(); }, 60);
    }

    function closeCustomerPicker() {
      document.getElementById('customerPickerModal').classList.add('hidden');
    }

    function renderCustomerPicker() {
      const host = document.getElementById('customerPickerList');
      if (!host) return;
      const searchEl = document.getElementById('customerPickerSearch');
      const q = cleanOptionValue(searchEl ? searchEl.value : '').toLowerCase();

      const rows = buildCustomerDirectory().filter(c => {
        if (!q) return true;
        const t = c.latest || {};
        return [c.name, t.Site_Location, t.Site_Contact_Phone, t.Technician_In_Charge, t.Sub_Department]
          .some(v => String(v || '').toLowerCase().includes(q));
      });

      host.innerHTML = '';
      if (rows.length === 0) {
        host.innerHTML =
          '<div class="py-10 text-center text-slate-400 bg-white rounded-2xl border border-dashed border-slate-300">' +
            '<i class="fa-solid fa-address-book text-3xl text-slate-300 mb-2 block"></i>' +
            '<p class="text-xs font-bold text-slate-600">' + (q ? 'ไม่พบลูกค้าที่ตรงกับคำค้น' : 'ยังไม่มีข้อมูลลูกค้าเก่า') + '</p>' +
            '<p class="text-[11px] text-slate-500 mt-1">ระบบดึงรายชื่อจากใบงานที่เคยบันทึกไว้ทั้งหมดโดยอัตโนมัติ</p>' +
          '</div>';
        return;
      }

      rows.slice(0, 200).forEach(c => {
        const t = c.latest || {};
        const line = (icon, color, text) => text ? '<span><i class="fa-solid ' + icon + ' ' + color + ' mr-1"></i>' + esc(text) + '</span>' : '';
        const card = document.createElement('button');
        card.type = 'button';
        card.className = 'w-full text-left bg-white rounded-2xl border border-slate-200 hover:border-pts-800 hover:shadow-md active:scale-[0.99] transition p-3.5';
        card.innerHTML =
          '<div class="flex items-start justify-between gap-3">' +
            '<div class="min-w-0">' +
              '<p class="text-sm font-bold text-slate-900 truncate">' + esc(c.name) + '</p>' +
              '<div class="mt-1 flex flex-wrap gap-x-3 gap-y-0.5 text-[11px] text-slate-600">' +
                line('fa-location-dot', 'text-rose-500', t.Site_Location) +
                line('fa-phone', 'text-emerald-600', t.Site_Contact_Phone) +
                line('fa-helmet-safety', 'text-amber-600', t.Technician_In_Charge) +
                line('fa-sitemap', 'text-pts-800', t.Sub_Department) +
                (t.Site_Map_Url ? '<span><i class="fa-solid fa-map text-blue-600 mr-1"></i>มีลิงก์แผนที่</span>' : '') +
                (!c.latest ? '<span class="italic text-slate-400">ยังไม่มีใบงาน — จะเติมเฉพาะชื่อ</span>' : '') +
              '</div>' +
            '</div>' +
            '<div class="text-right shrink-0">' +
              '<span class="inline-block text-[10px] font-bold text-slate-600 bg-slate-100 border border-slate-200 rounded-md px-1.5 py-0.5">' + c.jobs + ' ใบงาน</span>' +
              (t.Job_ID ? '<p class="text-[10px] font-mono text-slate-400 mt-1">ล่าสุด ' + esc(t.Job_ID) + '</p>' : '') +
            '</div>' +
          '</div>';
        card.onclick = () => pickCustomer(c);
        host.appendChild(card);
      });
    }

    function pickCustomer(customer) {
      const mode = customerPickerMode;
      const nameEl = document.getElementById(mode === 'new' ? 'newProjectName' : 'editProjectName');
      if (nameEl) { nameEl.value = customer.name; syncCombo(nameEl.id); }

      const filled = customer.latest ? applyCustomerToForm(mode, customer.latest) : [];
      closeCustomerPicker();

      if (filled.length > 0) {
        showAutofillHint(mode, 'เติม ' + filled.join(' · ') + ' จากใบงาน ' + customer.latest.Job_ID + ' ให้แล้ว (แก้ไขได้ตามต้องการ)');
      } else if (customer.latest) {
        showAutofillHint(mode, 'ช่องอื่นมีข้อมูลอยู่แล้ว จึงเติมเฉพาะชื่อลูกค้า (ใบงานล่าสุด ' + customer.latest.Job_ID + ')');
      } else {
        showAutofillHint(mode, '');
      }
    }

    /* =======================================================================
     * PHOTO DOWNLOAD - real image files, saved where the user wants them
     *
     * No .zip anywhere: every photo lands as its own .jpg, because a phone
     * cannot open a zip into its gallery and a technician wants the pictures
     * filed straight into a folder of their choosing.
     *
     * Three routes, picked by what the browser actually offers:
     *
     *  1. FOLDER PICKER (Chrome/Edge on Windows, macOS, Linux) - the user
     *     chooses any folder and the files are written into it directly,
     *     with no trip through the Downloads folder.
     *  2. OS SHARE SHEET (phones) - navigator.share hands the files to the
     *     system, so the user picks Photos, Files (any folder), Drive, LINE...
     *  3. PLAIN DOWNLOADS - the universal fallback, one save per photo.
     *
     * Bytes come from lh3.googleusercontent.com: the stored Drive links and
     * the thumbnail endpoint both refuse cross-origin reads, while the lh3
     * mirror of the same file id sends CORS headers. "=s0" is the stored
     * original, which the app already compressed to 900px before upload.
     * ===================================================================== */
    const PHOTO_FETCH_RETRIES = 2;
    const SEQUENTIAL_SAVE_DELAY_MS = 220;   // browsers drop saves fired back to back
    // An album is normally a few MB. This ceiling only ever trips on photos
    // that reached Drive without the app's compressor, and it stops a phone
    // from being pushed into an out-of-memory reload halfway through.
    const MEMORY_ROUTE_MAX_BYTES = 150 * 1024 * 1024;

    let photoSelectMode = false;
    let selectedPhotoIdx = new Set();
    let downloadToken = null;
    // Set the moment a download is requested. downloadToken only appears after
    // the folder dialog closes, so without this a second click while that
    // dialog is open would start a duplicate run.
    let downloadStarting = false;
    let pendingSaveFiles = [];              // fetched, still waiting for a destination
    let shareChunkSize = 0;                 // 0 = this device cannot share files
    let lastFetchReport = { failed: [], total: 0 };

    function supportsDirectoryPicker() {
      return typeof window.showDirectoryPicker === 'function';
    }

    /**
     * Largest batch this device will accept through the OS share sheet.
     * iOS in particular refuses large file sets, so probe downwards instead of
     * finding out only when the sheet fails to open.
     */
    function maxShareChunk(files) {
      if (!navigator.share || !navigator.canShare || files.length === 0) return 0;
      const sizes = [files.length, 20, 10, 5, 2, 1];
      for (const size of sizes) {
        if (size > files.length) continue;
        try {
          if (navigator.canShare({ files: files.slice(0, size) })) return size;
        } catch (err) { /* keep probing smaller */ }
      }
      return 0;
    }

    // Phones on a weak signal choke on four parallel image fetches
    function fetchConcurrency() {
      const conn = navigator.connection || navigator.mozConnection || navigator.webkitConnection;
      if (!conn) return 4;
      if (conn.saveData) return 2;
      const type = conn.effectiveType || '';
      if (type === 'slow-2g' || type === '2g') return 1;
      if (type === '3g') return 2;
      return 4;
    }

    function abortError() {
      const err = new Error('aborted');
      err.name = 'AbortError';
      return err;
    }

    /* ---------------- selection mode ---------------- */

    function setPhotoSelectMode(on) {
      photoSelectMode = !!on;
      if (!photoSelectMode) selectedPhotoIdx.clear();

      const bar = document.getElementById('detailPhotoSelectBar');
      if (bar) {
        bar.classList.toggle('hidden', !photoSelectMode);
        bar.classList.toggle('flex', photoSelectMode);
      }
      const btn = document.getElementById('detailSelectModeBtn');
      if (btn) {
        btn.innerHTML = photoSelectMode
          ? '<i class="fa-solid fa-xmark mr-1"></i>ออกจากโหมดเลือก'
          : '<i class="fa-regular fa-square-check mr-1"></i>เลือกรูป';
        btn.classList.toggle('bg-pts-50', photoSelectMode);
        btn.classList.toggle('border-pts-300', photoSelectMode);
      }
      updateSelectedPhotoCount();
      renderDetailPhotoThumbs();
    }

    function togglePhotoSelectMode() {
      setPhotoSelectMode(!photoSelectMode);
    }

    function togglePhotoSelected(idx) {
      if (selectedPhotoIdx.has(idx)) selectedPhotoIdx.delete(idx);
      else selectedPhotoIdx.add(idx);
      updateSelectedPhotoCount();
      renderDetailPhotoThumbs();
    }

    function selectAllPhotos(on) {
      selectedPhotoIdx.clear();
      if (on) (state.detailPhotos || []).forEach((_, i) => selectedPhotoIdx.add(i));
      updateSelectedPhotoCount();
      renderDetailPhotoThumbs();
    }

    function updateSelectedPhotoCount() {
      const n = selectedPhotoIdx.size;
      const a = document.getElementById('detailPhotoSelectedCount');
      const b = document.getElementById('detailDownloadSelectedCount');
      if (a) a.textContent = n;
      if (b) b.textContent = n;
      const btn = document.getElementById('detailDownloadSelectedBtn');
      if (btn) btn.disabled = n === 0;
    }

    /* ---------------- fetching the bytes ---------------- */

    // Candidate URLs to try, best quality first
    function photoDownloadUrls(entry) {
      const raw = typeof entry === 'string' ? entry : (entry && (entry.url || entry.dataUrl)) || '';
      if (!raw) return [];
      if (raw.indexOf('data:') === 0) return [raw];

      const id = getDriveFileId(raw);
      if (id) {
        return [
          'https://lh3.googleusercontent.com/d/' + id + '=s0',
          'https://lh3.googleusercontent.com/d/' + id + '=w1600'
        ];
      }
      return [raw];
    }

    /**
     * A photo, or null when it is genuinely unavailable.
     *
     * Retrying matters here: an 80-photo album on mobile data will hit a
     * dropped connection or a 429 from Google sooner or later, and one blip
     * used to cost the user that photo. A 404 or a non-image answer means the
     * file is deleted or unshared, so those fail fast instead of waiting out
     * the backoff.
     */
    /**
     * Shared health counter for one batch. A browser cannot tell a CORS refusal
     * apart from a dropped connection - both surface as "Failed to fetch" - so
     * a deleted Drive folder would otherwise make every single photo sit
     * through the full backoff. Once several photos in a row have failed with
     * nothing succeeding, retrying is clearly pointless and is switched off for
     * the rest of the batch.
     */
    function makeFetchBudget() {
      return { ok: 0, fail: 0, retries: PHOTO_FETCH_RETRIES, giveUp: false };
    }

    function noteFetchResult(budget, ok) {
      if (!budget) return;
      if (ok) {
        // Any success clears the streak: a patchy signal that drops two photos
        // and then recovers must keep its retries for the rest of the album.
        budget.ok++;
        budget.fail = 0;
        budget.retries = PHOTO_FETCH_RETRIES;
        return;
      }
      // Nothing coming back: stop retrying, then stop entirely rather than
      // grinding through a whole album that is clearly unreachable.
      budget.fail++;
      if (budget.fail >= 3) budget.retries = 0;
      if (budget.fail >= 5 && budget.ok === 0) budget.giveUp = true;
    }

    function albumUnreachableMessage() {
      // Throttling is listed first on purpose: it is by far the most common
      // cause after a few albums in a row, it looks identical to a permission
      // problem from the browser's side, and it clears itself in a minute.
      return 'ดึงรูปของใบงานนี้ไม่สำเร็จเลยสักรูป\n\n' +
        '1) ถ้าเพิ่งดาวน์โหลดรูปไปหลายชุดติดกัน Google จะจำกัดการโหลดชั่วคราว — ' +
        'พักสัก 1-2 นาทีแล้วกดใหม่ มักหายเอง\n' +
        '2) ถ้ายังไม่ได้เหมือนเดิม ให้ตรวจโฟลเดอร์ PTS_Engineering_Attachments ใน Google Drive ' +
        'ว่าถูกลบ ถูกย้าย หรือสิทธิ์แชร์ถูกเปลี่ยนเป็นส่วนตัวหรือไม่ครับ';
    }

    async function fetchPhotoBlob(entry, signal, budget) {
      const urls = photoDownloadUrls(entry);
      if (urls.length === 0) return null;
      const maxAttempts = budget ? budget.retries : PHOTO_FETCH_RETRIES;

      for (let attempt = 0; attempt <= maxAttempts; attempt++) {
        if (signal && signal.aborted) throw abortError();
        if (attempt > 0) await delay(600 * attempt);

        let worthRetrying = false;
        for (const url of urls) {
          if (signal && signal.aborted) throw abortError();
          try {
            const resp = await fetch(url, { signal, cache: 'no-store' });
            if (resp.ok) {
              const blob = await resp.blob();
              if (blob.size > 0 && /^image\//i.test(blob.type)) return blob;
              continue;   // an HTML error page - the file is gone, not a blip
            }
            if (resp.status === 429 || resp.status >= 500) worthRetrying = true;
          } catch (err) {
            if (err && err.name === 'AbortError') throw err;
            worthRetrying = true;   // network dropped mid-request
          }
        }
        if (!worthRetrying) return null;
      }
      return null;
    }

    function extForBlob(blob) {
      const type = (blob && blob.type) || '';
      if (/jpe?g/i.test(type)) return 'jpg';
      if (/png/i.test(type)) return 'png';
      if (/webp/i.test(type)) return 'webp';
      if (/gif/i.test(type)) return 'gif';
      return 'jpg';
    }

    // Windows/macOS/Android reject these in file names; Thai text is fine
    function sanitizeDownloadName(name) {
      return String(name || '').replace(/[\\/:*?"<>|]+/g, ' ').replace(/\s+/g, ' ').trim().slice(0, 80);
    }

    function photoFileName(jobId, idx, blob) {
      return `${jobId}_${String(idx + 1).padStart(2, '0')}.${extForBlob(blob)}`;
    }

    function saveBlob(blob, fileName) {
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = fileName;
      a.rel = 'noopener';
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      // Revoking immediately can cancel the save on slower devices
      setTimeout(() => URL.revokeObjectURL(url), 60000);
    }

    /* ---------------- progress dialog ---------------- */

    function showDownloadProgress(title, detail, percent, bytesText) {
      const note = document.getElementById('uploadProgressNote');
      if (note) note.textContent = 'ระบบกำลังดึงรูปจาก Google Drive ลงเครื่องของคุณ กรุณาอย่าเพิ่งปิดหน้าต่างเว็บ';
      const cancel = document.getElementById('uploadProgressCancelBtn');
      if (cancel) cancel.classList.remove('hidden');
      showUploadProgress(title, detail, percent, bytesText);
    }

    function hideDownloadProgress() {
      const note = document.getElementById('uploadProgressNote');
      if (note) note.textContent = 'ระบบกำลังส่งข้อมูลขึ้น Google Drive อย่างปลอดภัย กรุณารอสักครู่ อย่าเพิ่งปิดหน้าต่างเว็บจนกว่าจะเสร็จสิ้น';
      const cancel = document.getElementById('uploadProgressCancelBtn');
      if (cancel) cancel.classList.add('hidden');
      hideUploadProgress();
    }

    function cancelPhotoDownload() {
      if (downloadToken) downloadToken.controller.abort();
    }

    function reportFailedPhotos(failed, total) {
      if (!failed || failed.length === 0) return;
      alert(`บันทึกสำเร็จ ${total - failed.length} จาก ${total} รูป\n\n` +
        `โหลดไม่สำเร็จ: รูปที่ ${failed.join(', ')}\n\n` +
        'ลองกดดาวน์โหลดเฉพาะรูปที่ขาดอีกครั้งด้วยปุ่ม "เลือกรูป" — ส่วนใหญ่เป็นเพราะ Google ' +
        'จำกัดการโหลดชั่วคราวและจะได้ครบในรอบถัดไป ถ้ายังไม่ได้ แสดงว่าไฟล์นั้นถูกลบออกจาก Drive ครับ');
    }

    /* ---------------- the download itself ---------------- */

    async function downloadPhotos(scope) {
      if (downloadToken || downloadStarting) {
        alert('กำลังดาวน์โหลดชุดก่อนหน้าอยู่ กรุณารอให้เสร็จก่อน');
        return;
      }
      downloadStarting = true;
      try {
        await runPhotoDownload(scope);
      } finally {
        downloadStarting = false;
      }
    }

    async function runPhotoDownload(scope) {
      const photos = state.detailPhotos || [];
      const task = state.detailPhotoTask || {};

      let picked;
      if (scope === 'selected') {
        picked = Array.from(selectedPhotoIdx).sort((a, b) => a - b)
          .map(i => ({ idx: i, entry: photos[i] }))
          .filter(x => x.entry);
        if (picked.length === 0) {
          alert('ยังไม่ได้เลือกรูป — แตะที่รูปที่ต้องการก่อน แล้วค่อยกดดาวน์โหลด');
          return;
        }
      } else {
        picked = photos.map((entry, idx) => ({ idx, entry }));
      }
      if (picked.length === 0) return;

      const jobId = sanitizeDownloadName(task.Job_ID || 'PTS');

      // The folder picker needs the click to still be "fresh", so it runs
      // before any await. Anything else here would silently kill it.
      let dirHandle = null;
      if (supportsDirectoryPicker()) {
        try {
          dirHandle = await window.showDirectoryPicker({ id: 'pts-site-photos', mode: 'readwrite', startIn: 'pictures' });
        } catch (err) {
          if (err && err.name === 'AbortError') return;   // user closed the folder dialog
          dirHandle = null;                               // blocked by policy - use a fallback
        }
      }

      const controller = new AbortController();
      downloadToken = { controller };

      try {
        if (dirHandle) {
          await downloadIntoFolder(picked, jobId, dirHandle, controller);
        } else {
          await downloadThenChooseTarget(picked, jobId, controller);
        }
      } catch (err) {
        hideDownloadProgress();
        if (err && err.name === 'AbortError') {
          showSyncToast('ยกเลิกการดาวน์โหลดแล้ว');
        } else {
          alert('ดาวน์โหลดรูปไม่สำเร็จ\n\nอาจเป็นเพราะไฟล์ถูกลบออกจาก Google Drive หรือสิทธิ์แชร์ของโฟลเดอร์ถูกเปลี่ยน ' +
            'ลองกดเปิดดูรูปในใบงานเพื่อตรวจสอบ หรือเปิดไฟล์ต้นฉบับใน Google Drive โดยตรงครับ');
        }
      } finally {
        downloadToken = null;
      }
    }

    /**
     * Route 1 - write each photo straight into the folder the user chose.
     *
     * Photos are fetched a small batch at a time and then written in order.
     * Downloading one at a time made an 80-photo album take about three times
     * as long, while fetching the whole album first would put every byte in
     * memory - the batch keeps both the wait and the memory small.
     */
    async function downloadIntoFolder(picked, jobId, dirHandle, controller) {
      let done = 0;
      let bytes = 0;
      const failed = [];
      const budget = makeFetchBudget();
      const batchSize = Math.max(1, Math.min(fetchConcurrency(), picked.length));

      showDownloadProgress(`กำลังบันทึกลงโฟลเดอร์ ${dirHandle.name}...`, `0/${picked.length} รูป`, 0, '');

      for (let start = 0; start < picked.length; start += batchSize) {
        if (controller.signal.aborted || budget.giveUp) break;
        const batch = picked.slice(start, start + batchSize);

        let blobs;
        try {
          blobs = await Promise.all(batch.map(item => fetchPhotoBlob(item.entry, controller.signal, budget)));
        } catch (err) {
          if (err && err.name === 'AbortError') break;
          throw err;
        }
        blobs.forEach(blob => noteFetchResult(budget, !!blob));

        for (let i = 0; i < batch.length; i++) {
          const blob = blobs[i];
          if (blob) {
            try {
              const handle = await dirHandle.getFileHandle(photoFileName(jobId, batch[i].idx, blob), { create: true });
              const writable = await handle.createWritable();
              await writable.write(blob);
              await writable.close();
              bytes += blob.size;
            } catch (err) {
              // Folder went away, disk full, permission revoked mid-run
              hideDownloadProgress();
              alert(`เขียนไฟล์ลงโฟลเดอร์ "${dirHandle.name}" ไม่สำเร็จ (บันทึกไปแล้ว ${done - failed.length} รูป)\n\n` +
                'อาจเป็นเพราะโฟลเดอร์ถูกย้าย/ลบ พื้นที่ดิสก์เต็ม หรือสิทธิ์เขียนถูกยกเลิก กรุณาลองใหม่แล้วเลือกโฟลเดอร์อีกครั้งครับ');
              return;
            }
          } else {
            failed.push(batch[i].idx + 1);
          }
          done++;
        }

        showDownloadProgress(
          `กำลังบันทึกลงโฟลเดอร์ ${dirHandle.name}...`,
          `${done}/${picked.length} รูป`,
          (done / picked.length) * 100,
          formatBytes(bytes)
        );
      }

      hideDownloadProgress();

      if (controller.signal.aborted) {
        showSyncToast(`ยกเลิกแล้ว — บันทึกไปก่อนหน้า ${done - failed.length} รูป`);
        return;
      }
      const ok = done - failed.length;
      if (ok === 0) { alert(albumUnreachableMessage()); return; }
      reportFailedPhotos(failed, picked.length);
      if (failed.length === 0) {
        showSyncToast(`บันทึก ${ok} รูปลงโฟลเดอร์ "${dirHandle.name}" แล้ว (${formatBytes(bytes)})`);
      }
    }

    /**
     * Routes 2 and 3 - fetch everything first, then let the user say where it
     * goes. The destination is chosen by a fresh tap on the next dialog:
     * iOS in particular refuses navigator.share() if the tap that started the
     * work has already been "spent" on awaits.
     */
    async function downloadThenChooseTarget(picked, jobId, controller) {
      const files = [];
      const failed = [];
      let done = 0;
      let bytes = 0;
      let hitCeiling = false;
      const budget = makeFetchBudget();

      showDownloadProgress(`กำลังเตรียมรูป 0/${picked.length}...`, `รหัสงาน: ${jobId}`, 0, '');

      // A few at a time: parallel requests make Drive throttle and can stall a
      // phone, while one at a time is needlessly slow on a good connection.
      const queue = picked.slice();
      const worker = async () => {
        while (queue.length > 0) {
          if (controller.signal.aborted || hitCeiling || budget.giveUp) return;
          const item = queue.shift();
          const blob = await fetchPhotoBlob(item.entry, controller.signal, budget);
          noteFetchResult(budget, !!blob);
          if (blob) {
            if (bytes + blob.size > MEMORY_ROUTE_MAX_BYTES) { hitCeiling = true; return; }
            files.push({ order: item.idx, file: new File([blob], photoFileName(jobId, item.idx, blob), { type: blob.type }) });
            bytes += blob.size;
          } else {
            failed.push(item.idx + 1);
          }
          done++;
          showDownloadProgress(
            `กำลังเตรียมรูป ${done}/${picked.length}...`,
            `รหัสงาน: ${jobId}`,
            (done / picked.length) * 100,
            formatBytes(bytes)
          );
        }
      };
      await Promise.all(Array.from({ length: Math.min(fetchConcurrency(), picked.length) }, worker));

      hideDownloadProgress();
      if (controller.signal.aborted) { showSyncToast('ยกเลิกการดาวน์โหลดแล้ว'); return; }
      if (files.length === 0) { alert(albumUnreachableMessage()); return; }

      if (hitCeiling) {
        alert(`รูปชุดนี้มีขนาดรวมใหญ่มาก เครื่องอาจค้างถ้าโหลดทั้งหมดพร้อมกัน\n\n` +
          `ระบบเตรียมให้ ${files.length} รูปแรก (${formatBytes(bytes)}) ก่อน — ` +
          'บันทึกชุดนี้แล้วค่อยกลับมาเลือกรูปที่เหลือด้วยปุ่ม "เลือกรูป" อีกครั้งได้ครับ');
      }

      files.sort((a, b) => a.order - b.order);
      pendingSaveFiles = files.map(f => f.file);
      lastFetchReport = { failed: failed, total: picked.length };
      openSaveTargetModal(bytes);
    }

    /* ---------------- choosing where the files go ---------------- */

    function openSaveTargetModal(bytes) {
      const modal = document.getElementById('saveTargetModal');
      if (!modal) return;

      const summary = document.getElementById('saveTargetSummary');
      if (summary) {
        summary.textContent = `${pendingSaveFiles.length} รูป · ${formatBytes(bytes)}` +
          (lastFetchReport.failed.length > 0 ? ` (โหลดไม่สำเร็จ ${lastFetchReport.failed.length} รูป)` : '');
      }

      shareChunkSize = maxShareChunk(pendingSaveFiles);
      updateSaveTargetButtons();
      modal.classList.remove('hidden');
    }

    function updateSaveTargetButtons() {
      const shareBtn = document.getElementById('saveTargetShareBtn');
      const hint = document.getElementById('saveTargetHint');
      const shareable = shareChunkSize > 0;

      if (shareBtn) {
        shareBtn.classList.toggle('hidden', !shareable);
        if (shareable) {
          // Some phones (iOS especially) refuse a big file set, so the album is
          // handed over a chunk per tap - each tap is its own user gesture.
          const batches = Math.ceil(pendingSaveFiles.length / shareChunkSize);
          shareBtn.innerHTML = batches > 1
            ? `<i class="fa-solid fa-folder-open mr-1.5"></i>เลือกที่เก็บรูป (เหลือ ${pendingSaveFiles.length} รูป · อีก ${batches} ชุด)`
            : '<i class="fa-solid fa-folder-open mr-1.5"></i>เลือกที่เก็บรูป (คลังภาพ / ไฟล์ / Drive)';
        }
      }

      if (hint) {
        if (!shareable) {
          hint.textContent = 'ไฟล์จะถูกบันทึกลงโฟลเดอร์ดาวน์โหลดของเบราว์เซอร์ — ถ้าต้องการเลือกโฟลเดอร์เองทุกครั้ง ให้เปิด "ถามที่จัดเก็บไฟล์ก่อนดาวน์โหลด" ในการตั้งค่าเบราว์เซอร์';
        } else if (pendingSaveFiles.length > shareChunkSize) {
          hint.textContent = `เครื่องนี้รับได้ครั้งละ ${shareChunkSize} รูป กดปุ่มด้านบนซ้ำจนครบได้เลย — แต่ละครั้งเลือกที่เก็บได้อิสระ`;
        } else {
          hint.textContent = 'แนะนำ "เลือกที่เก็บรูป" — เลือกได้ทั้งคลังภาพ, แอปไฟล์ (เลือกโฟลเดอร์เองได้), Google Drive หรือส่งต่อทางแชต';
        }
      }
    }

    function closeSaveTargetModal() {
      const modal = document.getElementById('saveTargetModal');
      if (modal) modal.classList.add('hidden');
      pendingSaveFiles = [];
      shareChunkSize = 0;
    }

    async function saveFetchedPhotos(how) {
      if (pendingSaveFiles.length === 0) { closeSaveTargetModal(); return; }

      if (how === 'share' && shareChunkSize > 0) {
        const chunk = pendingSaveFiles.slice(0, shareChunkSize);
        try {
          await navigator.share({ files: chunk, title: 'รูปภาพหน้างาน' });
        } catch (err) {
          // Backing out of the OS sheet is a decision, not a fault
          if (err && (err.name === 'AbortError' || err.name === 'NotAllowedError')) return;
          // The sheet refused the files outright - stop offering it and save instead
          shareChunkSize = 0;
          updateSaveTargetButtons();
          alert('ระบบของเครื่องไม่รับไฟล์ชุดนี้ผ่านหน้าต่างแชร์\n\nระบบจะบันทึกลงโฟลเดอร์ดาวน์โหลดให้แทนครับ');
          const rest = pendingSaveFiles.slice();
          closeSaveTargetModal();
          await saveFilesSequentially(rest);
          reportFailedPhotos(lastFetchReport.failed, lastFetchReport.total);
          return;
        }

        pendingSaveFiles = pendingSaveFiles.slice(chunk.length);
        if (pendingSaveFiles.length === 0) {
          closeSaveTargetModal();
          showSyncToast('ส่งรูปให้เลือกที่เก็บครบแล้ว');
          reportFailedPhotos(lastFetchReport.failed, lastFetchReport.total);
        } else {
          updateSaveTargetButtons();
          showSyncToast(`ส่งแล้ว ${chunk.length} รูป — เหลืออีก ${pendingSaveFiles.length} รูป กดปุ่มเดิมต่อได้เลย`);
        }
        return;
      }

      const files = pendingSaveFiles.slice();
      closeSaveTargetModal();
      await saveFilesSequentially(files);
      reportFailedPhotos(lastFetchReport.failed, lastFetchReport.total);
    }

    // Route 3 - one save at a time, spaced out so no browser drops any of them
    async function saveFilesSequentially(files) {
      if (!files || files.length === 0) return;
      const controller = new AbortController();
      downloadToken = { controller };
      let saved = 0;
      try {
        for (let i = 0; i < files.length; i++) {
          if (controller.signal.aborted) break;
          if (files.length > 1) {
            showDownloadProgress(`กำลังบันทึก ${i + 1}/${files.length} รูป...`, files[i].name, ((i + 1) / files.length) * 100, '');
          }
          saveBlob(files[i], files[i].name);
          saved++;
          if (i < files.length - 1) await delay(SEQUENTIAL_SAVE_DELAY_MS);
        }
        showSyncToast(controller.signal.aborted
          ? `ยกเลิกแล้ว — บันทึกไปก่อนหน้า ${saved} รูป`
          : `บันทึก ${saved} รูปลงเครื่องแล้ว`);
      } catch (err) {
        alert(`บันทึกไฟล์ไม่สำเร็จ (บันทึกไปแล้ว ${saved} รูป)\n\nกรุณาลองใหม่อีกครั้งครับ`);
      } finally {
        downloadToken = null;
        hideDownloadProgress();
      }
    }

    // Save just the photo currently open in the lightbox
    async function downloadCurrentLightboxPhoto() {
      const entry = (state.lightboxPhotos || [])[state.lightboxIndex];
      if (!entry) return;
      const jobId = sanitizeDownloadName((state.lightboxLabel || 'PTS').split('·')[0]);
      const btn = document.getElementById('lightboxDownloadBtn');
      const original = btn ? btn.innerHTML : '';
      if (btn) { btn.disabled = true; btn.innerHTML = '<i class="fa-solid fa-circle-notch fa-spin mr-1.5"></i>กำลังดาวน์โหลด...'; }
      try {
        const blob = await fetchPhotoBlob(entry);
        if (!blob) throw new Error('fetch-failed');
        const name = photoFileName(jobId, state.lightboxIndex, blob);
        const file = new File([blob], name, { type: blob.type });

        // On a phone this opens the OS sheet so the photo can go to the
        // gallery or any folder; on a PC it saves straight away.
        if (maxShareChunk([file]) > 0) {
          try {
            await navigator.share({ files: [file], title: name });
            showSyncToast('ส่งรูปให้เลือกที่เก็บแล้ว');
            return;
          } catch (err) {
            if (err && (err.name === 'AbortError' || err.name === 'NotAllowedError')) return;
          }
        }
        saveBlob(blob, name);
        showSyncToast('ดาวน์โหลดรูปเรียบร้อย');
      } catch (err) {
        alert('ดาวน์โหลดรูปนี้ไม่สำเร็จ — ลองกด "เปิดไฟล์ต้นฉบับใน Google Drive" แล้วบันทึกจากหน้านั้นแทนครับ');
      } finally {
        if (btn) { btn.disabled = false; btn.innerHTML = original; }
      }
    }

    /* =======================================================================
     * TASKS / TO-DO LIST
     *
     * Deliberately mirrors the sales app (PTS Sales): Open/Done status,
     * yyyy-MM-dd dates with an optional HH:mm time, Low/Normal/High/Urgent
     * priority, a type, and the same three views - ของฉัน / ฉันมอบหมาย /
     * ทั้งทีม. Anyone who uses one app already knows this one.
     *
     * The sales app knows who you are from its login. This app has none, so
     * "who am I" is a name picked once and kept on the device. That is enough
     * to drive the three views and to fill in Created_By / Done_By; it is a
     * convenience, not a security boundary, and the whole team sees every task
     * either way.
     * ===================================================================== */
    const TASK_OPEN = 'Open';
    const TASK_DONE = 'Done';
    const TASK_PRIORITIES = ['Low', 'Normal', 'High', 'Urgent'];
    const DEFAULT_TASK_TYPES = ['ติดต่อลูกค้า', 'เข้าหน้างาน', 'ทำรายงาน', 'เบิก/สั่งอะไหล่', 'ส่งเอกสาร', 'ตรวจรับงาน', 'อื่นๆ'];
    const MY_NAME_KEY = 'pts_my_name';
    const TASK_SEEN_KEY = 'pts_task_seen';
    const TASK_NOTIFIED_KEY = 'pts_task_notified';

    state.jobTasks = [];
    let taskEditingId = null;
    let tasksFetched = false;
    // null = ยังไม่รู้ · true = หลังบ้านพร้อม · false = Apps Script ยังเป็นรุ่นเก่า
    let taskBackendReady = null;
    let taskScope = 'team';          // switched to 'mine' once a name is set
    let taskQuery = '';
    let taskShowDone = false;
    let taskSeenMap = null;

    /* ---------------- who am I ---------------- */

    function myName() {
      try { return cleanOptionValue(localStorage.getItem(MY_NAME_KEY) || ''); }
      catch (err) { return ''; }
    }

    function setMyName(name) {
      const clean = cleanOptionValue(name);
      const before = myName();
      try { localStorage.setItem(MY_NAME_KEY, clean); } catch (err) {}
      // The server routes notifications by name, so the device registration
      // has to follow a name change or alerts would go to the old person.
      if (clean && clean !== before && taskNotifyState() === 'granted') {
        enablePushOnThisDevice().then(updateNotifyButton);
      }
      // Without a name the personal views are meaningless - show the team view
      if (!clean && taskScope !== 'team') taskScope = 'team';
      if (clean && taskScope === 'team') taskScope = 'mine';
      renderMyNamePickers();
      refreshTaskViews();
    }

    function renderMyNamePickers() {
      const me = myName();
      const names = getOptionList('technicians');
      ['taskMeSelect', 'taskMeSelectMobile'].forEach(id => {
        const sel = document.getElementById(id);
        if (!sel) return;
        sel.innerHTML = '';
        const none = document.createElement('option');
        none.value = '';
        none.textContent = 'ยังไม่ได้เลือก';
        sel.appendChild(none);
        const list = me && !names.includes(me) ? names.concat([me]) : names;
        list.forEach(n => {
          const o = document.createElement('option');
          o.value = n;
          o.textContent = n;
          sel.appendChild(o);
        });
        sel.value = me;
      });
    }

    /* ---------------- date / time helpers ---------------- */

    function todayStr() {
      const d = new Date();
      const p = n => String(n).padStart(2, '0');
      return d.getFullYear() + '-' + p(d.getMonth() + 1) + '-' + p(d.getDate());
    }

    function nowHHMM() {
      const d = new Date();
      const p = n => String(n).padStart(2, '0');
      return p(d.getHours()) + ':' + p(d.getMinutes());
    }

    function addDaysStr(iso, n) {
      const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(String(iso || ''));
      if (!m) return '';
      const d = new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]) + n);
      const p = x => String(x).padStart(2, '0');
      return d.getFullYear() + '-' + p(d.getMonth() + 1) + '-' + p(d.getDate());
    }

    // Accepts yyyy-MM-dd (the stored shape) and dd/MM/yyyy (typed in the sheet)
    function taskDateOnly(value) {
      const s = String(value || '').replace(/^'/, '').trim();
      let m = /^(\d{4})-(\d{1,2})-(\d{1,2})$/.exec(s);
      if (m) return m[1] + '-' + String(m[2]).padStart(2, '0') + '-' + String(m[3]).padStart(2, '0');
      m = /^(\d{1,2})\/(\d{1,2})\/(\d{4})$/.exec(s);
      if (m) return m[3] + '-' + String(m[2]).padStart(2, '0') + '-' + String(m[1]).padStart(2, '0');
      return '';
    }

    function taskHHMM(value) {
      const m = /^(\d{1,2}):(\d{2})/.exec(String(value || '').replace(/^'/, '').trim());
      if (!m) return '';
      if (Number(m[1]) > 23 || Number(m[2]) > 59) return '';
      return String(m[1]).padStart(2, '0') + ':' + m[2];
    }

    function taskDateThai(iso) {
      const d = taskDateOnly(iso);
      return d ? d.split('-').reverse().join('/') : '';
    }

    function daysBetweenStr(fromIso, toIso) {
      const a = taskDateOnly(fromIso), b = taskDateOnly(toIso);
      if (!a || !b) return null;
      const pa = a.split('-').map(Number), pb = b.split('-').map(Number);
      return Math.round((new Date(pb[0], pb[1] - 1, pb[2]) - new Date(pa[0], pa[1] - 1, pa[2])) / 86400000);
    }

    const taskIsDone = t => String(t.Status || '').trim() === TASK_DONE;

    /**
     * How a task reads on screen - same buckets and wording as the sales app.
     * A task due today with a time that has passed counts as overdue, which is
     * the whole point of carrying a time at all.
     */
    function taskDue(t) {
      if (taskIsDone(t)) return { key: 'done', tone: 'green', label: 'เสร็จแล้ว' };
      const d = taskDateOnly(t.Due_Date), tm = taskHHMM(t.Due_Time), today = todayStr();
      const at = tm ? ' ' + tm + ' น.' : '';
      if (!d) return { key: 'later', tone: 'slate', label: 'ไม่ระบุวัน' };
      if (d < today) return { key: 'overdue', tone: 'red', label: 'เกินกำหนด ' + daysBetweenStr(d, today) + ' วัน' };
      if (d === today) {
        if (tm && nowHHMM() > tm) return { key: 'overdue', tone: 'red', label: 'เลยเวลา' + at };
        return { key: 'today', tone: 'amber', label: 'วันนี้' + at };
      }
      if (d === addDaysStr(today, 1)) return { key: 'soon', tone: 'blue', label: 'พรุ่งนี้' + at };
      return { key: 'later', tone: 'slate', label: taskDateThai(d) + at };
    }

    const priorityRank = p => {
      const i = TASK_PRIORITIES.indexOf(String(p || 'Normal'));
      return i === -1 ? 1 : (TASK_PRIORITIES.length - 1 - i);   // Urgent first
    };
    const taskSortKey = t => taskDateOnly(t.Due_Date) + ' ' + (taskHHMM(t.Due_Time) || '24:00');
    const byDue = (a, b) => taskSortKey(a).localeCompare(taskSortKey(b)) || priorityRank(a.Priority) - priorityRank(b.Priority);

    /* ---------------- unread dot ---------------- */

    const taskStamp = t => String(t.Updated_At || t.Created_At || '');

    function readTaskSeen() {
      if (taskSeenMap) return taskSeenMap;
      let raw = null;
      try { raw = localStorage.getItem(TASK_SEEN_KEY); } catch (err) {}
      if (raw) {
        try { taskSeenMap = JSON.parse(raw) || {}; } catch (err) { taskSeenMap = {}; }
        return taskSeenMap;
      }
      // First run on this device: treat everything already there as read, or
      // the whole list would light up at once and mean nothing.
      taskSeenMap = {};
      (state.jobTasks || []).forEach(t => { taskSeenMap[t.Task_ID] = taskStamp(t); });
      writeTaskSeen();
      return taskSeenMap;
    }

    function writeTaskSeen() {
      const live = {};
      (state.jobTasks || []).forEach(t => {
        if (taskSeenMap[t.Task_ID] !== undefined) live[t.Task_ID] = taskSeenMap[t.Task_ID];
      });
      taskSeenMap = live;   // drop ids that no longer exist
      try { localStorage.setItem(TASK_SEEN_KEY, JSON.stringify(taskSeenMap)); } catch (err) {}
    }

    function isTaskUnread(t) {
      const me = myName();
      if (taskIsDone(t) || !me || cleanOptionValue(t.Assignee) !== me) return false;
      return readTaskSeen()[t.Task_ID] !== taskStamp(t);
    }

    function markTaskRead(t) {
      readTaskSeen()[t.Task_ID] = taskStamp(t);
      writeTaskSeen();
    }

    /* ---------------- counts ---------------- */

    // Tasks that need attention now. With a name set this is "mine"; without
    // one it is the whole team, so the button still means something.
    function myTaskAlerts() {
      const me = myName();
      return (state.jobTasks || [])
        .filter(t => !taskIsDone(t) &&
          (!me || cleanOptionValue(t.Assignee) === me) &&
          ['overdue', 'today'].includes(taskDue(t).key))
        .sort(byDue);
    }

    function taskCounts() {
      const c = { late: 0, today: 0, soon: 0, open: 0, done: 0 };
      (state.jobTasks || []).forEach(t => {
        if (taskIsDone(t)) { c.done++; return; }
        c.open++;
        const k = taskDue(t).key;
        if (k === 'overdue') c.late++;
        else if (k === 'today') c.today++;
        else if (k === 'soon') c.soon++;
      });
      return c;
    }

    function scopeFilter(scope) {
      const me = myName();
      if (!me) return () => true;
      if (scope === 'mine') return t => cleanOptionValue(t.Assignee) === me;
      if (scope === 'byme') return t => cleanOptionValue(t.Created_By) === me && cleanOptionValue(t.Assignee) !== me;
      return () => true;
    }

    /* ---------------- server ---------------- */

    async function fetchJobTasks(options) {
      options = options || {};
      if (state.dbMode === 'mock' || !state.apiUrl) return false;

      const { json } = await apiGetDetailed({ tasks: '1' });
      // An old Code.gs answers ?tasks=1 with the plain job list, so a missing
      // `tasks` array is how we know the sheet side has not been updated yet.
      if (!json || json.status !== 'success' || !Array.isArray(json.tasks)) {
        if (json && json.status === 'success') {
          taskBackendReady = false;
          renderTaskBackendNotice();
        }
        return false;
      }

      taskBackendReady = true;
      const changed = JSON.stringify(state.jobTasks) !== JSON.stringify(json.tasks);
      state.jobTasks = json.tasks;
      tasksFetched = true;
      // The job table carries a Task column now, so it has to be redrawn when
      // the task list moves - otherwise a row shows a stale next action.
      if (changed) renderApp();
      refreshTaskViews();
      if (!options.silent) notifyDueTasks();
      return true;
    }

    function refreshTaskViews() {
      updateTaskAlertUi();
      if (!document.getElementById('tasksModal').classList.contains('hidden')) {
        renderTaskAssigneeFilter();
        renderScopeTabs();
        renderTasksList();
      }
    }

    async function saveJobTask(task) {
      const { json, error } = await apiPost({ action: 'TASK_SAVE', task: { ...task, Actor: myName() } });
      if (!json || json.status !== 'success') {
        // "Invalid action type" means the sheet side has no Task code yet -
        // a generic save error would send someone hunting the wrong problem.
        const msg = json && json.message ? String(json.message) : '';
        if (/invalid action/i.test(msg)) {
          taskBackendReady = false;
          renderTaskBackendNotice();
          alert('ยังบันทึก Task ไม่ได้ เพราะสคริปต์ฝั่ง Google Sheet ยังเป็นรุ่นเก่า\n\n' +
            'วิธีแก้ (ทำครั้งเดียว):\n' +
            '1. เปิดชีต > ส่วนขยาย > Apps Script\n' +
            '2. วางไฟล์ Code.gs ตัวใหม่ทับของเดิม แล้วกด Ctrl+S\n' +
            '3. Deploy > Manage deployments > ✏️ > Version: New version > Deploy');
          return false;
        }
        alert('บันทึก Task ไม่สำเร็จ\n\n' + (msg || apiErrorText(error)));
        return false;
      }
      await fetchJobTasks({ silent: true });
      markActivity();
      return true;
    }

    async function deleteJobTask(taskId) {
      const { json, error } = await apiPost({ action: 'TASK_DELETE', task_id: taskId });
      if (!json || json.status !== 'success') {
        alert('ลบ Task ไม่สำเร็จ\n\n' + (json && json.message ? json.message : apiErrorText(error)));
        return false;
      }
      await fetchJobTasks({ silent: true });
      markActivity();
      return true;
    }

    /* ---------------- header badge + banner ---------------- */

    function updateTaskAlertUi() {
      const alerts = myTaskAlerts();
      const urgent = alerts.length;
      const text = urgent > 99 ? '99+' : String(urgent);

      ['taskAlertBadge', 'taskBtnCount', 'taskFabCount'].forEach(id => {
        const el = document.getElementById(id);
        if (!el) return;
        el.textContent = text;
        el.classList.toggle('hidden', urgent === 0);
        el.classList.toggle('flex', urgent > 0);
      });

      const banner = document.getElementById('taskAlertBanner');
      if (!banner) return;
      if (urgent === 0) { banner.classList.add('hidden'); return; }

      const over = alerts.filter(t => taskDue(t).key === 'overdue').length;
      const late = over > 0;
      banner.className = 'rounded-2xl border p-3 sm:p-3.5 shadow-xs transition ' +
        (late ? 'bg-rose-50 border-rose-300 text-rose-900' : 'bg-amber-50 border-amber-300 text-amber-900');

      const icon = document.getElementById('taskAlertIcon');
      if (icon) {
        icon.className = 'w-9 h-9 rounded-xl flex items-center justify-center text-base shrink-0 ' +
          (late ? 'bg-rose-200 text-rose-800' : 'bg-amber-200 text-amber-900');
        icon.innerHTML = `<i class="fa-solid ${late ? 'fa-triangle-exclamation' : 'fa-bell'}"></i>`;
      }

      const me = myName();
      const parts = [];
      if (over > 0) parts.push(`เกินกำหนด ${over} รายการ`);
      if (urgent - over > 0) parts.push(`ครบกำหนดวันนี้ ${urgent - over} รายการ`);
      const titleEl = document.getElementById('taskAlertTitle');
      if (titleEl) titleEl.textContent = (me ? 'Task ของคุณ — ' : 'Task ของทีม — ') + parts.join(' · ');

      const detailEl = document.getElementById('taskAlertDetail');
      if (detailEl) {
        detailEl.textContent = alerts.slice(0, 3)
          .map(t => `${t.Title} (${taskDue(t).label})`).join(' · ');
      }
      banner.classList.remove('hidden');
    }

    /* ---------------- notification ---------------- */

    function taskNotifyState() {
      if (!('Notification' in window)) return 'unsupported';
      return Notification.permission;
    }

    function updateNotifyButton() {
      const btn = document.getElementById('taskNotifyBtn');
      if (!btn) return;
      const s = taskNotifyState();
      if (s === 'unsupported') { btn.classList.add('hidden'); return; }
      btn.classList.remove('hidden');

      // Whether the "test" button is useful depends on real push being live
      const testBtn = document.getElementById('taskTestPushBtn');
      const pushLive = !!(pushConfig && pushConfig.ready && pushConfig.registered);
      if (testBtn) testBtn.classList.toggle('hidden', !pushLive);

      if (s === 'granted') {
        btn.innerHTML = pushLive
          ? '<i class="fa-solid fa-bell mr-1.5 text-emerald-600"></i>แจ้งเตือนแม้ปิดแอป: เปิดอยู่'
          : '<i class="fa-solid fa-bell mr-1.5 text-amber-600"></i>แจ้งเตือนเฉพาะตอนเปิดแอป';
        btn.disabled = pushLive;
        btn.classList.toggle('opacity-70', pushLive);
        btn.classList.toggle('cursor-default', pushLive);
        return;
      }
      if (s === 'denied') {
        btn.innerHTML = '<i class="fa-solid fa-bell-slash mr-1.5 text-rose-500"></i>ถูกปิดไว้ในตั้งค่าเบราว์เซอร์';
        btn.disabled = true;
        btn.classList.add('opacity-70', 'cursor-default');
      } else {
        btn.innerHTML = '<i class="fa-regular fa-bell mr-1.5"></i>เปิดแจ้งเตือนบนเครื่องนี้';
        btn.disabled = false;
        btn.classList.remove('opacity-70', 'cursor-default');
      }
    }

    async function enableTaskNotifications() {
      if (!('Notification' in window)) {
        alert('เบราว์เซอร์นี้ไม่รองรับการแจ้งเตือน — ระบบจะยังเตือนผ่านตัวเลขบนปุ่ม Task และแถบสีด้านบนตามปกติครับ');
        return;
      }
      if (!myName()) {
        alert('กรุณาเลือกก่อนว่า "ฉันคือ" ใคร\n\nระบบจะได้ส่งเฉพาะงานของคุณมาที่เครื่องนี้ครับ');
        const sel = document.getElementById(isMobileViewport() ? 'taskMeSelectMobile' : 'taskMeSelect');
        if (sel) sel.focus();
        return;
      }

      const result = await Notification.requestPermission();
      updateNotifyButton();
      if (result === 'granted') {
        // Real push if the server is set up; otherwise in-app alerts only
        const res = await enablePushOnThisDevice();
        updateNotifyButton();
        if (res.ok) {
          showSyncToast('เปิดแจ้งเตือนแม้ปิดแอปแล้ว');
        } else if (res.reason === 'server') {
          showSyncToast('เปิดแจ้งเตือนแล้ว (เฉพาะตอนเปิดแอปค้างไว้)');
          alert('เปิดแจ้งเตือนบนเครื่องนี้แล้ว แต่ยังเตือนได้เฉพาะตอนเปิดแอปค้างไว้\n\n' +
            'ถ้าต้องการให้เตือนแม้ปิดแอป ต้องเปิด Apps Script ของชีตนี้ แล้วรันฟังก์ชัน setupPush หนึ่งครั้งครับ');
        } else {
          showSyncToast('เปิดแจ้งเตือนบนเครื่องนี้แล้ว');
        }
        notifyDueTasks(true);
      } else if (result === 'denied') {
        alert('การแจ้งเตือนถูกปฏิเสธ\n\nถ้าเปลี่ยนใจ เปิดได้ที่ไอคอนหน้าช่อง URL แล้วอนุญาต "การแจ้งเตือน" ครับ');
      }
    }

    // One notification per task per day, so a tab left open does not pop up on
    // every sync tick.
    function notifyDueTasks(force) {
      if (taskNotifyState() !== 'granted') return;
      const stamp = todayStr();
      let seen = {};
      try { seen = JSON.parse(localStorage.getItem(TASK_NOTIFIED_KEY) || '{}'); } catch (err) {}

      const fresh = myTaskAlerts().filter(t => force || seen[t.Task_ID] !== stamp);
      if (fresh.length === 0) return;

      try {
        const first = fresh[0];
        const body = fresh.length === 1
          ? `${first.Title}${first.Assignee ? ' — ' + first.Assignee : ''}\n${taskDue(first).label}`
          : fresh.slice(0, 3).map(t => `• ${t.Title} (${taskDue(t).label})`).join('\n') +
            (fresh.length > 3 ? `\n…และอีก ${fresh.length - 3} รายการ` : '');
        new Notification('PTS — Task ' + fresh.length + ' รายการ', { body: body, tag: 'pts-tasks', icon: 'logo.jpg' });

        const next = {};
        (state.jobTasks || []).forEach(t => {
          next[t.Task_ID] = fresh.some(f => f.Task_ID === t.Task_ID) ? stamp : (seen[t.Task_ID] || '');
        });
        localStorage.setItem(TASK_NOTIFIED_KEY, JSON.stringify(next));
      } catch (err) {
        // A blocked Notification must never break the app
      }
    }

    /* ---------------- the dialog ---------------- */

    function openTasksModal(preset) {
      if (myName() && taskScope === 'team') taskScope = 'mine';
      resetTaskForm(true);
      renderMyNamePickers();
      renderTaskTypeOptions();
      renderTaskJobOptions();
      renderTaskAssigneeFilter();
      if (preset === 'due') taskScope = myName() ? 'mine' : 'team';

      updateNotifyButton();
      renderScopeTabs();
      renderTasksList();

      // An empty list with the form folded away looks like a broken screen -
      // there is no Save button on it and nothing to edit. Open the form so
      // the first thing anyone sees is somewhere to type.
      if ((state.jobTasks || []).length === 0 && preset !== 'due') {
        const box = document.getElementById('taskFormBox');
        if (box) box.classList.remove('hidden');
        updateTaskFormToggle();
      }

      document.getElementById('tasksModal').classList.remove('hidden');
      renderTaskBackendNotice();
      if (!tasksFetched) fetchJobTasks({ silent: true });
      fetchPushStatus();   // refreshes the notification button once the server answers
    }

    function closeTasksModal() {
      resetTaskForm(true);
      document.getElementById('tasksModal').classList.add('hidden');
    }

    function setTaskScope(scope) {
      if (!myName() && scope !== 'team') {
        alert('กรุณาเลือกก่อนว่า "ฉันคือ" ใคร เพื่อให้ระบบแยกงานของคุณออกจากงานของทีมได้ครับ');
        const sel = document.getElementById(window.innerWidth < 640 ? 'taskMeSelectMobile' : 'taskMeSelect');
        if (sel) sel.focus();
        return;
      }
      taskScope = scope;
      renderScopeTabs();
      renderTasksList();
    }

    function renderScopeTabs() {
      const me = myName();
      const counts = { mine: 0, byme: 0, team: 0 };
      (state.jobTasks || []).forEach(t => {
        if (taskIsDone(t)) return;
        counts.team++;
        if (!me) return;
        if (cleanOptionValue(t.Assignee) === me) counts.mine++;
        else if (cleanOptionValue(t.Created_By) === me) counts.byme++;
      });

      ['mine', 'byme', 'team'].forEach(k => {
        const btn = document.querySelector(`[data-scope="${k}"]`);
        if (!btn) return;
        const active = taskScope === k;
        btn.className = 'px-2 py-2 rounded-lg text-xs font-bold transition ' +
          (active ? 'bg-white text-pts-900 shadow-sm border border-slate-200' : 'text-slate-600 hover:bg-white/60');
        btn.classList.toggle('opacity-50', !me && k !== 'team');
        const c = document.getElementById('taskScopeCount' + k.charAt(0).toUpperCase() + k.slice(1));
        if (c) c.textContent = counts[k] ? '(' + counts[k] + ')' : '';
      });
    }

    function renderTaskTypeOptions() {
      const sel = document.getElementById('taskTypeInput');
      if (!sel) return;
      const keep = sel.value;
      const list = getOptionList('task_types');
      const types = list.length ? list : DEFAULT_TASK_TYPES;
      sel.innerHTML = '';
      const none = document.createElement('option');
      none.value = '';
      none.textContent = 'ไม่ระบุประเภท';
      sel.appendChild(none);
      types.forEach(t => {
        const o = document.createElement('option');
        o.value = t;
        o.textContent = t;
        sel.appendChild(o);
      });
      if (keep) {
        if (!types.includes(keep) && keep !== '') {
          const o = document.createElement('option');
          o.value = keep;
          o.textContent = keep;
          sel.appendChild(o);
        }
        sel.value = keep;
      }
    }

    function renderTaskJobOptions() {
      const sel = document.getElementById('taskJobInput');
      if (!sel) return;
      const keep = sel.value;
      sel.innerHTML = '';
      const none = document.createElement('option');
      none.value = '';
      none.textContent = 'ไม่ผูกกับใบงาน (งานทั่วไป)';
      sel.appendChild(none);

      // Open jobs first - those are the ones people file tasks against
      const jobs = (state.tasks || []).slice().sort((a, b) => {
        const ac = isClosedStatus(a.Status) ? 1 : 0;
        const bc = isClosedStatus(b.Status) ? 1 : 0;
        if (ac !== bc) return ac - bc;
        return String(b.Job_ID).localeCompare(String(a.Job_ID), 'th');
      });
      jobs.forEach(j => {
        const o = document.createElement('option');
        o.value = j.Job_ID;
        o.textContent = `${j.Job_ID} · ${j.Project_Name || '-'}${isClosedStatus(j.Status) ? ' (ปิดงานแล้ว)' : ''}`;
        sel.appendChild(o);
      });
      if (keep) sel.value = keep;
    }

    function renderTaskAssigneeFilter() {
      const sel = document.getElementById('taskFilterAssignee');
      if (!sel) return;
      const keep = sel.value;
      const names = [...new Set((state.jobTasks || []).map(t => cleanOptionValue(t.Assignee)).filter(Boolean))]
        .sort((a, b) => a.localeCompare(b, 'th'));
      sel.innerHTML = '';
      const all = document.createElement('option');
      all.value = '';
      all.textContent = 'ทุกผู้รับผิดชอบ';
      sel.appendChild(all);
      names.forEach(n => {
        const o = document.createElement('option');
        o.value = n;
        o.textContent = n;
        sel.appendChild(o);
      });
      sel.value = names.includes(keep) ? keep : '';
    }

    function toggleTaskForm() {
      const box = document.getElementById('taskFormBox');
      if (!box) return;
      if (box.classList.contains('hidden')) {
        resetTaskForm();
        box.classList.remove('hidden');
        const el = document.getElementById('taskTitleInput');
        if (el) {
          // Bring the form into view, or on a small screen it opens off-screen
          box.scrollIntoView({ block: 'start', behavior: 'smooth' });
          if (!isMobileViewport()) el.focus();
        }
      } else {
        resetTaskForm(true);
      }
      updateTaskFormToggle();
    }

    /** The button has to say what it will do, or an open form looks stuck */
    function updateTaskFormToggle() {
      const btn = document.getElementById('taskNewBtn');
      const box = document.getElementById('taskFormBox');
      if (!btn || !box) return;
      const open = !box.classList.contains('hidden');
      btn.innerHTML = open
        ? '<i class="fa-solid fa-chevron-up mr-1.5"></i>ปิดฟอร์มเพิ่ม Task'
        : '<i class="fa-solid fa-plus mr-1.5"></i>เพิ่ม Task';
    }

    /**
     * The sheet side carries the Tasks tab. Until Code.gs is updated there,
     * nothing here can save - so say that plainly instead of letting someone
     * fill in a form that is going to fail.
     */
    function renderTaskBackendNotice() {
      const notice = document.getElementById('taskBackendNotice');
      if (!notice) return;
      const blocked = taskBackendReady === false;
      notice.classList.toggle('hidden', !blocked);

      const box = document.getElementById('taskFormBox');
      const addBtn = document.getElementById('taskNewBtn');
      const saveBtn = document.getElementById('taskSaveBtn');
      if (blocked) {
        if (box) box.classList.add('hidden');
        if (addBtn) { addBtn.disabled = true; addBtn.classList.add('opacity-50', 'cursor-not-allowed'); }
        if (saveBtn) saveBtn.disabled = true;
      } else {
        if (addBtn) { addBtn.disabled = false; addBtn.classList.remove('opacity-50', 'cursor-not-allowed'); }
        if (saveBtn) saveBtn.disabled = false;
      }
      updateTaskFormToggle();
    }

    function resetTaskForm(hide) {
      taskEditingId = null;
      const set = (id, v) => { const el = document.getElementById(id); if (el) el.value = v; };
      set('taskTitleInput', '');
      set('taskAssigneeInput', myName());      // most tasks are for yourself
      set('taskDueInput', todayStr());
      set('taskTimeInput', '');
      set('taskDetailInput', '');
      const pr = document.getElementById('taskPriorityInput');
      if (pr) pr.value = 'Normal';
      const ty = document.getElementById('taskTypeInput');
      if (ty) ty.value = '';
      const job = document.getElementById('taskJobInput');
      if (job) job.value = '';

      const label = document.getElementById('taskSaveBtnLabel');
      if (label) label.textContent = 'บันทึก Task';
      if (hide) {
        const box = document.getElementById('taskFormBox');
        if (box) box.classList.add('hidden');
      }
    }

    function startEditTask(taskId) {
      const t = (state.jobTasks || []).find(x => x.Task_ID === taskId);
      if (!t) return;
      markTaskRead(t);
      taskEditingId = taskId;

      renderTaskTypeOptions();
      renderTaskJobOptions();

      const set = (id, v) => { const el = document.getElementById(id); if (el) el.value = v || ''; };
      set('taskTitleInput', t.Title);
      set('taskAssigneeInput', t.Assignee);
      set('taskDueInput', taskDateOnly(t.Due_Date));
      set('taskTimeInput', taskHHMM(t.Due_Time));
      set('taskDetailInput', t.Detail);
      const pr = document.getElementById('taskPriorityInput');
      if (pr) pr.value = TASK_PRIORITIES.includes(t.Priority) ? t.Priority : 'Normal';

      const ty = document.getElementById('taskTypeInput');
      if (ty) {
        if (t.Type && !Array.from(ty.options).some(o => o.value === t.Type)) {
          const o = document.createElement('option');
          o.value = t.Type; o.textContent = t.Type;
          ty.appendChild(o);
        }
        ty.value = t.Type || '';
      }

      const job = document.getElementById('taskJobInput');
      if (job) {
        // The job may be gone - keep the value selectable rather than dropping it
        if (t.Job_ID && !Array.from(job.options).some(o => o.value === t.Job_ID)) {
          const o = document.createElement('option');
          o.value = t.Job_ID; o.textContent = t.Job_ID + ' (ไม่พบใบงานนี้แล้ว)';
          job.appendChild(o);
        }
        job.value = t.Job_ID || '';
      }

      const label = document.getElementById('taskSaveBtnLabel');
      if (label) label.textContent = 'บันทึกการแก้ไข';
      const box = document.getElementById('taskFormBox');
      if (box) box.classList.remove('hidden');
      const el = document.getElementById('taskTitleInput');
      if (el) { el.focus(); el.scrollIntoView({ block: 'center', behavior: 'smooth' }); }
      renderTasksList();
    }

    async function submitTaskForm() {
      const get = id => { const el = document.getElementById(id); return el ? el.value : ''; };
      const focusField = id => { const el = document.getElementById(id); if (el) el.focus(); };

      const title = cleanOptionValue(get('taskTitleInput'));
      if (!title) { alert('กรุณากรอกชื่อ Task'); focusField('taskTitleInput'); return; }
      const assignee = cleanOptionValue(get('taskAssigneeInput'));
      if (!assignee) { alert('กรุณาระบุผู้รับผิดชอบ'); focusField('taskAssigneeInput'); return; }
      const due = taskDateOnly(get('taskDueInput'));
      if (!due) { alert('กรุณาระบุวันที่ต้องทำ'); focusField('taskDueInput'); return; }

      const payload = {
        Task_ID: taskEditingId || '',
        Title: title,
        Detail: String(get('taskDetailInput') || '').trim(),
        Assignee: assignee,
        Due_Date: due,
        Due_Time: taskHHMM(get('taskTimeInput')),
        Job_ID: get('taskJobInput'),
        Priority: get('taskPriorityInput') || 'Normal',
        Type: get('taskTypeInput')
      };

      const ok = await withBusy('taskSaveBtn', 'กำลังบันทึก...', () => saveJobTask(payload));
      if (!ok) return;

      // A newly typed name should appear in the pickers everywhere else.
      // Best-effort: the task already saved, so a hiccup here is not an alarm.
      if (!getOptionList('technicians').includes(assignee)) {
        await addOptionValue('technicians', assignee, true);
        renderOptionLists();
        renderMyNamePickers();
      }

      showSyncToast(taskEditingId ? 'บันทึกการแก้ไขแล้ว' : `เพิ่ม Task "${title}" แล้ว`);
      resetTaskForm(true);
      renderTaskAssigneeFilter();
      renderScopeTabs();
      renderTasksList();
    }

    async function toggleTaskDone(taskId, done) {
      const t = (state.jobTasks || []).find(x => x.Task_ID === taskId);
      if (!t) return;
      const { json, error } = await apiPost({ action: 'TASK_DONE', Task_ID: taskId, Done: !!done, Actor: myName() });
      if (!json || json.status !== 'success') {
        alert('อัปเดตสถานะ Task ไม่สำเร็จ\n\n' + (json && json.message ? json.message : apiErrorText(error)));
        return;
      }
      await fetchJobTasks({ silent: true });
      markActivity();
      showSyncToast(done ? 'ทำเสร็จแล้ว' : 'ย้ายกลับไปเป็นงานค้าง');
    }

    async function removeTask(taskId) {
      const t = (state.jobTasks || []).find(x => x.Task_ID === taskId);
      if (!t) return;
      if (!confirm(`ลบ Task "${t.Title}" ออกจากระบบใช่หรือไม่?\n\nการลบนี้ย้อนกลับไม่ได้ครับ`)) return;
      if (await deleteJobTask(taskId)) showSyncToast('ลบ Task แล้ว');
    }

    function onTaskSearch(value) {
      taskQuery = value;
      renderTasksList();
    }

    function toggleShowDoneTasks() {
      taskShowDone = !taskShowDone;
      renderTasksList();
    }

    /* ---------------- list ---------------- */

    function taskSearchFields(t) {
      const job = t.Job_ID ? (state.tasks || []).find(j => j.Job_ID === t.Job_ID) : null;
      return [t.Title, t.Detail, t.Assignee, t.Created_By, t.Type, t.Job_ID, job && job.Project_Name];
    }

    function taskCardHtml(t) {
      const due = taskDue(t);
      const done = taskIsDone(t);
      const job = t.Job_ID ? (state.tasks || []).find(j => j.Job_ID === t.Job_ID) : null;
      const by = t.Created_By && cleanOptionValue(t.Created_By) !== cleanOptionValue(t.Assignee)
        ? ' · มอบหมายโดย ' + esc(t.Created_By) : '';
      const tones = {
        red:   { card: 'bg-rose-50 border-rose-300',   pill: 'bg-rose-600 text-white' },
        amber: { card: 'bg-amber-50 border-amber-300', pill: 'bg-amber-500 text-white' },
        blue:  { card: 'bg-white border-slate-200',    pill: 'bg-blue-100 text-blue-800' },
        slate: { card: 'bg-white border-slate-200',    pill: 'bg-slate-100 text-slate-600' },
        green: { card: 'bg-emerald-50/60 border-emerald-200', pill: 'bg-emerald-600 text-white' }
      };
      const tone = tones[due.tone] || tones.slate;
      const pr = t.Priority === 'Urgent' ? 'bg-rose-600 text-white'
        : t.Priority === 'High' ? 'bg-amber-500 text-white' : '';

      return `
        <div data-task-id="${esc(t.Task_ID)}" class="rounded-2xl border p-3 sm:p-3.5 shadow-xs transition ${tone.card} ${taskEditingId === t.Task_ID ? 'ring-2 ring-pts-800' : ''}">
          <div class="flex items-start gap-3">
            <button type="button" data-done class="w-7 h-7 shrink-0 rounded-lg border-2 flex items-center justify-center transition active:scale-90 ${done ? 'bg-emerald-600 border-emerald-600 text-white' : 'bg-white border-slate-300 text-transparent hover:border-emerald-500'}" title="${done ? 'ย้ายกลับเป็นงานค้าง' : 'ทำเสร็จแล้ว'}">
              <i class="fa-solid fa-check text-xs"></i>
            </button>

            <div class="min-w-0 flex-1">
              <p class="text-sm font-bold text-slate-900 ${done ? 'line-through text-slate-500' : ''}">
                ${isTaskUnread(t) ? '<span class="inline-block w-2 h-2 rounded-full bg-rose-600 mr-1.5 align-middle" title="ยังไม่ได้เปิดดู"></span>' : ''}${esc(t.Title)}
              </p>
              <div class="mt-1 flex flex-wrap items-center gap-x-2.5 gap-y-1 text-[11px]">
                <span class="inline-flex items-center gap-1 px-2 py-0.5 rounded-md font-bold ${tone.pill}">${esc(due.label)}</span>
                ${pr && !done ? `<span class="px-2 py-0.5 rounded-md font-bold ${pr}">${esc(t.Priority)}</span>` : ''}
                <span class="text-slate-600"><i class="fa-solid fa-user mr-1 text-slate-400"></i>${esc(t.Assignee)}${by}</span>
                ${t.Type ? `<span class="text-slate-500"><i class="fa-solid fa-tag mr-1 text-slate-400"></i>${esc(t.Type)}</span>` : ''}
                ${t.Job_ID ? `<button type="button" data-job class="text-pts-800 font-semibold hover:underline"><i class="fa-solid fa-link mr-1"></i>${esc(t.Job_ID)}${job ? ' · ' + esc(job.Project_Name || '') : ' (ไม่พบใบงาน)'}</button>` : ''}
              </div>
              ${t.Detail ? `<p class="mt-1.5 text-[11px] text-slate-600 whitespace-pre-wrap">${esc(t.Detail)}</p>` : ''}
              ${done && t.Done_By ? `<p class="mt-1 text-[10px] text-emerald-700 font-semibold">เสร็จโดย ${esc(t.Done_By)}${t.Done_At ? ' · ' + esc(t.Done_At) : ''}</p>` : ''}
            </div>

            <div class="flex items-center gap-1 shrink-0">
              <button type="button" data-edit class="w-8 h-8 rounded-lg bg-white/80 hover:bg-white border border-slate-300 text-slate-600 transition" title="แก้ไข"><i class="fa-solid fa-pen-to-square text-xs"></i></button>
              <button type="button" data-del class="w-8 h-8 rounded-lg bg-white/80 hover:bg-rose-50 border border-slate-300 text-rose-600 transition" title="ลบ"><i class="fa-solid fa-trash-can text-xs"></i></button>
            </div>
          </div>
        </div>`;
    }

    function renderTasksList() {
      const host = document.getElementById('tasksList');
      if (!host) return;

      const assignee = (document.getElementById('taskFilterAssignee') || {}).value || '';
      const priority = (document.getElementById('taskFilterPriority') || {}).value || '';
      const q = cleanOptionValue(taskQuery).toLowerCase();

      let list = (state.jobTasks || []).filter(scopeFilter(taskScope)).filter(t =>
        (!assignee || cleanOptionValue(t.Assignee) === assignee) &&
        (!priority || (t.Priority || 'Normal') === priority));
      if (q) list = list.filter(t => taskSearchFields(t).some(v => String(v || '').toLowerCase().includes(q)));

      const open = list.filter(t => !taskIsDone(t)).sort(byDue);
      const done = list.filter(taskIsDone).sort((a, b) => String(b.Done_At).localeCompare(String(a.Done_At)));
      const group = k => open.filter(t => taskDue(t).key === k);
      const over = group('overdue');

      const section = (title, items, tone) => items.length ? `
        <section class="space-y-2">
          <h4 class="text-xs font-bold uppercase tracking-wider flex items-center gap-2 ${tone || 'text-slate-600'}">
            ${esc(title)}<span class="px-1.5 py-0.5 rounded-md bg-white border border-slate-200 text-[10px] font-mono">${items.length}</span>
          </h4>
          ${items.map(taskCardHtml).join('')}
        </section>` : '';

      const empty = q ? `ไม่พบ Task ที่มีคำว่า “${esc(q)}”`
        : (assignee || priority) ? 'ไม่มี Task ที่ตรงกับตัวกรอง'
        : taskScope === 'mine' ? 'ไม่มีงานค้างของคุณ 🎉'
        : taskScope === 'byme' ? 'คุณยังไม่ได้มอบหมายงานให้ใคร'
        : (state.jobTasks || []).length === 0 ? 'ยังไม่มี Task ในระบบ — กด "เพิ่ม Task" ด้านบนได้เลย' : 'ไม่มีงานค้าง';

      host.innerHTML =
        (open.length ? '' : `<div class="py-10 text-center bg-white rounded-2xl border border-dashed border-slate-300">
            <i class="fa-solid fa-clipboard-check text-3xl text-slate-300 mb-2 block"></i>
            <p class="text-xs font-bold text-slate-600">${empty}</p>
          </div>`)
        + section('เกินกำหนด', over, 'text-rose-700')
        + section('วันนี้', group('today'), 'text-amber-700')
        + section('พรุ่งนี้', group('soon'), 'text-blue-700')
        + section('ถัดไป', group('later'))
        + (done.length ? `<button type="button" data-toggle-done class="w-full px-4 py-2.5 rounded-xl border border-slate-300 bg-white hover:bg-slate-100 text-slate-700 text-xs font-bold transition">
             ${taskShowDone ? 'ซ่อน' : 'แสดง'}งานที่เสร็จแล้ว (${done.length})</button>` : '')
        + (taskShowDone ? section('เสร็จแล้ว', done, 'text-emerald-700') : '');

      // Wire each card by its own id, never by position: the sections are
      // rebuilt on every render and an index-based pairing would quietly send
      // a click to the wrong task.
      host.querySelectorAll('[data-task-id]').forEach(card => {
        const t = (state.jobTasks || []).find(x => x.Task_ID === card.dataset.taskId);
        if (!t) return;
        card.querySelector('[data-done]').onclick = () => toggleTaskDone(t.Task_ID, !taskIsDone(t));
        card.querySelector('[data-edit]').onclick = () => startEditTask(t.Task_ID);
        card.querySelector('[data-del]').onclick = () => removeTask(t.Task_ID);
        const jobBtn = card.querySelector('[data-job]');
        if (jobBtn) jobBtn.onclick = () => {
          const job = (state.tasks || []).find(j => j.Job_ID === t.Job_ID);
          if (!job) { alert(`ไม่พบใบงาน ${t.Job_ID} ในระบบแล้ว (อาจถูกลบไป)`); return; }
          markTaskRead(t);
          closeTasksModal();
          openDetailModal(t.Job_ID);
        };
      });
      const toggle = host.querySelector('[data-toggle-done]');
      if (toggle) toggle.onclick = toggleShowDoneTasks;

      const sub = document.getElementById('taskHeaderSub');
      if (sub) {
        sub.textContent = `${open.length} งานค้าง` + (over.length ? ` · เกินกำหนด ${over.length}` : '') +
          (myName() ? '' : ' · ยังไม่ได้เลือกว่าฉันคือใคร');
      }
    }

    /* =======================================================================
     * WEB PUSH - แจ้งเตือนถึงเครื่อง แม้ปิดแอปอยู่
     *
     * ยกมาจากแอพงานขาย: หน้าเว็บลงทะเบียน endpoint ของเครื่องไว้ที่หลังบ้าน
     * แล้วตัวตั้งเวลาใน Apps Script เป็นคนส่ง ข้อความจึงมาถึงแม้ปิดแอป
     *
     * ถ้าหลังบ้านยังไม่ได้รัน setupPush() ทุกอย่างในนี้จะเงียบและไม่ทำอะไร
     * แอปยังเตือนผ่านตัวเลขบนปุ่ม Task กับแถบสีด้านบนตามปกติ
     * ===================================================================== */
    let pushConfig = null;        // { ready, running, publicKey, registered }
    let swRegistration = null;

    function pushSupported() {
      return 'serviceWorker' in navigator && 'PushManager' in window &&
             'Notification' in window && window.isSecureContext;
    }

    async function registerPushServiceWorker() {
      if (!pushSupported()) return null;
      if (swRegistration) return swRegistration;
      try {
        swRegistration = await navigator.serviceWorker.register('sw.js', { updateViaCache: 'none' });
        // Ask once per app start for a newer worker; without this a device can
        // sit on an old one for hours.
        if (swRegistration.update) swRegistration.update().catch(() => {});
        await navigator.serviceWorker.ready;
        return swRegistration;
      } catch (err) {
        swRegistration = null;
        return null;
      }
    }

    async function currentPushSubscription() {
      const reg = await registerPushServiceWorker();
      if (!reg) return null;
      try { return await reg.pushManager.getSubscription(); } catch (err) { return null; }
    }

    /** Ask the server whether push is configured, and whether this device is on the list */
    async function fetchPushStatus() {
      if (!pushSupported() || state.dbMode === 'mock' || !state.apiUrl) { pushConfig = null; return null; }
      const sub = await currentPushSubscription();
      const { json } = await apiGetDetailed({ push: '1', endpoint: sub ? sub.endpoint : '' });
      pushConfig = (json && json.status === 'success' && json.push) ? json.push : null;
      updateNotifyButton();
      return pushConfig;
    }

    function urlBase64ToUint8Array(base64) {
      const padded = String(base64 || '').replace(/-/g, '+').replace(/_/g, '/');
      const raw = atob(padded + '='.repeat((4 - padded.length % 4) % 4));
      const out = new Uint8Array(raw.length);
      for (let i = 0; i < raw.length; i++) out[i] = raw.charCodeAt(i);
      return out;
    }

    const subKey = (sub, name) => {
      const raw = sub.getKey(name);
      if (!raw) return '';
      return btoa(String.fromCharCode.apply(null, new Uint8Array(raw)))
        .replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '');
    };

    /**
     * Subscribe this device and tell the server. Needs a name, because the
     * server sends each person only their own tasks.
     */
    async function enablePushOnThisDevice() {
      if (!pushSupported()) return { ok: false, reason: 'unsupported' };
      if (!pushConfig) await fetchPushStatus();
      if (!pushConfig || !pushConfig.ready || !pushConfig.publicKey) return { ok: false, reason: 'server' };
      if (!myName()) return { ok: false, reason: 'noname' };

      const reg = await registerPushServiceWorker();
      if (!reg) return { ok: false, reason: 'sw' };

      let sub = null;
      try {
        sub = await reg.pushManager.getSubscription();
        // A key change on the server makes an old subscription undeliverable
        if (sub) {
          const same = subKey(sub, 'p256dh') && btoa(String.fromCharCode.apply(null, new Uint8Array(sub.options.applicationServerKey || new ArrayBuffer(0))))
            .replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '') === pushConfig.publicKey;
          if (!same) { try { await sub.unsubscribe(); } catch (e) {} sub = null; }
        }
        if (!sub) {
          sub = await reg.pushManager.subscribe({
            userVisibleOnly: true,
            applicationServerKey: urlBase64ToUint8Array(pushConfig.publicKey)
          });
        }
      } catch (err) {
        return { ok: false, reason: 'subscribe' };
      }

      const { json } = await apiPost({
        action: 'PUSH_SUB_SAVE',
        endpoint: sub.endpoint,
        p256dh: subKey(sub, 'p256dh'),
        auth: subKey(sub, 'auth'),
        name: myName(),
        ua: navigator.userAgent.slice(0, 200)
      });
      if (!json || json.status !== 'success') return { ok: false, reason: 'save' };

      pushConfig.registered = true;
      updateNotifyButton();
      return { ok: true };
    }

    /**
     * Re-register silently when this device is already allowed but the server
     * has lost it, or the owner's name changed - otherwise the person quietly
     * stops getting notifications and never finds out why.
     */
    async function refreshPushRegistration() {
      if (!pushConfig || !pushConfig.ready || !myName()) return;
      if (taskNotifyState() !== 'granted') return;
      const sub = await currentPushSubscription();
      if (!sub && pushConfig.registered) return;
      await enablePushOnThisDevice();
    }

    async function disablePushOnThisDevice() {
      const sub = await currentPushSubscription();
      if (!sub) return;
      try { await apiPost({ action: 'PUSH_SUB_DELETE', endpoint: sub.endpoint }); } catch (err) {}
      try { await sub.unsubscribe(); } catch (err) {}
      if (pushConfig) pushConfig.registered = false;
      updateNotifyButton();
    }

    /** Send a test notification to this device, after a delay so the app can be closed */
    async function sendTestPush() {
      const sub = await currentPushSubscription();
      if (!sub) { alert('ยังไม่ได้เปิดการแจ้งเตือนบนเครื่องนี้'); return; }
      const wait = confirm('จะส่งข้อความทดสอบมาที่เครื่องนี้\n\nกด "ตกลง" แล้วปิดแอป/สลับไปแอปอื่นภายใน 10 วินาที เพื่อดูว่าเด้งจริงแม้ปิดแอป\n(กด "ยกเลิก" เพื่อส่งทันที)');
      const btn = document.getElementById('taskTestPushBtn');
      const original = btn ? btn.innerHTML : '';
      if (btn) { btn.disabled = true; btn.innerHTML = '<i class="fa-solid fa-circle-notch fa-spin mr-1.5"></i>กำลังส่ง...'; }
      try {
        const { json, error } = await apiPost({ action: 'PUSH_TEST', endpoint: sub.endpoint, delay: wait ? 10 : 0 });
        if (json && json.status === 'success') showSyncToast('ส่งข้อความทดสอบแล้ว');
        else alert('ส่งข้อความทดสอบไม่สำเร็จ\n\n' + (json && json.message ? json.message : apiErrorText(error)));
      } finally {
        if (btn) { btn.disabled = false; btn.innerHTML = original; }
      }
    }

    /**
     * The service worker tells the page when a push arrives or a notification
     * is tapped, so an app left open reacts straight away.
     */
    function listenToServiceWorker() {
      if (!('serviceWorker' in navigator)) return;
      navigator.serviceWorker.addEventListener('message', event => {
        const data = event.data || {};
        if (data.type === 'push') {
          fetchJobTasks({ silent: true });
        } else if (data.type === 'open-tasks') {
          fetchJobTasks({ silent: true }).then(() => openTasksModal('due'));
        }
      });
      // Opened from a notification on a cold start
      try {
        if (new URLSearchParams(location.search).get('view') === 'tasks') {
          setTimeout(() => openTasksModal('due'), 1200);
        }
      } catch (err) {}
    }

    /* =======================================================================
     * TASK OVERVIEW PER JOB
     *
     * The board answers "what is the state of this job"; it could not answer
     * "and what does someone still have to DO about it". These helpers put the
     * next action for each job right in the row, and the full list inside the
     * job's own dialog.
     * ===================================================================== */

    /** Open tasks for one job, most urgent first */
    function jobOpenTasks(jobId) {
      const id = String(jobId || '').trim();
      if (!id) return;
      return (state.jobTasks || [])
        .filter(t => String(t.Job_ID || '').trim() === id && !taskIsDone(t))
        .sort(byDue);
    }

    function jobTaskSummary(jobId) {
      const open = jobOpenTasks(jobId) || [];
      if (open.length === 0) return { count: 0 };
      const worst = open[0];
      const due = taskDue(worst);
      return {
        count: open.length,
        next: worst,
        due: due,
        late: open.filter(t => taskDue(t).key === 'overdue').length,
        today: open.filter(t => taskDue(t).key === 'today').length
      };
    }

    /** The "Task ที่ต้องทำ" cell in the job table */
    function jobTaskCellHtml(jobId) {
      const s = jobTaskSummary(jobId);
      if (s.count === 0) {
        return `<button type="button" data-job-tasks class="text-xs text-slate-400 hover:text-pts-800 transition" title="ยังไม่มี Task — กดเพื่อเพิ่ม">
            <i class="fa-regular fa-square-plus mr-1"></i>เพิ่ม
          </button>`;
      }
      const tone = s.late > 0 ? 'bg-rose-100 text-rose-900 border-rose-300'
        : s.today > 0 ? 'bg-amber-100 text-amber-900 border-amber-300'
        : 'bg-slate-100 text-slate-700 border-slate-300';
      const icon = s.late > 0 ? 'fa-triangle-exclamation' : s.today > 0 ? 'fa-bell' : 'fa-list-check';
      return `<button type="button" data-job-tasks class="text-left w-full group" title="กดเพื่อดู Task ทั้งหมดของใบงานนี้">
          <span class="inline-flex items-center gap-1 px-2 py-0.5 rounded-lg text-[11px] font-bold border ${tone}">
            <i class="fa-solid ${icon}"></i>${s.count} งาน
          </span>
          <p class="mt-1 text-[11px] text-slate-700 font-semibold leading-tight line-clamp-2 group-hover:text-pts-800">${esc(s.next.Title)}</p>
          <p class="text-[10px] text-slate-500">${esc(s.due.label)}${s.next.Assignee ? ' · ' + esc(s.next.Assignee) : ''}</p>
        </button>`;
    }

    /** Jump into the Task dialog already filtered to one job */
    function openTasksForJobModal(jobId) {
      closeDetailModal();
      openTasksModal();
      taskScope = 'team';                       // a job's tasks may belong to anyone
      const search = document.getElementById('taskSearchInput');
      if (search) search.value = jobId;
      taskQuery = jobId;
      const box = document.getElementById('taskFormBox');
      if (box) box.classList.add('hidden');
      updateTaskFormToggle();
      renderScopeTabs();
      renderTasksList();
    }

    /** Start a new task already tied to one job */
    function addTaskForJob(jobId) {
      closeDetailModal();
      openTasksModal();
      resetTaskForm();
      renderTaskJobOptions();
      const box = document.getElementById('taskFormBox');
      if (box) box.classList.remove('hidden');
      updateTaskFormToggle();
      const sel = document.getElementById('taskJobInput');
      if (sel) sel.value = jobId;
      const title = document.getElementById('taskTitleInput');
      if (title) { title.focus(); title.scrollIntoView({ block: 'center', behavior: 'smooth' }); }
    }

    /** The Task section inside a job's detail dialog */
    function renderDetailTasks(jobId) {
      const host = document.getElementById('detailTasksList');
      const countEl = document.getElementById('detailTaskCount');
      if (!host) return;

      const all = (state.jobTasks || []).filter(t => String(t.Job_ID || '').trim() === String(jobId || '').trim());
      const open = all.filter(t => !taskIsDone(t)).sort(byDue);
      const done = all.filter(taskIsDone);
      if (countEl) countEl.textContent = `(ค้าง ${open.length}${done.length ? ' · เสร็จ ' + done.length : ''})`;

      if (all.length === 0) {
        host.innerHTML = `<p class="text-xs text-slate-400 italic p-1">ยังไม่มี Task ของใบงานนี้ — กด "เพิ่ม Task" เพื่อระบุสิ่งที่ต้องทำต่อ</p>`;
        return;
      }

      const tones = { red: 'bg-rose-50 border-rose-300', amber: 'bg-amber-50 border-amber-300', blue: 'bg-white border-slate-200', slate: 'bg-white border-slate-200', green: 'bg-emerald-50/60 border-emerald-200' };
      const pills = { red: 'bg-rose-600 text-white', amber: 'bg-amber-500 text-white', blue: 'bg-blue-100 text-blue-800', slate: 'bg-slate-100 text-slate-600', green: 'bg-emerald-600 text-white' };

      host.innerHTML = open.concat(done).map(t => {
        const due = taskDue(t);
        const isDone = taskIsDone(t);
        return `<div class="flex items-start gap-2.5 p-2.5 rounded-xl border ${tones[due.tone] || tones.slate}">
            <i class="fa-solid ${isDone ? 'fa-circle-check text-emerald-600' : 'fa-circle-dot text-slate-400'} mt-0.5 text-xs shrink-0"></i>
            <div class="min-w-0 flex-1">
              <p class="text-xs font-bold text-slate-900 ${isDone ? 'line-through text-slate-500' : ''}">${esc(t.Title)}</p>
              <div class="mt-0.5 flex flex-wrap items-center gap-x-2 gap-y-0.5 text-[10px]">
                <span class="px-1.5 py-0.5 rounded font-bold ${pills[due.tone] || pills.slate}">${esc(due.label)}</span>
                <span class="text-slate-600">${esc(t.Assignee || 'ไม่ระบุผู้รับผิดชอบ')}</span>
                ${t.Type ? `<span class="text-slate-500">${esc(t.Type)}</span>` : ''}
              </div>
            </div>
          </div>`;
      }).join('');
    }
