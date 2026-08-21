const MAX_PHOTOS_PER_TASK = 100;
    // Base64 inflates a file by about a third, and Apps Script answers an
    // oversized POST with an HTML error page rather than JSON. Batching keeps a
    // *set* of files small, but one single file always travels alone - so the
    // per-file ceiling has to stay under the limit on its own.
    const MAX_ATTACHMENT_MB = 6;

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
      "แจ้งเข้าทำงาน",
      "เข้าปฏิบัติงานหน้างาน",
      "รอ Testing & Commissioning",
      "ติดปัญหาหน้างาน / รออะไหล่",
      "ค้างทำรายงาน",
      "ส่งมอบ / เอกสารเรียบร้อย"
    ];

    const ALL_STATUSES = [
      "วางแผน / เตรียมอุปกรณ์",
      "แจ้งเข้าทำงาน",
      "เข้าปฏิบัติงานหน้างาน",
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

      // 3. เข้าปฏิบัติงานหน้างาน
      if (s === 'เข้าปฏิบัติงานหน้างาน' || s === 'ระหว่างปฏิบัติงาน' || s === 'กำลังปฏิบัติงาน' || s.includes('ปฏิบัติงาน') || s.includes('หน้างาน') || s.includes('ทีมช่างอยู่หน้างาน')) {
        return 'เข้าปฏิบัติงานหน้างาน';
      }

      // 2. แจ้งเข้าทำงาน
      if (s === 'แจ้งเข้าทำงาน' || s === 'แจ้งเข้าทำงานแล้ว' || s.includes('แจ้งเข้า') || s.includes('ส่งหนังสือเข้างาน') || s.includes('อนุมัติเข้าพื้นที่') || s.includes('ดำเนินการแก้ไข')) {
        return 'แจ้งเข้าทำงาน';
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
      currentDeliveryJobId: null
    };

    // INITIALIZATION
    window.addEventListener('DOMContentLoaded', () => {
      localStorage.setItem('pts_db_mode', state.dbMode);

      loadState();
      renderApp();
      startAutoRefresh();

      if (apiUrlMigratedFrom) {
        showSyncToast('เปลี่ยนไปใช้ลิงก์ Apps Script ใหม่ในไฟล์แล้ว (กด ⚙️ ตั้งค่า API เพื่อดู/แก้)');
      }
    });

    function loadState() {
      if (state.dbMode === 'mock') {
        const stored = localStorage.getItem('pts_mock_tasks');
        if (stored) {
          try {
            state.tasks = JSON.parse(stored);
          } catch(e) {
            state.tasks = [...INITIAL_SAMPLE_TASKS];
          }
        } else {
          state.tasks = [...INITIAL_SAMPLE_TASKS];
          saveMockStorage();
        }
      } else {
        fetchTasks();
      }
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

    async function uploadAttachmentsInBackground(attachFields, fileCount) {
      const jobId = attachFields.Job_ID;
      const settled = {};
      const queue = [];

      ATTACHMENT_FIELDS.forEach(field => {
        settled[field] = [];
        parseAttachmentValue(attachFields[field]).forEach(item => {
          const src = typeof item === 'string' ? item : (item && (item.dataUrl || item.url)) || '';
          if (String(src).indexOf('data:') === 0) queue.push({ field, item, chars: String(src).length });
          else settled[field].push(item);
        });
      });

      let uploaded = 0;

      while (queue.length > 0) {
        const batch = [];
        let chars = 0;
        // Always take at least one, even if that single file is over the budget
        while (queue.length > 0 && (batch.length === 0 || chars + queue[0].chars <= UPLOAD_BATCH_CHARS)) {
          const next = queue.shift();
          chars += next.chars;
          batch.push(next);
        }

        showSyncToast(`กำลังอัปโหลดไฟล์แนบ ${uploaded + batch.length}/${fileCount} ...`);

        const data = { Job_ID: jobId };
        ATTACHMENT_FIELDS.forEach(field => {
          const additions = batch.filter(b => b.field === field).map(b => b.item);
          data[field] = JSON.stringify(settled[field].concat(additions));
        });

        const { json, error } = await apiPost({ action: 'UPDATE', data: data });

        if (!json || json.status !== 'success') {
          state.lastError = error || 'not-json';
          updateStatusBadge();
          alert("บันทึกใบงานสำเร็จแล้ว แต่อัปโหลดไฟล์แนบไม่สำเร็จ\n\n" +
            (json && json.message ? json.message : apiErrorText(error)) +
            `\n\nอัปโหลดสำเร็จแล้ว ${uploaded} จาก ${fileCount} ไฟล์ — ข้อมูลใบงานและไฟล์ที่ขึ้นไปแล้วไม่หาย` +
            "\nเปิดใบงานนี้แล้วแนบไฟล์ที่เหลือใหม่อีกครั้งได้เลย");
          fetchTasks({ silent: true, fresh: true }).then(markActivity);
          return;
        }

        reportSaveWarnings(json.warnings);
        uploaded += batch.length;

        if (json.attachments) {
          ATTACHMENT_FIELDS.forEach(field => {
            if (typeof json.attachments[field] !== 'undefined') {
              settled[field] = parseAttachmentValue(json.attachments[field]);
            }
          });
        } else if (queue.length > 0) {
          alert("อัปโหลดไฟล์แนบต่อไม่ได้\n\nสคริปต์ฝั่ง Google ยังเป็นเวอร์ชันเก่า " +
            "กรุณาวาง Code.gs ใหม่แล้ว Deploy > Manage deployments > New version\n\n" +
            `(อัปโหลดสำเร็จแล้ว ${uploaded} จาก ${fileCount} ไฟล์)`);
          fetchTasks({ silent: true, fresh: true }).then(markActivity);
          return;
        }
      }

      showSyncToast(`อัปโหลดไฟล์แนบเรียบร้อย (${fileCount} ไฟล์)`);
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
      let m = /^(\d{4})-(\d{1,2})-(\d{1,2})/.exec(s);
      if (m) return new Date(+m[1], +m[2] - 1, +m[3]);

      m = /^(\d{1,2})\/(\d{1,2})\/(\d{4})/.exec(s);
      if (m) return new Date(+m[3], +m[2] - 1, +m[1]);

      const d = new Date(s);
      return isNaN(d.getTime()) ? null : new Date(d.getFullYear(), d.getMonth(), d.getDate());
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

      return (state.tasks || []).filter(t => {
        const isClosed = isClosedStatus(t.Status);
        const norm = normalizeStatus(t.Status);

        // Separate closed vs active jobs:
        if (statusFilterVal === 'ACTIVE_ALL' || statusFilterVal === 'ALL') {
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
    }

    function handleStatusFilterChange() {
      renderApp();
    }
    function renderApp() {
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
      const clearedCount = all.filter(t => normalizeStatus(t.Status) === 'แจ้งเข้าทำงาน').length;
      const onSiteCount = all.filter(t => normalizeStatus(t.Status) === 'เข้าปฏิบัติงานหน้างาน').length;
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
        { id: 'card-metric-2', key: 'แจ้งเข้าทำงาน' },
        { id: 'card-metric-3', key: 'เข้าปฏิบัติงานหน้างาน' },
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
            <td colspan="10" class="text-center py-8 text-slate-400 text-sm">
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
          if (isCompleted) {
            contractBadge = `<span class="inline-block px-2.5 py-0.5 rounded-lg text-xs font-bold bg-slate-100 text-slate-700">${esc(task.Contract_Expiry_Date)}</span>`;
          } else if (daysRem !== null && daysRem < 0) {
            contractBadge = `<span class="inline-block px-2.5 py-0.5 rounded-lg text-xs font-bold bg-rose-100 text-pts-900 border border-rose-300 animate-pulse">🚨 เลย ${Math.abs(daysRem)} วัน</span>`;
          } else if (daysRem !== null && daysRem <= 7) {
            contractBadge = `<span class="inline-block px-2.5 py-0.5 rounded-lg text-xs font-bold bg-amber-100 text-amber-950 border border-amber-300 animate-pulse">⚠️ ครบสัญญาใน ${daysRem} วัน</span>`;
          } else {
            contractBadge = `<span class="inline-block px-2.5 py-0.5 rounded-lg text-xs font-medium bg-slate-100 text-slate-700">${esc(task.Contract_Expiry_Date)}</span>`;
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
                ${isCompleted && task.Completion_Date ? `<span class="inline-block px-2 py-0.5 rounded-lg text-[10px] font-bold bg-emerald-100 text-emerald-900">เสร็จ: ${esc(task.Completion_Date)}</span>` : (task.Contract_Expiry_Date ? contractBadge : '')}
              </div>
            </div>
          </td>
          <td class="hidden md:table-cell py-3 px-4 text-xs font-semibold text-slate-600 align-top">${esc(task.Sub_Department || '-')}</td>
          <td class="py-3 px-2 sm:px-4 text-xs font-medium text-slate-700 align-top">
            <i class="fa-solid fa-user-gear text-slate-400 mr-1"></i>${esc(task.Technician_In_Charge || '-')}
          </td>
          <td class="hidden sm:table-cell py-3 px-4 text-xs font-mono text-slate-600 align-top">${esc(task.PO_Approval_Date || '-')}</td>
          <td class="py-3 px-2 sm:px-4 align-top">
            ${isCompleted ? `
              <span class="inline-flex items-center px-2.5 py-1 rounded-xl text-xs font-bold font-mono bg-emerald-50 text-emerald-950 border border-emerald-300">
                <i class="fa-solid fa-calendar-check text-emerald-600 mr-1"></i>${esc(task.Completion_Date || 'เสร็จสิ้น')}
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
                ${deliveryFiles.length > 0 ? `<span class="px-1.5 py-0.5 rounded bg-emerald-50 text-emerald-800 font-black border border-emerald-200" title="มีใบส่งมอบงาน"><i class="fa-solid fa-file-signature mr-0.5 text-emerald-600"></i>${deliveryFiles.length}</span>` : ''}
                ${videos.length > 0 ? `<span class="px-1.5 py-0.5 rounded bg-purple-50 text-purple-700"><i class="fa-solid fa-video mr-0.5"></i>${videos.length}</span>` : ''}
                ${photos.length === 0 && docs.length === 0 && deliveryFiles.length === 0 && videos.length === 0 ? '<span class="text-slate-400 font-normal">-</span>' : ''}
              </div>
            </div>
          </td>
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

      // Anything much larger than this turns into a Base64 body Apps Script
      // rejects, and the user only saw an opaque network error.
      const tooBig = files.filter(f => f.size > MAX_ATTACHMENT_MB * 1024 * 1024);
      const usable = files.filter(f => f.size <= MAX_ATTACHMENT_MB * 1024 * 1024);
      if (tooBig.length > 0) {
        alert(`ไฟล์ต่อไปนี้ใหญ่เกิน ${MAX_ATTACHMENT_MB} MB จึงแนบผ่านแอปไม่ได้:\n\n` +
              tooBig.map(f => `• ${f.name} (${formatBytes(f.size)})`).join('\n') +
              `\n\nเป็นข้อจำกัดของ Google Apps Script ที่รับไฟล์ต่อครั้งได้จำกัด\n` +
              `วิธีแก้สำหรับคลิปยาว: อัปโหลดขึ้น Google Drive เองแล้ววางลิงก์ไว้ในช่อง "หมายเหตุ"`);
      }
      if (usable.length === 0) return;

      if (fileCategory === 'photo') {
        const currentPhotos = mode === 'new' ? state.tempNewPhotos : state.tempEditPhotos;
        const availableSlots = MAX_PHOTOS_PER_TASK - currentPhotos.length;

        if (availableSlots <= 0) {
          alert(`ใบงานนี้แนบรูปภาพครบโควต้าสูงสุดแล้ว (${MAX_PHOTOS_PER_TASK} รูป)`);
          return;
        }

        let filesToProcess = usable;
        if (filesToProcess.length > availableSlots) {
          alert(`คุณเลือกรูปภาพเกินโควต้า ระบบจะเพิ่มรูปให้เพียง ${availableSlots} รูปแรกเท่านั้น`);
          filesToProcess = filesToProcess.slice(0, availableSlots);
        }

        // Wait for all of them: compression finishes out of order, so pushing as
        // each one lands used to shuffle the photos. Promise.all keeps the order
        // the user picked, and a null means that file could not be decoded.
        const results = await Promise.all(filesToProcess.map(f => compressImage(f, PHOTO_MAX_WIDTH, PHOTO_QUALITY)));
        const target = mode === 'new' ? state.tempNewPhotos : state.tempEditPhotos;
        results.forEach(dataUrl => { if (dataUrl) target.push(dataUrl); });
        renderPhotosPreview(mode);

        const failed = results.filter(r => !r).length;
        if (failed > 0) alert(`อ่านไฟล์รูปภาพไม่สำเร็จ ${failed} ไฟล์ (ไฟล์อาจเสียหายหรือเป็นชนิดที่เบราว์เซอร์ไม่รองรับ เช่น HEIC)`);
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
      if (mode === 'new') {
        const container = document.getElementById('newPhotosPreviewContainer');
        const text = document.getElementById('newPhotoCountText');
        container.innerHTML = '';
        const count = state.tempNewPhotos.length;
        text.textContent = `(${count} / ${MAX_PHOTOS_PER_TASK} รูป)`;

        state.tempNewPhotos.forEach((imgUrl, idx) => {
          const div = document.createElement('div');
          div.className = "relative group w-14 h-14 rounded-xl overflow-hidden border border-slate-300 shadow-sm shrink-0";
          div.innerHTML = `
            <img src="${escUrl(photoSrc(imgUrl))}" class="w-full h-full object-cover" alt="Preview">
            <button type="button" onclick="removeAttachment('new', 'photo', ${idx})" class="absolute top-0.5 right-0.5 bg-rose-600 text-white rounded-full w-6 h-6 sm:w-4 sm:h-4 text-[11px] sm:text-[9px] flex items-center justify-center shadow hover:bg-rose-700 active:scale-90 transition">
              <i class="fa-solid fa-xmark"></i>
            </button>
          `;
          container.appendChild(div);
        });
      } else if (mode === 'edit') {
        const container = document.getElementById('editPhotosGalleryContainer');
        const badge = document.getElementById('editPhotoCountBadge');
        container.innerHTML = '';
        const count = state.tempEditPhotos.length;
        badge.textContent = `(${count} / ${MAX_PHOTOS_PER_TASK} รูป)`;

        if (count === 0) {
          container.innerHTML = `<p class="text-xs text-slate-400 p-1 italic">ยังไม่มีรูปภาพแนบ</p>`;
          return;
        }

        state.tempEditPhotos.forEach((imgUrl, idx) => {
          const div = document.createElement('div');
          div.className = "relative group w-16 h-16 rounded-xl overflow-hidden border border-slate-300 shadow-sm cursor-pointer shrink-0";
          div.onclick = () => openPhotoLightbox(state.tempEditPhotos, idx, state.currentEditingTask);
          div.innerHTML = `
            <img src="${escUrl(photoSrc(imgUrl))}" class="w-full h-full object-cover hover:scale-110 transition duration-200" alt="Site Photo">
            <button type="button" onclick="event.stopPropagation(); removeAttachment('edit', 'photo', ${idx})" class="absolute top-0.5 right-0.5 bg-rose-600 text-white rounded-full w-6 h-6 sm:w-4 sm:h-4 text-[11px] sm:text-[9px] flex items-center justify-center shadow hover:bg-rose-700 active:scale-90 transition">
              <i class="fa-solid fa-xmark"></i>
            </button>
          `;
          container.appendChild(div);
        });
      }
    }

    // RENDER DOCUMENTS PREVIEW (PDF, WORD, EXCEL)
    function renderDocsPreview(mode) {
      const docs = mode === 'new' ? state.tempNewDocs : state.tempEditDocs;
      const container = document.getElementById(mode === 'new' ? 'newDocsPreviewContainer' : 'editDocsGalleryContainer');
      const label = document.getElementById(mode === 'new' ? 'newDocCountText' : 'editDocCountBadge');
      
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
      const videos = mode === 'new' ? state.tempNewVideos : state.tempEditVideos;
      const container = document.getElementById(mode === 'new' ? 'newVideosPreviewContainer' : 'editVideosGalleryContainer');
      const label = document.getElementById(mode === 'new' ? 'newVideoCountText' : 'editVideoCountBadge');

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
      const docs = mode === 'new' ? (state.tempNewDelivery || []) : (state.tempEditDelivery || []);
      const container = document.getElementById(mode === 'new' ? 'newDeliveryPreviewContainer' : 'editDeliveryGalleryContainer');
      const label = document.getElementById(mode === 'new' ? 'newDeliveryCountText' : 'editDeliveryCountBadge');
      
      if (!container || !label) return;
      container.innerHTML = '';
      label.textContent = `(${docs.length} ไฟล์)`;

      if (docs.length === 0) {
        container.innerHTML = `<p class="text-xs text-blue-600/70 p-1 italic">ยังไม่มีไฟล์ใบส่งมอบงานแนบ</p>`;
        return;
      }

      docs.forEach((doc, idx) => {
        const ext = getFileExt(doc.name || '');
        let iconClass = "fa-file-lines text-slate-600";
        let bgClass = "bg-white border-blue-200";

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
            <a href="${escUrl(fileSrc(doc)) || '#'}" ${fileSrc(doc).indexOf('data:') === 0 ? `download="${esc(doc.name)}"` : 'rel="noopener"'} target="_blank" class="p-1 text-blue-600 hover:text-blue-900 transition" title="เปิด / ดาวน์โหลด">
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
      renderDetailPhotos(task, photos);
      renderDetailDocs(docs);
      renderDetailVideos(videos);

      document.getElementById('detailModal').classList.remove('hidden');
    }

    function closeDetailModal() {
      state.currentDetailJobId = null;
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
        contractLine = `<span class="text-emerald-700 font-bold"><i class="fa-solid fa-flag-checkered mr-1"></i>ส่งมอบงานเรียบร้อย (${esc(task.Completion_Date || '-')})</span>`;
      } else if (daysRem !== null && daysRem < 0) {
        contractLine = `<span class="text-rose-700 font-bold"><i class="fa-solid fa-triangle-exclamation mr-1"></i>เลยกำหนดสัญญาแล้ว ${Math.abs(daysRem)} วัน (${esc(task.Contract_Expiry_Date)})</span>`;
      } else if (daysRem !== null && daysRem <= 7) {
        contractLine = `<span class="text-amber-800 font-bold"><i class="fa-solid fa-clock mr-1"></i>ครบกำหนดสัญญาในอีก ${daysRem} วัน (${esc(task.Contract_Expiry_Date)})</span>`;
      } else if (daysRem !== null) {
        contractLine = `<span class="text-slate-600"><i class="fa-regular fa-calendar-check mr-1"></i>ครบกำหนดสัญญา: ${esc(task.Contract_Expiry_Date)}</span>`;
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
        cell('วันอนุมัติ PO', task.PO_Approval_Date, 'fa-file-signature text-amber-600'),
        cell('วันปฏิบัติงาน (Target)', task.Target_Date, 'fa-calendar-day'),
        cell('วันครบกำหนดสัญญา', task.Contract_Expiry_Date, 'fa-calendar-xmark text-pts-800'),
        cell('วันที่งานเสร็จจริง', task.Completion_Date, 'fa-flag-checkered text-emerald-700'),
        `<div class="sm:col-span-2 p-3 rounded-2xl bg-slate-50 border border-slate-200">
           <p class="text-[10px] font-bold text-slate-500 uppercase tracking-wider mb-0.5">
             <i class="fa-solid fa-clipboard-list mr-1 text-slate-400"></i>รายละเอียดงาน
           </p>
           <p class="text-xs text-slate-800 whitespace-pre-wrap">${esc(task.Task_Detail || '-')}</p>
         </div>`
      ].join('');
    }

    function renderDetailPhotos(task, photos) {
      const container = document.getElementById('detailPhotosGallery');
      document.getElementById('detailPhotoCount').textContent = `(${photos.length} รูป)`;
      container.innerHTML = '';

      if (photos.length === 0) {
        container.innerHTML = `<p class="col-span-full text-xs text-slate-400 italic p-1">ยังไม่มีรูปภาพแนบในใบงานนี้</p>`;
        return;
      }

      photos.forEach((p, idx) => {
        const src = photoSrc(p);
        const div = document.createElement('div');
        div.className = "relative aspect-square rounded-xl overflow-hidden border border-slate-200 bg-slate-100 cursor-zoom-in group";
        div.onclick = () => openPhotoLightbox(photos, idx, task);
        div.innerHTML = `
          <img src="${escUrl(src)}" loading="lazy" alt="Site photo ${idx + 1}" class="w-full h-full object-cover group-hover:scale-110 transition duration-200">
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
      document.getElementById('createJobForm').reset();
      document.getElementById('newJobId').value = suggestNextJobId();

      const todayStr = new Date().toISOString().split('T')[0];
      document.getElementById('newPoApprovalDate').value = todayStr;
      document.getElementById('newTargetDate').value = todayStr;
      
      const futureDate = new Date();
      futureDate.setDate(futureDate.getDate() + 30);
      document.getElementById('newContractExpiryDate').value = futureDate.toISOString().split('T')[0];

      state.tempNewPhotos = [];
      state.tempNewDocs = [];
      state.tempNewVideos = [];
      state.tempNewDelivery = [];

      document.getElementById('newSiteContactPhone').value = '';
      document.getElementById('newSiteMapUrl').value = '';

      renderPhotosPreview('new');
      renderDocsPreview('new');
      renderVideosPreview('new');
      renderDeliveryPreview('new');

      document.getElementById('createModal').classList.remove('hidden');
    }

    function closeCreateModal() {
      document.getElementById('createModal').classList.add('hidden');
    }

    async function handleCreateJob(e) {
      e.preventDefault();

      const jobId = document.getElementById('newJobId').value.trim();
      if (!jobId) {
        alert("กรุณาระบุรหัสใบงาน");
        return;
      }
      if (state.tasks.some(t => String(t.Job_ID).trim() === jobId)) {
        alert(`รหัสใบงาน ${jobId} ถูกใช้ไปแล้ว กรุณาเปลี่ยนรหัสใหม่\n(รหัสว่างถัดไป: ${suggestNextJobId()})`);
        return;
      }

      const newTask = {
        Job_ID: jobId,
        Project_Name: document.getElementById('newProjectName').value,
        Sub_Department: document.getElementById('newSubDept').value,
        Technician_In_Charge: document.getElementById('newTechnician').value,
        Task_Detail: document.getElementById('newTaskDetail').value,
        Status: document.getElementById('newStatus').value,
        JSA_Completed: document.getElementById('newJsaCompleted').checked ? 'Yes' : 'No',
        Priority: document.getElementById('newPriority').value,
        Target_Date: document.getElementById('newTargetDate').value,
        PO_Approval_Date: document.getElementById('newPoApprovalDate').value,
        Contract_Expiry_Date: document.getElementById('newContractExpiryDate').value,
        Completion_Date: '',
        Site_Location: document.getElementById('newSiteLocation').value,
        Site_Contact_Phone: document.getElementById('newSiteContactPhone').value,
        Site_Map_Url: document.getElementById('newSiteMapUrl').value,
        Notes_Issues: `[${getNowFormatted()}] เปิดใบงานใหม่ในระบบ`,
        Updated_At: getNowFormatted(),
        Site_Photos: JSON.stringify(state.tempNewPhotos),
        Document_Files: JSON.stringify(state.tempNewDocs),
        Video_Files: JSON.stringify(state.tempNewVideos),
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

      setVal('editProjectName', task.Project_Name);
      setVal('editTechnician', task.Technician_In_Charge);
      setVal('editSubDept', task.Sub_Department || 'งานโครงการ');
      setVal('editPriority', task.Priority || 'Medium');
      setVal('editSiteLocation', task.Site_Location);
      setVal('editSiteContactPhone', task.Site_Contact_Phone);
      setVal('editSiteMapUrl', task.Site_Map_Url);
      setVal('editTaskDetail', task.Task_Detail);
      
      setVal('editPoApprovalDate', task.PO_Approval_Date);
      setVal('editContractExpiryDate', task.Contract_Expiry_Date);
      setVal('editCompletionDate', task.Completion_Date);

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
              <p class="text-[11px] text-emerald-700">วันที่อนุมัติ PO: ${task.PO_Approval_Date || '-'} | ครบสัญญา: ${task.Contract_Expiry_Date || '-'} | งานเสร็จจริง: ${task.Completion_Date || '-'}</p>
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
              <p class="text-[11px] text-rose-700">วันครบกำหนดสัญญา: ${task.Contract_Expiry_Date} | โปรดเร่งรัดดำเนินการส่งมอบงาน</p>
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
              <p class="text-[11px] text-amber-800">วันครบกำหนดสัญญา: ${task.Contract_Expiry_Date} (วันอนุมัติ PO: ${task.PO_Approval_Date || '-'})</p>
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
              <p class="text-[11px] text-slate-500">วันอนุมัติ PO: ${task.PO_Approval_Date || '-'} | ครบกำหนดสัญญา: ${task.Contract_Expiry_Date || '-'}</p>
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

      const getVal = (id, def = '') => {
        const el = document.getElementById(id);
        return el ? el.value : def;
      };

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

    // MODAL 4: API CONFIGURATION
    function openConfigModal() {
      document.getElementById('apiUrlInput').value = state.apiUrl || DEFAULT_LIVE_API_URL;
      syncConfigOpenLink();
      showConfigMsg('', '');
      const modeRadio = document.querySelector(`input[name="dbMode"][value="${state.dbMode}"]`);
      if (modeRadio) modeRadio.checked = true;
      document.getElementById('configModal').classList.remove('hidden');
    }

    function closeConfigModal() {
      document.getElementById('configModal').classList.add('hidden');
    }

    function syncConfigOpenLink() {
      const link = document.getElementById('configOpenLink');
      const url = normalizeApiUrl(document.getElementById('apiUrlInput').value);
      link.href = /^https:\/\/script\.google\.com\//.test(url) ? url : '#';
    }

    function showConfigMsg(text, kind) {
      const box = document.getElementById('configTestMsg');
      const skin = {
        ok:      'bg-emerald-50 text-emerald-900 border border-emerald-300',
        error:   'bg-rose-50 text-pts-900 border border-rose-300',
        warn:    'bg-amber-50 text-amber-900 border border-amber-300',
        pending: 'bg-slate-100 text-slate-700 border border-slate-300'
      }[kind];

      if (!text || !skin) {
        box.className = 'hidden p-3 rounded-xl text-xs font-medium whitespace-pre-line';
        box.textContent = '';
        return;
      }
      box.className = 'p-3 rounded-xl text-xs font-medium whitespace-pre-line ' + skin;
      box.textContent = text;
    }

    function useBuiltInApiUrl() {
      document.getElementById('apiUrlInput').value = DEFAULT_LIVE_API_URL;
      syncConfigOpenLink();
      showConfigMsg('ใส่ลิงก์ที่มากับไฟล์แล้ว กด "ทดสอบการเชื่อมต่อ" เพื่อตรวจสอบ', 'pending');
    }

    /**
     * Checks the URL that is currently TYPED in the box, not the saved one, so
     * a link can be verified before it is committed. Reports which stage failed
     * - shape, reachability, or reading the sheet - because each has a different
     * fix and the old dialog reported none of them.
     */
    async function testConnection() {
      const btn = document.getElementById('configTestBtn');
      const typed = normalizeApiUrl(document.getElementById('apiUrlInput').value);
      syncConfigOpenLink();

      if (!typed) {
        showConfigMsg('ยังไม่ได้ใส่ URL', 'error');
        return;
      }

      const shapeProblem = validateApiUrlShape(typed);
      if (shapeProblem && !typed.endsWith('/dev')) {
        showConfigMsg('❌ ' + shapeProblem, 'error');
        return;
      }

      btn.disabled = true;
      btn.classList.add('opacity-60');
      showConfigMsg('กำลังทดสอบ...', 'pending');

      // Point the transport at the typed URL for the duration of the test only
      const previousUrl = state.apiUrl;
      state.apiUrl = typed;
      try {
        const ping = await apiGetDetailed({ ping: '1' });
        if (!ping.json) {
          showConfigMsg('❌ ติดต่อ Apps Script ไม่ได้\n\n' + apiErrorText(ping.error) +
            '\n\nลองกด "เปิดลิงก์ในแท็บใหม่" — ถ้าเห็นหน้าล็อกอินหรือ "ไม่พบเพจ" แปลว่าต้อง Deploy ใหม่โดยตั้ง Who has access = Anyone', 'error');
          return;
        }

        const full = await apiGetDetailed({ fresh: '1' });
        if (!full.json || full.json.status !== 'success' || !Array.isArray(full.json.data)) {
          showConfigMsg('⚠️ ต่อถึง Apps Script ได้ แต่อ่านชีตไม่สำเร็จ\n\n' +
            (full.json && full.json.message ? full.json.message : apiErrorText(full.error)) +
            '\n\nมักเกิดจากยังไม่ได้รัน authorizeDriveAccess หรือแท็บชีตไม่ได้ชื่อ Engineering_Tasks', 'warn');
          return;
        }

        const rows = full.json.data.length;
        showConfigMsg('✅ เชื่อมต่อสำเร็จ\n\n' +
          'จำนวนใบงานในชีต: ' + rows + ' รายการ' +
          (rows === 0 ? ' (ชีตใหม่ที่ยังว่าง — พร้อมให้ทดสอบบันทึกข้อมูล)' : '') +
          '\nrevision: ' + (typeof full.json.rev !== 'undefined' ? full.json.rev : '-') +
          (shapeProblem ? '\n\n⚠️ ' + shapeProblem : '') +
          '\n\nกด "บันทึกการตั้งค่า" เพื่อใช้ลิงก์นี้', 'ok');
      } finally {
        state.apiUrl = previousUrl;
        btn.disabled = false;
        btn.classList.remove('opacity-60');
      }
    }

    function saveConfigSettings() {
      const mode = document.querySelector('input[name="dbMode"]:checked').value;
      const url = normalizeApiUrl(document.getElementById('apiUrlInput').value);

      if (mode === 'live' && !url) {
        alert("กรุณากรอก Google Apps Script Web App URL ก่อนเปิดใช้งาน Live Mode");
        return;
      }

      if (mode === 'live') {
        const shapeProblem = validateApiUrlShape(url);
        if (shapeProblem && !confirm('⚠️ ' + shapeProblem + '\n\nต้องการบันทึกลิงก์นี้ต่อไปหรือไม่?')) return;
      }

      state.dbMode = mode;
      state.apiUrl = url;
      state.lastError = null;

      localStorage.setItem('pts_db_mode', mode);
      localStorage.setItem(API_URL_KEY, url);
      // Records the shipped default in force at the moment the user chose this
      // URL, so a future update only overrides links nobody customised.
      localStorage.setItem(API_URL_ADOPTED_KEY, DEFAULT_LIVE_API_URL);

      updateStatusBadge();
      closeConfigModal();

      if (mode === 'live') {
        fetchTasks();
      } else {
        loadState();
      }
    }

    function loadSampleDataPrompt() {
      if (confirm("ต้องการรีเซ็ตข้อมูล Mock Sandbox ให้เป็นข้อมูลตัวอย่างเริ่มต้น 6 รายการหรือไม่?")) {
        state.tasks = [...INITIAL_SAMPLE_TASKS];
        saveMockStorage();
        renderApp();
        closeConfigModal();
      }
    }

    // HELPER UTILS
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

      document.getElementById('deliveryJobIdBadge').textContent = task.Job_ID;
      document.getElementById('deliverySubtitle').textContent = `${task.Project_Name || '-'} · ${task.Sub_Department || '-'}`;

      // Form Sheet Data
      document.getElementById('formSheetJobId').textContent = task.Job_ID;
      document.getElementById('formSheetDate').textContent = `วันที่ส่งมอบ: ${task.Completion_Date || task.Target_Date || getNowFormatted().split(' ')[0]}`;
      document.getElementById('formSheetProject').textContent = task.Project_Name || '-';
      document.getElementById('formSheetSubDept').textContent = task.Sub_Department || '-';
      document.getElementById('formSheetLocation').textContent = task.Site_Location || '-';
      document.getElementById('formSheetContact').textContent = task.Site_Contact_Phone || '-';
      document.getElementById('formSheetPoDate').textContent = task.PO_Approval_Date || '-';
      document.getElementById('formSheetExpiryDate').textContent = task.Contract_Expiry_Date || '-';
      document.getElementById('formSheetTaskDetail').textContent = task.Task_Detail || 'ดำเนินการติดตั้ง / ตรวจสอบระบบตามสัญญาเสร็จสิ้นสมบูรณ์';
      
      const isJsaOk = (task.JSA_Completed || '').toLowerCase() === 'yes';
      document.getElementById('formSheetJsa').innerHTML = isJsaOk
        ? '<i class="fa-solid fa-shield-check mr-1 text-emerald-600"></i>ผ่านการประเมินความปลอดภัยเรียบร้อย (JSA Approved)'
        : '<i class="fa-solid fa-triangle-exclamation mr-1 text-amber-600"></i>ยังไม่ได้รับการรับรอง JSA (Pending)';
      
      document.getElementById('formSheetCompletionDate').textContent = task.Completion_Date || task.Target_Date || '-';
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
            <p class="text-[11px] text-slate-500 mt-1">สามารถกดปุ่ม <span class="font-semibold text-blue-700">+ แนบไฟล์ลงระบบ</span> ด้านบนเพื่อเลือกไฟล์ PDF, สแกน หรือรูปภาพมาบันทึกเก็บไว้ได้ทันที</p>
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

      if (deliveryFiles.length > 0) {
        switchDeliveryTab('files');
      } else {
        switchDeliveryTab('form');
      }

      document.getElementById('deliveryModal').classList.remove('hidden');
    }

    function closeDeliveryModal() {
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
      if (jobId) openEditModal(jobId);
    }

    async function handleDeliveryModalUpload(event) {
      const jobId = state.currentDeliveryJobId;
      if (!jobId) return;
      const task = state.tasks.find(t => t.Job_ID === jobId);
      if (!task) return;

      const input = event.target;
      const files = Array.from(input.files || []);
      input.value = '';
      if (files.length === 0) return;

      const results = await Promise.all(files.map(file => readFileAsDataUrl(file).then(dataUrl => (
        dataUrl ? {
          name: file.name,
          size: formatBytes(file.size),
          type: file.type || getFileExt(file.name),
          dataUrl: dataUrl
        } : null
      ))));

      const usable = results.filter(Boolean);
      if (usable.length === 0) return;

      const current = getTaskAttachments(task, 'Delivery_Doc');
      usable.forEach(item => current.push(item));
      task.Delivery_Doc = JSON.stringify(current);

      await saveTaskToDb(task, false);
      openDeliveryDocModal(jobId);
      showSyncToast('บันทึกเอกสารใบส่งมอบงานเรียบร้อย');
    }
