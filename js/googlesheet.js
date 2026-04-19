/**
 * googlesheet.js — โหลดรายชื่อจาก Google Sheet
 * ==============================================
 * ใช้ CSV export URL (Public Sheet) ไม่ต้องใช้ API Key
 */

const GoogleSheet = (() => {

  /** สร้าง URL สำหรับดึง CSV จาก Google Sheet */
  function buildCsvUrl(sheetId, gid = '0') {
    return `https://docs.google.com/spreadsheets/d/${sheetId}/export?format=csv&gid=${gid}`;
  }

  /**
   * แปลง CSV text เป็น Array of Objects
   * @param {string} csvText - Raw CSV string
   * @returns {Array} - array ของแถวข้อมูล (ข้ามแถวแรกที่เป็น header)
   */
  function parseCsv(csvText) {
    const lines = csvText.trim().split('\n');
    if (lines.length < 2) return [];

    // ข้ามแถวแรก (header)
    return lines.slice(1).map(line => {
      // จัดการ field ที่มี comma ใน quotes
      const cols = [];
      let inQuote = false;
      let cell = '';
      for (let i = 0; i < line.length; i++) {
        const ch = line[i];
        if (ch === '"') {
          inQuote = !inQuote;
        } else if (ch === ',' && !inQuote) {
          cols.push(cell.trim());
          cell = '';
        } else {
          cell += ch;
        }
      }
      cols.push(cell.trim());
      return cols;
    }).filter(row => row.some(cell => cell.length > 0)); // กรองแถวว่าง
  }

  /**
   * แปลง CSV row เป็น Officer object
   * @param {Array} row - array ของ cell values
   * @returns {Object} officer object
   */
  function rowToOfficer(row) {
    const c = CONFIG.COLUMNS;
    return {
      rank: (row[c.RANK]       || '').trim(),
      name: (row[c.NAME]       || '').trim(),
      dept: (row[c.DEPARTMENT] || '').trim(),
      id:   (row[c.ID]         || '').trim(),
    };
  }

  /**
   * โหลดรายชื่อตำรวจจาก Google Sheet
   * @returns {Promise<{officers: Array, source: string}>}
   */
  async function loadOfficers() {
    const sheetId = CONFIG.SHEET_ID;

    // ถ้ายังไม่ได้ตั้ง Sheet ID → ใช้ Demo data
    if (!sheetId || sheetId === 'YOUR_GOOGLE_SHEET_ID_HERE') {
      console.info('[GoogleSheet] Using demo data (no Sheet ID configured)');
      return {
        officers: CONFIG.DEMO_OFFICERS,
        source: 'demo'
      };
    }

    const url = buildCsvUrl(sheetId, CONFIG.SHEET_GID);
    console.info('[GoogleSheet] Fetching:', url);

    try {
      const resp = await fetch(url, { cache: 'no-cache' });
      if (!resp.ok) throw new Error(`HTTP ${resp.status}: ${resp.statusText}`);

      const csvText = await resp.text();
      const rows    = parseCsv(csvText);
      const officers = rows
        .map(rowToOfficer)
        .filter(o => o.name.length > 0); // กรองแถวที่ไม่มีชื่อ

      if (officers.length === 0) throw new Error('ไม่พบข้อมูลในชีท');

      console.info(`[GoogleSheet] Loaded ${officers.length} officers`);
      return { officers, source: 'sheet' };

    } catch (err) {
      console.warn('[GoogleSheet] Load failed, using demo data:', err.message);
      return {
        officers: CONFIG.DEMO_OFFICERS,
        source: 'demo',
        error: err.message
      };
    }
  }

  // Public API
  return { loadOfficers };

})();
