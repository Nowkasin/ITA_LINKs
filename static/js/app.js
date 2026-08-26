const DEPARTMENTS = window.DEPARTMENTS || [];
const METHODS = window.METHODS || [];
const REASONS = window.REASONS || [];
const PURCHASING_UNITS = window.PURCHASING_UNITS || [];
const THAI_MONTHS = ["มกราคม","กุมภาพันธ์","มีนาคม","เมษายน","พฤษภาคม","มิถุนายน","กรกฎาคม","สิงหาคม","กันยายน","ตุลาคม","พฤศจิกายน","ธันวาคม"];

// เก็บผลการค้นหาล่าสุดไว้ในตัวแปรนี้ เพื่อใช้เปิดฟอร์มแก้ไขโดยไม่ต้อง fetch ใหม่
let CURRENT_ROWS = [];
let DELETE_TARGET_ID = null;

// state ของ unit type filter
let ACTIVE_GROUP_REPORT = '';   // ประเภทหน่วยงาน (บนสุด) — ไม่มี "ทั้งหมด" แล้ว ต้องเป็นค่าใดค่าหนึ่งจาก Group_Report_ms เสมอ
const GROUP_REPORT_LIST = window.GROUP_REPORT_LIST || []; // จาก Group_Report_ms: [{value, label}]

const K2_PARAM_KEYS = {
    dept: 'CUR_INDEPT_LNAME_TH',   // ส่วนงาน (ชื่อหน่วยงานภาษาไทย)
    fiscalYear: 'Year',            // ปีงบประมาณ
    month: 'Month',                // เดือน (1-12)
    reportGroup: 'Group_Report',   // กลุ่มรายงาน
    purchasingUnit: 'PlantName',   // หน่วยจัดซื้อ (ชื่อ Plant)
};

let K2_DEPT_OVERRIDE = null;   // id หน่วยงานที่ resolve ได้จากชื่อที่ K2 ส่งมา (ถ้ามี)
let K2_REPORT_GROUP = null;    // กลุ่มรายงานดิบที่ K2 ส่งมา ส่งต่อให้ backend กรอง

// หมายเหตุ: DEPT_MAP_PLANT/SNO_DEPT_IDS (เวอร์ชันก่อนหน้า) ถูกเอาออกแล้ว — ตอนนี้กรองตามกลุ่มหน่วยงาน
// ใช้ Group_Report ส่งไป backend ให้ SQL กรองให้ตรงๆ ผ่าน fetch_plant_list_by_group() แทน ไม่ต้อง hack ฝั่ง JS
// UNIT TYPE FILTER (ปุ่มไดนามิกจาก Group_Report_ms)

function initGroupReportFilter() {
    // สร้างปุ่ม "ประเภทหน่วยงาน" แบบไดนามิกจาก Group_Report_ms — ไม่มีปุ่ม "ทั้งหมด" แล้ว
    // ต้องมีค่าใดค่าหนึ่งเลือกอยู่เสมอ -> ตั้งค่าเริ่มต้นเป็นแถวแรกของตาราง
    const container = document.getElementById('groupReportBtnGroup');
    if (container) {
        container.innerHTML = '';
        GROUP_REPORT_LIST.forEach((g, idx) => {
            const btn = document.createElement('button');
            btn.type = 'button';
            btn.className = 'unit-btn' + (idx === 0 ? ' active' : '');
            btn.textContent = g.label;
            btn.dataset.groupReport = g.value;
            btn.addEventListener('click', () => {
                ACTIVE_GROUP_REPORT = g.value;
                container.querySelectorAll('.unit-btn').forEach(x => x.classList.remove('active'));
                btn.classList.add('active');

                applyPurchasingUnitGroupFilter();
                updateUnitSummary();
            });
            container.appendChild(btn);
        });
        if (GROUP_REPORT_LIST.length > 0) {
            ACTIVE_GROUP_REPORT = GROUP_REPORT_LIST[0].value;
        }
    }

    applyPurchasingUnitGroupFilter();
    updateUnitSummary();
}

// อัปเดตช่องแสดงผล "ส่วนงาน :" ให้ตรงกับหน่วยงาน/คณะที่เลือกอยู่ตอนนี้
// (มาแทนที่ dropdown "หน่วยงาน (ส่วนจัดซื้อ)" เดิม — ตอนนี้เป็นช่องแสดงผลอย่างเดียว)
function updateUnitSummary() {
    const nameEl = document.getElementById('deptDisplay');
    const codeEl = document.getElementById('deptDisplayCode');
    if (!nameEl || !codeEl) return;

    // ถ้า K2 ส่งหน่วยงานมาแบบ override ให้แสดงตามนั้นก่อนเสมอ ไม่ให้ปุ่ม filter ทับ
    if (K2_DEPT_OVERRIDE !== null) return;

    if (ACTIVE_GROUP_REPORT !== '') {
        const found = GROUP_REPORT_LIST.find(g => g.value === ACTIVE_GROUP_REPORT);
        nameEl.value = found ? found.label : ('กลุ่มรายงาน ' + ACTIVE_GROUP_REPORT);
        codeEl.textContent = '';
        return;
    }

    nameEl.value = '-- ทุกหน่วยงาน --';
    codeEl.textContent = '';
}

function getUnitFilterParams() {
    // ตอนนี้ไม่มี dropdown "เลือกหน่วยงานเจาะจง" แล้ว (เอาออกตามที่ขอ)
    // การกรองหน่วยงาน/กลุ่มทำผ่าน ACTIVE_GROUP_REPORT (ส่งเป็น Group_Report ใน currentParams()) เท่านั้น
    // Search_Dept_ID จะถูกส่งก็ต่อเมื่อ K2 ส่ง dept ที่ resolve เป็น id ได้มาโดยตรง (ดู K2_DEPT_OVERRIDE ใน currentParams())
    return {};
}

// FILTER DROPDOWNS (สำหรับฟอร์มค้นหาด้านบน)

function buildYearOptions() {
    const sel = document.getElementById('fiscalYear');
    const now = new Date();
    const currentBE = now.getFullYear() + 543;
    for (let y = currentBE + 1; y >= currentBE - 5; y--) {
        const opt = document.createElement('option');
        opt.value = y;
        opt.textContent = y;
        if (y === currentBE) opt.selected = true;
        sel.appendChild(opt);
    }
}

function buildMonthOptions() {
    const sel = document.getElementById('month');
    THAI_MONTHS.forEach((name, idx) => {
        const opt = document.createElement('option');
        opt.value = idx + 1;
        opt.textContent = name;
        sel.appendChild(opt);
    });
}

// เพิ่มใหม่: build dropdown "หน่วยจัดซื้อ" จาก window.PURCHASING_UNITS — เดิมไม่เคยถูกเรียกเลย
function buildPurchasingUnitOptions() {
    const sel = document.getElementById('purchasingUnit');
    if (!sel) return;
    PURCHASING_UNITS.forEach(u => {
        const opt = document.createElement('option');
        opt.value = u.id;
        opt.textContent = u.name;
        opt.dataset.group = u.group; // '0' | '1' | '2' จาก Plant_ms.Group_Report — ใช้กรองตามปุ่ม "ประเภทหน่วยงาน"
        sel.appendChild(opt);
    });
}

// ซ่อน/แสดง option ใน dropdown "หน่วยจัดซื้อ" ให้ตรงกับปุ่ม "ประเภทหน่วยงาน" ที่เลือกอยู่ (ACTIVE_GROUP_REPORT)
// 0 = มหาวิทยาลัย -> แสดงทุก Plant, 1 = สนอ. -> เฉพาะ Plant กลุ่ม 1, 2 = คณะ -> เฉพาะ Plant กลุ่ม 2
function applyPurchasingUnitGroupFilter() {
    const sel = document.getElementById('purchasingUnit');
    if (!sel) return;

    Array.from(sel.options).forEach(opt => {
        if (!opt.value) return; // option "ทุกหน่วยจัดซื้อ" ไม่ต้องซ่อน
        if (isUniversityWideGroup(ACTIVE_GROUP_REPORT)) {
            opt.hidden = false;
        } else {
            opt.hidden = (opt.dataset.group !== ACTIVE_GROUP_REPORT);
        }
    });

    // ถ้า option ที่เลือกอยู่ถูกซ่อนไป (สลับกลุ่มแล้วของเดิมไม่อยู่ในกลุ่มใหม่) ให้ reset กลับเป็น "ทุกหน่วยจัดซื้อ"
    const selectedOpt = sel.options[sel.selectedIndex];
    if (selectedOpt && selectedOpt.hidden) sel.value = '';
}

// เช็คว่าค่า Group_Report ที่เลือกคือแถว "มหาวิทยาลัยศรีนครินทรวิโรฒ" (ระดับทั้งหมด) หรือไม่
// ยืนยันแล้วว่ารหัสจริงใน Group_Report_ms คือ: 0 = มหาวิทยาลัย (ทั้งหมด), 1 = สำนักงานอธิการบดี, 2 = คณะ/สถาบัน/สำนัก
// เช็คค่า '0' เป็นหลัก และเผื่อ fallback เช็คจาก label ด้วย เผื่อรหัสไม่ตรงกับที่คาดไว้ในบาง environment
function isUniversityWideGroup(value) {
    if (String(value) === '0') return true;
    const found = GROUP_REPORT_LIST.find(g => g.value === value);
    return !!(found && found.label && found.label.indexOf('มหาวิทยาลัย') !== -1);
}

function currentParams() {
    const year = document.getElementById('fiscalYear').value;
    const month = document.getElementById('month').value;
    const purchasingUnit = document.getElementById('purchasingUnit').value;
    const params = new URLSearchParams();
    if (year) params.set('FiscalYear', year);
    if (month) params.set('Month', month);
    if (purchasingUnit) params.set('PlantName_id', purchasingUnit);

    // K2 ส่งกลุ่มรายงานมา override ก่อนเสมอ ถ้าไม่มีค่อยใช้ปุ่มที่ผู้ใช้เลือกในหน้าเว็บ
    // ยกเว้นกลุ่ม "มหาวิทยาลัย..." (ทั้งหมด) -> ไม่ส่ง Group_Report เลย เพื่อให้ได้ Plant ครบทุกตัว ไม่ถูกกรอง
    const groupReport = K2_REPORT_GROUP || (ACTIVE_GROUP_REPORT || null);
    if (groupReport && !isUniversityWideGroup(groupReport)) {
        params.set('Group_Report', groupReport);
    }

    const unitFilter = K2_DEPT_OVERRIDE
        ? { Search_Dept_ID: K2_DEPT_OVERRIDE }
        : getUnitFilterParams();

    if (unitFilter.Search_Dept_ID) {
        params.set('Search_Dept_ID', unitFilter.Search_Dept_ID);
    }

    return params;
}

function setDownloadLinks(enabled) {
    const params = currentParams();
    const pdfBtn = document.getElementById('downloadPdf');
    const csvBtn = document.getElementById('downloadCsv');
    if (enabled) {
        pdfBtn.href = '/procurement-report.pdf?' + params.toString();
        csvBtn.href = '/procurement-report.xlsx?' + params.toString();
        pdfBtn.removeAttribute('aria-disabled');
        csvBtn.removeAttribute('aria-disabled');
    } else {
        pdfBtn.href = '#';
        csvBtn.href = '#';
        pdfBtn.setAttribute('aria-disabled', 'true');
        csvBtn.setAttribute('aria-disabled', 'true');
    }
}

// SEARCH + TABLE
async function runSearch() {
    const resultArea = document.getElementById('resultArea');
    const meta = document.getElementById('resultMeta');
    const btn = document.getElementById('searchBtn');

    const year = document.getElementById('fiscalYear').value;
    if (!year) {
        resultArea.innerHTML = '<div class="error-state">กรุณาเลือกปีงบประมาณก่อนค้นหา</div>';
        return;
    }

    btn.disabled = true;
    btn.textContent = 'กำลังค้นหา...';
    setDownloadLinks(false);
    meta.textContent = '';
    resultArea.innerHTML = '<div class="loading-state"><div class="spinner"></div>กำลังดึงข้อมูล...</div>';

    try {
        const params = currentParams();
        const res = await fetch('/api/procurement-data?' + params.toString());
        if (!res.ok) throw new Error('เรียกข้อมูลไม่สำเร็จ (HTTP ' + res.status + ')');
        const data = await res.json();
        CURRENT_ROWS = data.rows || [];
        renderTable(CURRENT_ROWS);
        setDownloadLinks(true);
    } catch (err) {
        resultArea.innerHTML = '<div class="error-state">เกิดข้อผิดพลาด: ' + (err.message || 'ไม่สามารถดึงข้อมูลได้') + '</div>';
        setDownloadLinks(false);
    } finally {
        btn.disabled = false;
        btn.textContent = 'ค้นหา';
    }
}

function renderTable(rows) {
    const resultArea = document.getElementById('resultArea');
    const meta = document.getElementById('resultMeta');
    meta.innerHTML = '<span class="badge-count">' + rows.length + ' รายการ</span>';

    if (!rows.length) {
        resultArea.innerHTML = '<div class="empty-state">ไม่พบข้อมูลตามเงื่อนไขที่เลือก</div>';
        return;
    }

    let html = '<div class="table-scroll"><table><thead><tr>' +
        '<th>ลำดับ</th><th>หน่วยงาน</th><th>งานที่จัดซื้อ/จัดจ้าง</th><th>วงเงิน</th><th>ราคากลาง</th>' +
        '<th>วิธีซื้อ/จ้าง</th><th>ผู้เสนอราคา/ผู้ชนะ</th><th>เหตุผลที่คัดเลือก</th><th>เลขที่เอกสาร</th>' +
        '<th>จัดการ</th>' +
        '</tr></thead><tbody>';

    rows.forEach(r => {
        // ปุ่มลบแสดงเฉพาะ admin (สอดคล้องกับ @admin_required ฝั่ง backend)
        const deleteBtn = window.IS_ADMIN
            ? '<button class="btn-danger btn-sm" type="button" onclick="openDeleteModal(' + r.id + ')">ลบ</button>'
            : '';
        html += '<tr>' +
            '<td class="center">' + escapeHtml(r.seq) + '</td>' +
            '<td>' + escapeHtml(r.dept_name) + '</td>' +
            '<td>' + escapeHtml(r.work) + '</td>' +
            '<td class="num">' + escapeHtml(r.budget) + '</td>' +
            '<td class="num">' + escapeHtml(r.median) + '</td>' +
            '<td class="center">' + escapeHtml(r.method) + '</td>' +
            '<td>' + escapeHtml(r.winner) + '</td>' +
            '<td>' + escapeHtml(r.reason) + '</td>' +
            '<td>' + escapeHtml(r.refdoc) + '</td>' +
            '<td class="center">' +
                '<div class="row-actions">' +
                    '<button class="btn-outline btn-sm" type="button" onclick="openEditModal(' + r.id + ')">แก้ไข</button>' +
                    deleteBtn +
                '</div>' +
            '</td>' +
            '</tr>';
    });

    html += '</tbody></table></div>';
    resultArea.innerHTML = html;
}

function escapeHtml(str) {
    if (str === null || str === undefined) return '-';
    return String(str)
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;');
}

// MODAL: เพิ่ม/แก้ไขข้อมูล
function buildFormYearOptions() {
    const sel = document.getElementById('f_fiscal_year');
    sel.innerHTML = '';
    const now = new Date();
    const currentBE = now.getFullYear() + 543;
    for (let y = currentBE + 1; y >= currentBE - 5; y--) {
        const opt = document.createElement('option');
        opt.value = y;
        opt.textContent = y;
        if (y === currentBE) opt.selected = true;
        sel.appendChild(opt);
    }
}

function buildFormDeptOptions() {
    const sel = document.getElementById('f_dept_id');
    sel.innerHTML = '<option value="">-- เลือกหน่วยงาน --</option>';
    DEPARTMENTS.forEach(d => {
        const opt = document.createElement('option');
        opt.value = d.id;
        opt.textContent = d.name;
        sel.appendChild(opt);
    });
}

function buildFormMethodOptions() {
    const sel = document.getElementById('f_method_id');
    sel.innerHTML = '<option value="">-- เลือกวิธีซื้อ/จ้าง --</option>';
    METHODS.forEach(m => {
        const opt = document.createElement('option');
        opt.value = m.id;
        opt.textContent = m.name;
        sel.appendChild(opt);
    });
}

function buildFormReasonOptions() {
    const sel = document.getElementById('f_reason_id');
    sel.innerHTML = '<option value="">-- เลือกเหตุผลที่คัดเลือก --</option>';
    REASONS.forEach(r => {
        const opt = document.createElement('option');
        opt.value = r.id;
        opt.textContent = r.name;
        sel.appendChild(opt);
    });
}

// เพิ่มใหม่: build dropdown "หน่วยจัดซื้อ" ในฟอร์มเพิ่ม/แก้ไข — เดิมไม่เคยถูกเรียกเลย ทำให้ #f_purchasing_unit_id ว่างตลอด
function buildFormPurchasingUnitOptions() {
    const sel = document.getElementById('f_purchasing_unit_id');
    if (!sel) return;
    sel.innerHTML = '<option value="">-- เลือกหน่วยจัดซื้อ --</option>';
    PURCHASING_UNITS.forEach(u => {
        const opt = document.createElement('option');
        opt.value = u.id;
        opt.textContent = u.name;
        sel.appendChild(opt);
    });
}

function clearForm() {
    document.getElementById('f_id').value = '';
    document.getElementById('f_work').value = '';
    document.getElementById('f_budget').value = '';
    document.getElementById('f_median').value = '';
    document.getElementById('f_bidder_list').value = '';
    document.getElementById('f_winner').value = '';
    document.getElementById('f_refdoc').value = '';
    const contractDateEl = document.getElementById('f_contract_date');
    if (contractDateEl) contractDateEl.value = '';
    document.getElementById('f_dept_id').value = '';
    document.getElementById('f_method_id').value = '';
    document.getElementById('f_reason_id').value = '';
    const purchasingUnitEl = document.getElementById('f_purchasing_unit_id');
    if (purchasingUnitEl) purchasingUnitEl.value = '';
    hideFormError();
}

function showFormError(message) {
    const box = document.getElementById('formError');
    box.textContent = message;
    box.hidden = false;
}

function hideFormError() {
    const box = document.getElementById('formError');
    box.hidden = true;
    box.textContent = '';
}

function openFormModal() {
    document.getElementById('formModalOverlay').classList.add('open');
}

function closeFormModal() {
    document.getElementById('formModalOverlay').classList.remove('open');
}

function openAddModal() {
    clearForm();
    document.getElementById('modalTitle').textContent = 'เพิ่มข้อมูลจัดซื้อจัดจ้าง';
    buildFormYearOptions();
    openFormModal();
}

// เปิดฟอร์มแก้ไข โดยดึงข้อมูลจาก CURRENT_ROWS ที่ค้นหามาแล้ว (ไม่ fetch ซ้ำ)
function openEditModal(id) {
    const row = CURRENT_ROWS.find(r => r.id === id);
    if (!row) {
        alert('ไม่พบข้อมูลรายการนี้ กรุณาค้นหาใหม่อีกครั้ง');
        return;
    }

    clearForm();
    document.getElementById('modalTitle').textContent = 'แก้ไขข้อมูลจัดซื้อจัดจ้าง';
    buildFormYearOptions();

    document.getElementById('f_id').value = row.id;
    document.getElementById('f_fiscal_year').value = row.fiscal_year || '';
    document.getElementById('f_dept_id').value = row.dept_id || '';
    document.getElementById('f_work').value = row.work === '-' ? '' : row.work;
    document.getElementById('f_budget').value = parseNumberOrEmpty(row.budget);
    document.getElementById('f_median').value = parseNumberOrEmpty(row.median);
    document.getElementById('f_method_id').value = row.method_id || '';
    document.getElementById('f_reason_id').value = row.reason_id || '';
    document.getElementById('f_bidder_list').value = row.bidder_list === '-' ? '' : row.bidder_list;
    document.getElementById('f_winner').value = row.winner === '-' ? '' : row.winner;
    document.getElementById('f_refdoc').value = row.refdoc === '-' ? '' : row.refdoc;
    const contractDateEl = document.getElementById('f_contract_date');
    if (contractDateEl) contractDateEl.value = row.contract_date
        ? row.contract_date.split(' ')[0]
        : '';

    openFormModal();
}

// ตัวเลขที่ผ่าน format_currency มาแล้วจะมี comma เช่น "5,000" ต้องลอกออกก่อนใส่ใน input type=number
function parseNumberOrEmpty(text) {
    if (!text || text === '-') return '';
    const cleaned = String(text).replace(/,/g, '');
    const num = parseFloat(cleaned);
    return isNaN(num) ? '' : num;
}

async function saveForm() {
    hideFormError();

    const id = document.getElementById('f_id').value;
    const fiscalYear = document.getElementById('f_fiscal_year').value;
    const work = document.getElementById('f_work').value.trim();
    const deptId = document.getElementById('f_dept_id').value;

    if (!work) {
        showFormError('กรุณากรอกงานที่จัดซื้อ/จัดจ้าง');
        return;
    }
    if (!fiscalYear) {
        showFormError('กรุณาเลือกปีงบประมาณ');
        return;
    }
    if (!deptId) {
        showFormError('กรุณาเลือกหน่วยงาน');
        return;
    }

    const payload = {
        id: id ? parseInt(id, 10) : null,
        fiscal_year: parseInt(fiscalYear, 10),
        dept_id: deptId,
        work: work,
        budget: emptyToNull(document.getElementById('f_budget').value),
        median: emptyToNull(document.getElementById('f_median').value),
        method_id: emptyToNull(document.getElementById('f_method_id').value),
        reason_id: emptyToNull(document.getElementById('f_reason_id').value),
        bidder_list: document.getElementById('f_bidder_list').value.trim() || null,
        winner: document.getElementById('f_winner').value.trim() || null,
        refdoc: document.getElementById('f_refdoc').value.trim() || null,
        contract_date: (document.getElementById('f_contract_date') || {}).value || null,
        use_or_not: true,
    };

    const saveBtn = document.getElementById('modalSaveBtn');
    saveBtn.disabled = true;
    saveBtn.textContent = 'กำลังบันทึก...';

    try {
        const res = await fetch('/api/procurement-data', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(payload),
        });
        const data = await res.json();

        if (!res.ok) {
            showFormError(data.error || 'บันทึกข้อมูลไม่สำเร็จ');
            return;
        }

        closeFormModal();
        // บันทึกสำเร็จ -> ค้นหาใหม่เพื่อให้ตารางอัปเดตข้อมูลล่าสุด
        await runSearch();
    } catch (err) {
        showFormError('เกิดข้อผิดพลาดในการเชื่อมต่อ: ' + (err.message || ''));
    } finally {
        saveBtn.disabled = false;
        saveBtn.textContent = 'บันทึก';
    }
}

function emptyToNull(value) {
    if (value === '' || value === null || value === undefined) return null;
    return value;
}

// MODAL: ลบข้อมูล
function openDeleteModal(id) {
    DELETE_TARGET_ID = id;
    document.getElementById('deleteModalOverlay').classList.add('open');
}

function closeDeleteModal() {
    DELETE_TARGET_ID = null;
    document.getElementById('deleteModalOverlay').classList.remove('open');
}

async function confirmDelete() {
    if (!DELETE_TARGET_ID) return;

    const btn = document.getElementById('deleteConfirmBtn');
    btn.disabled = true;
    btn.textContent = 'กำลังลบ...';

    try {
        const res = await fetch('/api/procurement-data/' + DELETE_TARGET_ID, {
            method: 'DELETE',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({}),
        });
        const data = await res.json();

        if (!res.ok) {
            alert(data.error || 'ลบข้อมูลไม่สำเร็จ');
            return;
        }

        closeDeleteModal();
        await runSearch();
    } catch (err) {
        alert('เกิดข้อผิดพลาดในการเชื่อมต่อ: ' + (err.message || ''));
    } finally {
        btn.disabled = false;
        btn.textContent = 'ลบข้อมูล';
    }
}

// K2 INTEGRATION — รับค่าที่ส่งมาผ่าน URL query string
function findByName(list, name) {
    if (!list || !name) return null;
    const target = name.trim();
    return list.find(item => (item.name || '').trim() === target) || null;
}

function loadParamsFromK2() {
    const qs = new URLSearchParams(window.location.search);
    const deptName = qs.get(K2_PARAM_KEYS.dept);
    const year     = qs.get(K2_PARAM_KEYS.fiscalYear);
    const month    = qs.get(K2_PARAM_KEYS.month);
    const group    = qs.get(K2_PARAM_KEYS.reportGroup);
    const unitName = qs.get(K2_PARAM_KEYS.purchasingUnit);

    if (!deptName && !year && !month && !group && !unitName) return; // ไม่มี query จาก K2 มา ไม่ต้องทำอะไร

    if (year)  document.getElementById('fiscalYear').value = year;
    if (month) document.getElementById('month').value = month;

    if (unitName) {
        const foundUnit = findByName(PURCHASING_UNITS, unitName);
        if (foundUnit) {
            document.getElementById('purchasingUnit').value = foundUnit.id;
        } else {
            console.warn('ไม่พบหน่วยจัดซื้อที่ตรงกับชื่อจาก K2:', unitName);
        }
    }

    if (deptName) {
        const foundDept = findByName(DEPARTMENTS, deptName);
        if (foundDept) {
            K2_DEPT_OVERRIDE = foundDept.id;
            document.getElementById('deptDisplay').value = foundDept.name;
        } else {
            // หา id ไม่เจอ แต่ยังส่งชื่อดิบให้ backend ลอง resolve เองอีกทีผ่าน CUR_INDEPT_LNAME_TH
            document.getElementById('deptDisplay').value = deptName;
            console.warn('ไม่พบหน่วยงานที่ตรงกับชื่อจาก K2 ในรายการฝั่งหน้าเว็บ:', deptName);
        }
    }

    if (group) {
        K2_REPORT_GROUP = group;
        ACTIVE_GROUP_REPORT = group;
        const container = document.getElementById('groupReportBtnGroup');
        if (container) {
            container.querySelectorAll('.unit-btn').forEach(b => {
                b.classList.toggle('active', b.dataset.groupReport === group);
            });
        }
        applyPurchasingUnitGroupFilter();
    }

    // ค้นหาอัตโนมัติเมื่อมีปีงบมาครบ (API บังคับต้องมีปีงบ)
    if (year) runSearch();
}

// INIT
document.addEventListener('DOMContentLoaded', () => {
    buildYearOptions();
    buildMonthOptions();
    buildPurchasingUnitOptions();
    buildFormDeptOptions();
    buildFormMethodOptions();
    buildFormReasonOptions();
    buildFormPurchasingUnitOptions();
    initGroupReportFilter();

    // ต้องเรียกหลัง build dropdown ทั้งหมดแล้วเท่านั้น เพราะต้องเซ็ตค่าลง dropdown ที่มี option ครบแล้ว
    loadParamsFromK2();

    document.getElementById('searchBtn').addEventListener('click', runSearch);
    document.getElementById('addBtn').addEventListener('click', openAddModal);

    document.getElementById('modalCloseBtn').addEventListener('click', closeFormModal);
    document.getElementById('modalCancelBtn').addEventListener('click', closeFormModal);
    document.getElementById('modalSaveBtn').addEventListener('click', saveForm);

    document.getElementById('deleteModalCloseBtn').addEventListener('click', closeDeleteModal);
    document.getElementById('deleteCancelBtn').addEventListener('click', closeDeleteModal);
    document.getElementById('deleteConfirmBtn').addEventListener('click', confirmDelete);

    // ปิด modal เมื่อคลิกพื้นหลังสีเทา (นอกกล่อง)
    document.getElementById('formModalOverlay').addEventListener('click', (e) => {
        if (e.target.id === 'formModalOverlay') closeFormModal();
    });
    document.getElementById('deleteModalOverlay').addEventListener('click', (e) => {
        if (e.target.id === 'deleteModalOverlay') closeDeleteModal();
    });

    const logoutBtn = document.getElementById('logoutBtn');
    if (logoutBtn) {
        logoutBtn.addEventListener('click', async () => {
            try {
                const res = await fetch('/api/logout', { method: 'POST' });
                const data = await res.json();
                window.location.href = data.redirect || '/login';
            } catch (err) {
                window.location.href = '/login';
            }
        });
    }
});