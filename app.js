const LOGIN_USER="admin",LOGIN_PASS_DEFAULT="tools123",RESET_PIN="4256";function savedPass(){return localStorage.getItem("shipment_tools_password")||LOGIN_PASS_DEFAULT}function savePass(v){localStorage.setItem("shipment_tools_password",v)}function doLogin(){const u=document.getElementById("username")?.value?.trim(),p=document.getElementById("password")?.value||"",r=document.getElementById("remember");u===LOGIN_USER&&p===savedPass()?(r?.checked?localStorage.setItem("shipment_tools_saved_user",LOGIN_USER):localStorage.removeItem("shipment_tools_saved_user"),localStorage.setItem("shipment_tools_logged_in","yes"),window.location.href="dashboard.html"):alert("Wrong Username or Password")}function bootstrapLogin(){const s=localStorage.getItem("shipment_tools_saved_user");if(s&&document.getElementById("username")){document.getElementById("username").value=s;const r=document.getElementById("remember");r&&(r.checked=!0)}}function logoutNow(){return localStorage.removeItem("shipment_tools_logged_in"),window.location.href="index.html",!1}function protectPage(){const i=location.pathname.endsWith("index.html")||"/"===location.pathname||location.pathname.endsWith("/");i||"yes"===localStorage.getItem("shipment_tools_logged_in")||(window.location.href="index.html")}function openForgot(){document.getElementById("forgotModal")?.classList.add("open")}function closeForgot(){document.getElementById("forgotModal")?.classList.remove("open"),document.getElementById("resetStep2")?.classList.add("hidden");const p=document.getElementById("pinInput");p&&(p.value="");const n=document.getElementById("newPasswordInput");n&&(n.value="")}function verifyPin(){(document.getElementById("pinInput")?.value||"").trim()===RESET_PIN?document.getElementById("resetStep2")?.classList.remove("hidden"):alert("Wrong PIN")}function saveNewPassword(){const n=document.getElementById("newPasswordInput")?.value||"";n?(savePass(n),alert("Password changed successfully"),closeForgot()):alert("Enter new password")}const DEFAULT_DB=[{company:"diamond traders",ntn:"4967890"},{company:"vision exporters",ntn:"3746594"},{company:"pearl embroidery",ntn:"7812459"},{company:"classic sports",ntn:"4567812"}];function getDB(){const r=localStorage.getItem("shipment_tools_ntn_db");if(r)try{return JSON.parse(r)}catch(e){}return localStorage.setItem("shipment_tools_ntn_db",JSON.stringify(DEFAULT_DB)),DEFAULT_DB.slice()}function saveDB(d){localStorage.setItem("shipment_tools_ntn_db",JSON.stringify(d))}function titleCase(s){return String(s||"").replace(/\b\w/g,c=>c.toUpperCase())}function normalize(s){return String(s||"").toLowerCase().trim().replace(/\s+/g," ")}function digitsOnly(s){return String(s??"").replace(/\D/g,"")}function firstExisting(r,n){for(const x of n)if(void 0!==r[x]&&null!==r[x]&&""!==String(r[x]).trim())return r[x];return""}function numVal(v){const n=parseFloat(String(v??"").replace(/,/g,"").trim());return isNaN(n)?0:n}function setText(id,v){const e=document.getElementById(id);e&&(e.textContent=v)}function escapeHtml(v){return String(v??"").replace(/[&<>"]/g,m=>({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;"}[m]))}function renderTable(id,h){const e=document.getElementById(id);e&&(e.innerHTML=h)}function setFileName(i,l){const input=document.getElementById(i),label=document.getElementById(l);input&&label&&input.addEventListener("change",()=>{label.textContent=input.files[0]?input.files[0].name:"No file selected"})}function downloadRows(rows,filename,sheetName){if("undefined"==typeof XLSX)return alert("Excel library not loaded");const ws=XLSX.utils.json_to_sheet(rows),wb=XLSX.utils.book_new();XLSX.utils.book_append_sheet(wb,ws,sheetName),XLSX.writeFile(wb,filename)}function parseExcel(file,cb){const r=new FileReader;r.onload=e=>{const data=new Uint8Array(e.target.result),wb=XLSX.read(data,{type:"array"}),ws=wb.Sheets[wb.SheetNames[0]];cb(XLSX.utils.sheet_to_json(ws,{defval:""}))},r.readAsArrayBuffer(file)}let hsRows=[];function initHS(){setFileName("hsFile","hsFileName");const p=document.getElementById("hsProcessBtn");p&&(p.onclick=()=>{const file=document.getElementById("hsFile").files[0],country=(document.getElementById("hsCountry").value||"").trim().toUpperCase();file?country?parseExcel(file,rows=>{hsRows=rows.filter(r=>String(firstExisting(r,["Recip Cntry","Country","Country Code"])).trim().toUpperCase()===country).filter(r=>""!==String(firstExisting(r,["CE Commodity Description","Commodity Description","Description"])).trim()).map(r=>{const hs=digitsOnly(firstExisting(r,["Commodity Harmonized Code","HS Code","Harmonized Code"]));return{...r,"Commodity Harmonized Code":hs,__bad:hs.length<10,"HS Code Status":hs.length<10?`HS Code ${hs.length} digits`:"Valid"}}).sort((a,b)=>Number(b.__bad)-Number(a.__bad)),renderTable("hsBody",hsRows.slice(0,20).map(r=>`<tr class="${r.__bad?"row-alert":""}"><td>${escapeHtml(firstExisting(r,["Tracking Number","Shipment Number","Invoice No"]))}</td><td>${escapeHtml(firstExisting(r,["Recip Cntry","Country","Country Code"]))}</td><td>${escapeHtml(r["Commodity Harmonized Code"])}</td><td class="${r.__bad?"status-warn":"status-valid"}">${escapeHtml(r["HS Code Status"])}</td></tr>`).join("")||'<tr><td colspan="4">No rows matched</td></tr>'),setText("hsTotal",hsRows.length),setText("hsInvalid",hsRows.filter(x=>x.__bad).length)}):alert("Enter country code"):alert("Upload Excel file first")},document.getElementById("hsExportBtn").onclick=()=>{hsRows.length?downloadRows(hsRows.map(({__bad,...x})=>x),"HS_Code_Result.xlsx","HS Code Verification"):alert("No data to export")})}let missingRows = [];
let bucketRows = [];

function initMissing() {
  setFileName("missingFile", "missingFileName");
  const p = document.getElementById("missingProcessBtn");
  
  p && (p.onclick = () => {
    const file = document.getElementById("missingFile").files[0];
    
    if (file) {
      parseExcel(file, rows => {
        missingRows = rows.filter(r => {
          const ntn = normalize(firstExisting(r, ["NTN", "NTN Number", "Company NTN"]));
          const company = String(firstExisting(r, ["Shipper Company", "Shipper Name", "Company"])).trim();
          const desc = String(firstExisting(r, ["CE Commodity Description", "Commodity Description", "Description"])).trim();
          const value = numVal(firstExisting(r, ["Value", "Declared Value", "Customs Value"]));
          const upper = company.toUpperCase().trim();

          // Filter logic
          return !ntn &&                          // No NTN
            !/\d/.test(company) &&                // Company name doesn't contain digits (for NTN validation)
            !!desc &&                             // Description must not be blank
            !upper.endsWith("-A") &&              // Remove rows where company name ends with "-A"
            !upper.endsWith(" -A") &&             // Remove rows where company name ends with " -A"
            !upper.endsWith("E FORM") &&
            !upper.endsWith("- E FORM") &&
            !upper.endsWith("-EFORM") &&
            !upper.endsWith("E FORM B") &&
            !upper.endsWith("-E FORM B") &&
            value < 500;                          // Remove rows with value >= 500
        }).map(r => ({ ...r, NTN: "MISSING" }));

         // Render table with the filtered rows
        renderTable("missingBody", missingRows.slice(0, 20).map(r => `
          <tr>
            <td>${escapeHtml(firstExisting(r, ["Tracking Number","Shipment Number","Invoice No"]))}</td>
            <td>${escapeHtml(firstExisting(r, ["Shipper Company","Shipper Name","Company"]))}</td>
            <td class="status-missing">MISSING</td>
            <td>${escapeHtml(firstExisting(r, ["CE Commodity Description", "Commodity Description", "Description"]))}</td>
          </tr>
        `).join("") || '<tr><td colspan="4">No rows matched</td></tr>');

        setText("missingTotal", missingRows.length);
      });
    } else {
      alert("Upload Excel file first");
    }
  });

  // Export filtered rows
const missingExportBtn = document.getElementById("missingExportBtn");

if (missingExportBtn) {
  missingExportBtn.onclick = () => {
    missingRows.length
      ? downloadRows(missingRows, "NTN_Missing_Result.xlsx", "NTN Missing")
      : alert("No result to export");
  };
}
}

function initBucket(){

setFileName("bucketFile","bucketFileName");

const p=document.getElementById("bucketProcessBtn");

p&&(p.onclick=()=>{

const file=document.getElementById("bucketFile").files[0];

file?parseExcel(file,rows=>{

bucketRows = rows.filter(r => {

  // Description (Commodity Description)
  const desc = String(firstExisting(r, ["CE Commodity Description", "Commodity Description", "Description"])).trim();

  // Shipper Company (and modify order here if needed)
  const company = String(firstExisting(r, ["Shipper Company", "Shipper Name", "Company"])).trim();

  // City (Modify City data filter)
  const city = String(firstExisting(r, ["Shpr City", "City", "Shipper City"])).toUpperCase();

  // Reference Number
  const ref = String(firstExisting(r, ["Shipper Ref", "Reference Number", "Reference"])).trim();

  // Value for filtering
  const value = numVal(firstExisting(r, ["Customs Value", "Declared Value", "Value"]));

  // Modify the order of company and tracking number here if required
  const trackingNumber = String(firstExisting(r, ["Tracking Number", "Shipment Number", "Invoice No"])).trim();
  
  const upperCompany = company.toUpperCase();

  // description blank
if (!desc) return false;

// company contains digits (most NTN patterns)
if (/\d{4,}/.test(company)) return false;

// NTN text detection
if (/NTN/i.test(company)) return false;

// NTN inside brackets
if (/\([A-Z]?\d{5,}\)/i.test(company)) return false;

// NTN with hyphen patterns
if (/\d{5,}-\d+/i.test(company)) return false;

// EFORM patterns
if (/-E\s*FORM/i.test(company) || /-EFORM/i.test(company)) return false;

// value check
if (value >= 500) return false;

// Sialkot city filter
if (city.includes("SIALKOT") || city.includes("SKT")) { return true;
  
// lahore ko force reject
if (city.includes("LAHORE")) return false;

  return true;
});
/* render table */

renderTable(
  "bucketBody",
  bucketRows.slice(0, 20).map(r => `
    <tr>
      <td>${escapeHtml(firstExisting(r, ["Tracking Number", "Shipment Number", "Invoice No"]))}</td> <!-- Tracking Number -->
      <td>${escapeHtml(firstExisting(r, ["Shipper Company", "Shipper Name", "Company"]))}</td> <!-- Shipper Company -->
      <td>${escapeHtml(firstExisting(r, ["CE Commodity Description", "Commodity Description", "Description"]))}</td> <!-- Description -->
      <td>${escapeHtml(firstExisting(r, ["City", "Shipper City", "Address", "Shpr Addl Addr"]))}</td> <!-- City -->
    </tr>
  `).join("") || '<tr><td colspan="4">No rows matched</td></tr>'
);

setText("bucketTotal",bucketRows.length);

}):alert("Upload Excel file first");

});

const bucketExportBtn = document.getElementById("bucketExportBtn");

if(bucketExportBtn){

  bucketExportBtn.onclick = ()=>{

    bucketRows.length
    ? downloadRows(bucketRows,"Bucket_Shop_Result.xlsx","Bucket Shop")
    : alert("No result to export");

  };

}

}


let dupRows=[];function initDuplicate(){setFileName("dupFile","dupFileName");const p=document.getElementById("dupProcessBtn");p&&(p.onclick=()=>{const file=document.getElementById("dupFile").files[0];file?parseExcel(file,rows=>{const count={};rows.forEach(r=>{const t=normalize(firstExisting(r,["Tracking Number","Shipment Number","Invoice No"]));t&&(count[t]=(count[t]||0)+1)}),dupRows=rows.map(r=>{const t=normalize(firstExisting(r,["Tracking Number","Shipment Number","Invoice No"]));return{...r,__dup:!!(t&&count[t]>1)}}).sort((a,b)=>Number(b.__dup)-Number(a.__dup)),renderTable("dupBody",dupRows.slice(0,20).map(r=>`<tr class="${r.__dup?"row-alert":""}"><td>${escapeHtml(firstExisting(r,["Tracking Number","Shipment Number","Invoice No"]))}</td><td>${escapeHtml(firstExisting(r,["Shipper Company","Shipper Name","Company"]))}</td><td class="${r.__dup?"status-missing":"status-valid"}">${r.__dup?"DUPLICATE":"UNIQUE"}</td></tr>`).join("")||'<tr><td colspan="3">No rows found</td></tr>'),setText("dupTotal",dupRows.length),setText("dupCount",dupRows.filter(x=>x.__dup).length)}):alert("Upload Excel file first")},document.getElementById("dupExportBtn").onclick=()=>{dupRows.length?downloadRows(dupRows.filter(x=>x.__dup).map(({__dup,...r})=>r),"Duplicate_Shipments.xlsx","Duplicate"):alert("No result to export")})}function initSearch(){const btn=document.getElementById("searchBtn");btn&&(btn.onclick=()=>{const q=normalize(document.getElementById("companySearch").value),found=getDB().filter(x=>x.company.includes(q));setText("ntnSearchInfo",`Showing ${found.length} result${1===found.length?"":"s"} for "${q}":`),renderTable("ntnSearchBody",found.length?found.map(r=>`<tr><td>${escapeHtml(titleCase(r.company))}</td><td>${escapeHtml(r.ntn)}</td></tr>`).join(""):'<tr><td colspan="2">No result found</td></tr>')});const addBtn=document.getElementById("addNtnBtn");addBtn&&(addBtn.onclick=()=>{const c=normalize(document.getElementById("newCompany").value),n=String(document.getElementById("newNtn").value||"").trim();if(!c||!n)return alert("Enter company and NTN");const db=getDB();db.push({company:c,ntn:n}),saveDB(db),alert("Company and NTN saved"),document.getElementById("newCompany").value="",document.getElementById("newNtn").value=""})}let autoRows=[];function initAutoUpdate(){setFileName("autoFile","autoFileName");const p=document.getElementById("autoProcessBtn");p&&(p.onclick=()=>{const file=document.getElementById("autoFile").files[0];file?parseExcel(file,rows=>{const db=getDB();autoRows=rows.map(r=>{const company=normalize(firstExisting(r,["Shipper Company","Shipper Name","Company"])),found=db.find(x=>x.company===company);return{...r,NTN:found?found.ntn:"",__missing:!found}}).sort((a,b)=>Number(b.__missing)-Number(a.__missing)),renderTable("autoBody",autoRows.slice(0,20).map(r=>`<tr class="${r.__missing?"row-alert":""}"><td>${escapeHtml(firstExisting(r,["Shipper Company","Shipper Name","Company"]))}</td><td>${escapeHtml(r.NTN)}</td><td class="${r.__missing?"status-missing":"status-valid"}">${r.__missing?"MISSING":"FILLED"}</td></tr>`).join("")||'<tr><td colspan="3">No rows found</td></tr>'),setText("autoTotal",autoRows.length),setText("autoMissing",autoRows.filter(x=>x.__missing).length)}):alert("Upload Excel file first")},document.getElementById("autoExportBtn").onclick=()=>{autoRows.length?downloadRows(autoRows.map(({__missing,...r})=>r),"Auto_NTN_Updated.xlsx","Auto NTN Update"):alert("No result to export")})}document.addEventListener("DOMContentLoaded",()=>{protectPage(),bootstrapLogin(),initHS(),initMissing(),initBucket(),initDuplicate(),initSearch(),initAutoUpdate()});


const SEARCH_DB_KEY = "shipment_tools_ntn_db";
const SEARCH_DB_VERSION_KEY = "shipment_tools_ntn_db_version";
const SEARCH_DB_VERSION = "client_detail2_seed_v1";
let companyEditIndex = -1;

function getSeedData(){
  return Array.isArray(window.SEARCH_SEED_DATA) ? window.SEARCH_SEED_DATA : [];
}
function normalizeSearchRecord(item){
  return {
    reff: String(item?.reff || item?.REFF || item?.Ref || item?.ref || "").trim(),
    company: normalize(item?.company || item?.["COMPANY NAMES"] || item?.["Company Name"] || item?.Company || item?.Name || ""),
    cnic: String(item?.cnic || item?.CNIC || item?.["CNIC "] || "").trim(),
    ntn: String(item?.ntn || item?.NTN || item?.["NTN Number"] || "").trim(),
    expired: !!item?.expired
  };
}
function getSearchDB(){
  try {
    const raw = localStorage.getItem(SEARCH_DB_KEY);
    if (!raw) return [];
    const parsed = JSON.parse(raw);
    return Array.isArray(parsed) ? parsed.map(normalizeSearchRecord).filter(r => r.company || r.reff || r.cnic || r.ntn) : [];
  } catch (e) {
    return [];
  }
}
function saveSearchDB(rows){
  localStorage.setItem(SEARCH_DB_KEY, JSON.stringify(rows.map(normalizeSearchRecord)));
}
function seedSearchDBIfNeeded(){
  const existing = getSearchDB();
  const version = localStorage.getItem(SEARCH_DB_VERSION_KEY);
  const defaultOnly = existing.length && existing.length <= 4 && existing.every(r => ["diamond traders","vision exporters","pearl embroidery","classic sports"].includes(r.company));
  if (!existing.length || defaultOnly || version !== SEARCH_DB_VERSION) {
    const seed = getSeedData().map(normalizeSearchRecord).filter(r => r.company || r.reff || r.cnic || r.ntn);
    if (seed.length) {
      saveSearchDB(seed);
      localStorage.setItem(SEARCH_DB_VERSION_KEY, SEARCH_DB_VERSION);
    }
  }
}
function getDB(){
  const seeded = getSearchDB();
  return seeded.length ? seeded : (typeof DEFAULT_DB !== 'undefined' ? DEFAULT_DB.slice() : []);
}
function saveDB(d){
  saveSearchDB(d);
}
function titleCompany(v){
  return titleCase(String(v || '').replace(/\s+/g,' ').trim());
}
function renderSearchStats(){
  const db = getDB();
  setText('ntnTotalRecords', db.length);
  setText('ntnExpiredCount', db.filter(x => x.expired).length);
}
function matchSearch(item, query){
  const q = normalize(query);
  if (!q) return 0;
  const reff = normalize(item.reff);
  const company = normalize(item.company);
  const cnic = normalize(item.cnic);
  const ntn = normalize(item.ntn);
  if (reff === q || company === q || cnic === q || ntn === q) return 1000;
  if (reff.startsWith(q) || company.startsWith(q) || cnic.startsWith(q) || ntn.startsWith(q)) return 700;
  if (reff.includes(q) || company.includes(q) || cnic.includes(q) || ntn.includes(q)) return 400;
  return 0;
}
function actionButtons(index, item){
  const ntnBtn = item.ntn ? `<button type="button" class="mini-btn" onclick="copySearchValue(${index},'ntn')">Copy NTN</button>` : '';
  const cnicBtn = item.cnic ? `<button type="button" class="mini-btn" onclick="copySearchValue(${index},'cnic')">Copy CNIC</button>` : '';
  const reffBtn = item.reff ? `<button type="button" class="mini-btn" onclick="copySearchValue(${index},'reff')">Copy REFF</button>` : '';
  const expBtn = `<button type="button" class="mini-btn ${item.expired ? 'warn' : ''}" onclick="toggleExpireCompany(${index})">${item.expired ? 'Unexpire' : 'Expire NTN'}</button>`;
  const editBtn = `<button type="button" class="mini-btn primary" onclick="editCompany(${index})">Edit</button>`;
  return [reffBtn, cnicBtn, ntnBtn, expBtn, editBtn].filter(Boolean).join(' ');
}
function renderSearchResults(list, query){
  const body = document.getElementById('ntnSearchBody');
  const info = document.getElementById('ntnSearchInfo');
  const empty = document.getElementById('ntnEmptyState');
  if (!body || !info) return;
  if (!query) {
    info.textContent = 'Start typing to see live results.';
    body.innerHTML = '<tr><td colspan="6">No results yet</td></tr>';
    empty?.classList.add('hidden-block');
    renderSearchStats();
    return;
  }
  info.textContent = `Showing ${list.length} result${list.length === 1 ? '' : 's'} for "${query}"`;
  if (!list.length) {
    body.innerHTML = '<tr><td colspan="6">No result found</td></tr>';
    empty?.classList.remove('hidden-block');
    renderSearchStats();
    return;
  }
  empty?.classList.add('hidden-block');
  const db = getDB();
  body.innerHTML = list.map(item => {
    const realIndex = db.findIndex(x => x.reff === item.reff && x.company === item.company && x.cnic === item.cnic && x.ntn === item.ntn && !!x.expired === !!item.expired);
    return `<tr>
      <td>${escapeHtml(item.reff || '-')}</td>
      <td>${escapeHtml(titleCompany(item.company) || '-')}</td>
      <td>${escapeHtml(item.cnic || '-')}</td>
      <td>${escapeHtml(item.ntn || '-')}</td>
      <td><span class="status-pill ${item.expired ? 'expired' : 'active'}">${item.expired ? 'NTN EXPIRE' : 'ACTIVE'}</span></td>
      <td class="action-cell">${actionButtons(realIndex, item)}</td>
    </tr>`;
  }).join('');
  renderSearchStats();
}
function runLiveCompanySearch(){
  const input = document.getElementById('companySearch');
  if (!input) return;
  const q = normalize(input.value);
  const db = getDB();
  if (!q) {
    renderSearchResults([], '');
    return;
  }
  const found = db.map(item => ({ ...item, _score: matchSearch(item, q) }))
    .filter(item => item._score > 0)
    .sort((a, b) => b._score - a._score || a.company.localeCompare(b.company));
  renderSearchResults(found, input.value.trim());
}
function copySearchValue(index, field){
  const db = getDB();
  const row = db[index];
  if (!row) return;
  const value = String(row[field] || '').trim();
  if (!value) return;
  navigator.clipboard?.writeText(value).catch(() => {});
}
function renderManageCompanies(filterText = ''){
  const body = document.getElementById('companyTable');
  if (!body) return;
  const q = normalize(filterText);
  const db = getDB();
  const filtered = db.filter(item => !q || [item.reff, item.company, item.cnic, item.ntn].some(v => normalize(v).includes(q)));
  body.innerHTML = filtered.length ? filtered.map(item => {
    const realIndex = db.findIndex(x => x.reff === item.reff && x.company === item.company && x.cnic === item.cnic && x.ntn === item.ntn && !!x.expired === !!item.expired);
    return `<tr>
      <td><input type="checkbox" class="company-select" data-index="${realIndex}"></td>
      <td>${escapeHtml(item.reff || '-')}</td>
      <td>${escapeHtml(titleCompany(item.company) || '-')}</td>
      <td>${escapeHtml(item.cnic || '-')}</td>
      <td>${escapeHtml(item.ntn || '-')}</td>
      <td><span class="status-pill ${item.expired ? 'expired' : 'active'}">${item.expired ? 'NTN EXPIRE' : 'ACTIVE'}</span></td>
      <td class="action-cell">
        <button type="button" class="mini-btn primary" onclick="editCompany(${realIndex})">Edit</button>
        <button type="button" class="mini-btn ${item.expired ? 'warn' : ''}" onclick="toggleExpireCompany(${realIndex})">${item.expired ? 'Unexpire' : 'Expire NTN'}</button>
      </td>
    </tr>`;
  }).join('') : '<tr><td colspan="7">No companies found</td></tr>';
  renderSearchStats();
}
function openManage(){
  const m = document.getElementById('manageModal');
  if (m) m.style.display = 'flex';
  const s = document.getElementById('companySearchManage');
  if (s) s.value = '';
  const a = document.getElementById('selectAllCompanies');
  if (a) a.checked = false;
  renderManageCompanies('');
}
function closeManage(){
  const m = document.getElementById('manageModal');
  if (m) m.style.display = 'none';
}
function openAddNew(){
  companyEditIndex = -1;
  document.getElementById('addEditTitle').textContent = 'Add New Company';
  document.getElementById('companyReffInput').value = '';
  document.getElementById('companyNameInput').value = '';
  document.getElementById('companyCNICInput').value = '';
  document.getElementById('companyNTNInput').value = '';
  document.getElementById('companyExpiredInput').checked = false;
  document.getElementById('addEditModal').style.display = 'flex';
}
function closeAddNew(){
  const m = document.getElementById('addEditModal');
  if (m) m.style.display = 'none';
}
function editCompany(index){
  const db = getDB();
  const item = db[index];
  if (!item) return;
  companyEditIndex = index;
  document.getElementById('addEditTitle').textContent = 'Edit Company';
  document.getElementById('companyReffInput').value = item.reff || '';
  document.getElementById('companyNameInput').value = titleCompany(item.company);
  document.getElementById('companyCNICInput').value = item.cnic || '';
  document.getElementById('companyNTNInput').value = item.ntn || '';
  document.getElementById('companyExpiredInput').checked = !!item.expired;
  document.getElementById('addEditModal').style.display = 'flex';
}
function saveCompanyFromModal(){
  const db = getDB();
  const record = normalizeSearchRecord({
    reff: document.getElementById('companyReffInput').value,
    company: document.getElementById('companyNameInput').value,
    cnic: document.getElementById('companyCNICInput').value,
    ntn: document.getElementById('companyNTNInput').value,
    expired: document.getElementById('companyExpiredInput').checked
  });
  if (!record.company) return alert('Enter company name');
  const duplicateIndex = db.findIndex((x, idx) => idx !== companyEditIndex && x.company === record.company && x.reff === record.reff);
  if (duplicateIndex !== -1) return alert('This company + REFF already exists');
  if (companyEditIndex === -1) db.unshift(record);
  else db[companyEditIndex] = record;
  saveDB(db);
  closeAddNew();
  renderManageCompanies(document.getElementById('companySearchManage')?.value || '');
  runLiveCompanySearch();
}
function deleteSelectedCompanies(){
  const checks = Array.from(document.querySelectorAll('.company-select:checked'));
  if (!checks.length) return alert('No company selected');
  const db = getDB();
  const indexes = checks.map(c => Number(c.dataset.index)).sort((a, b) => b - a);
  indexes.forEach(i => { if (i >= 0) db.splice(i, 1); });
  saveDB(db);
  renderManageCompanies(document.getElementById('companySearchManage')?.value || '');
  runLiveCompanySearch();
}
function toggleExpireCompany(index){
  const db = getDB();
  const item = db[index];
  if (!item) return;
  item.expired = !item.expired;
  saveDB(db);
  renderManageCompanies(document.getElementById('companySearchManage')?.value || '');
  runLiveCompanySearch();
}
function importCompaniesFromExcel(file){
  parseExcel(file, rows => {
    const imported = rows.map(r => normalizeSearchRecord({
      reff: firstExisting(r, ['REFF', 'Ref', 'Reference', 'Reference Number']),
      company: firstExisting(r, ['COMPANY NAMES', 'Company Name', 'Company', 'Name']),
      cnic: firstExisting(r, ['CNIC ', 'CNIC', 'Cnic']),
      ntn: firstExisting(r, ['NTN', 'NTN Number', 'NTN No'])
    })).filter(r => r.company || r.reff || r.cnic || r.ntn);
    if (!imported.length) return alert('No valid rows found in Excel');
    saveDB(imported);
    localStorage.setItem(SEARCH_DB_VERSION_KEY, SEARCH_DB_VERSION + '_imported');
    renderManageCompanies('');
    runLiveCompanySearch();
  });
}
function initSearchEnhanced(){
  seedSearchDBIfNeeded();
  if (!document.getElementById('companySearch')) return;
  renderSearchStats();
  const input = document.getElementById('companySearch');
  const searchBtn = document.getElementById('searchBtn');
  const saveModalBtn = document.getElementById('saveCompanyBtn');
  const excelUpload = document.getElementById('excelUpload');
  const manageSearch = document.getElementById('companySearchManage');
  const selectAll = document.getElementById('selectAllCompanies');
  input?.addEventListener('input', runLiveCompanySearch);
  searchBtn && (searchBtn.onclick = runLiveCompanySearch);
  saveModalBtn && (saveModalBtn.onclick = saveCompanyFromModal);
  excelUpload?.addEventListener('change', () => {
    const file = excelUpload.files[0];
    if (file) importCompaniesFromExcel(file);
  });
  manageSearch?.addEventListener('input', function(){ renderManageCompanies(this.value); });
  selectAll?.addEventListener('change', function(){ document.querySelectorAll('.company-select').forEach(cb => cb.checked = this.checked); });
  renderSearchResults([], '');
}
document.addEventListener('DOMContentLoaded', initSearchEnhanced);
