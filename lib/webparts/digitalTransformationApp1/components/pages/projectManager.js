// ✅ Universal dynamic site URL detection
import Swal from 'sweetalert2';
const managerContent = `
<div class="manager-page fade-in">
  <div class="page-header">
    <h1>👤 Project Managers</h1>
  </div>
  <div style="text-align:right; margin-bottom:15px;">
    <button class="btn-primary show-form-btn">➕ Add Manager</button>
  </div>

  <!-- Table Section -->
  <section class="table-section card">
    <div class="table-actions">
      <input type="text" class="search-input" placeholder="Search Managers" />
      <button class="btn-primary export-btn">Export CSV</button>
    </div>
    <div style="overflow-x:auto;">
      <table class="modern-table">
        <thead>
          <tr>
            <th>Name</th>
            <th>Department</th>
            <th>Email</th>
            <th>Phone</th>
            <th>Role</th>
            <th>Actions</th>
          </tr>
        </thead>
        <tbody class="manager-table-body"></tbody>
      </table>
    </div>
    <div class="pagination"></div>
  </section>

  <!-- Form Section -->
  <section class="form-section card" style="display:none;">
    <div class="form-header"><h2>➕ Add Manager</h2></div>
    <form id="managerForm">
      <div class="form-row">
        <div class="form-group"><label>Name</label><input type="text" id="Title" required /></div>
        <div class="form-group"><label>Department</label><input type="text" id="Department" required /></div>
      </div>
      <div class="form-row">
        <div class="form-group"><label>Email</label><input type="email" id="Email" required /></div>
        <div class="form-group"><label>Phone</label><input type="text" id="Phone" required /></div>
      </div>
      <div class="form-row">
        <div class="form-group"><label>Role</label><input type="text" id="Role" required /></div>
      </div>
      <div class="form-actions">
        <button type="submit" class="btn-primary">Save Manager</button>
        <button type="button" class="btn-secondary back-to-table">Back</button>
      </div>
    </form>
  </section>
</div>
`;
export default {
    content: managerContent,
    style: `
.manager-page { font-family:'Segoe UI',Arial,sans-serif; color:#333; animation:fadeIn 0.4s ease-in; }
.page-header h1 { font-size:22px; font-weight:600; color:#0078d4; }
.card { background:#fff; border-radius:10px; box-shadow:0 2px 8px rgba(0,0,0,0.08); padding:15px; margin-bottom:20px; }
.table-actions { display:flex; justify-content:space-between; margin-bottom:10px; flex-wrap:wrap; gap:10px; }
.modern-table { width:100%; border-collapse:collapse; font-size:14px; }
.modern-table th { background:#f3f2f1; text-align:left; padding:12px; font-weight:600; border-bottom:2px solid #e1e1e1; }
.modern-table td { padding:10px; border-bottom:1px solid #eee; }
.modern-table tr:hover { background:#fafafa; }
.form-row { display:flex; flex-wrap:wrap; gap:20px; margin-bottom:15px; }
.form-group { flex:1 1 calc(50% - 10px); display:flex; flex-direction:column; }
.form-group label { font-weight:500; margin-bottom:6px; }
.form-group input { padding:10px 12px; border-radius:6px; border:1px solid #ccc; }
.form-group input:focus { border-color:#0078d4; outline:none; box-shadow:0 0 3px rgba(0,120,212,0.4); }
.form-actions { text-align:center; margin-top:25px; }
.btn-primary, .btn-secondary { padding:10px 20px; border-radius:6px; border:none; cursor:pointer; font-weight:500; font-size:14px; transition:all 0.3s; }
.btn-primary { background:#0078d4; color:#fff; }
.btn-primary:hover { background:#005a9e; }
.btn-secondary { background:#e1e1e1; color:#333; margin-left:10px; }
.btn-secondary:hover { background:#ccc; }
.pagination { margin-top:15px; display:flex; gap:5px; justify-content:flex-end; flex-wrap:wrap; }
.pagination button { padding:6px 12px; border:none; background:#e1e1e1; border-radius:4px; cursor:pointer; }
.pagination button.active { background:#0078d4; color:#fff; }
.toast { position:fixed; bottom:20px; right:20px; padding:12px 20px; background:#0078d4; color:#fff; border-radius:6px; box-shadow:0 2px 6px rgba(0,0,0,0.2); opacity:0; transition:opacity 0.3s; z-index:1000; }
.actions-cell { display:flex; gap:5px; }
.edit-btn, .delete-btn { width:28px; height:28px; border:none; border-radius:4px; cursor:pointer; display:flex; align-items:center; justify-content:center; }
.edit-btn { background:#ff9800; color:#fff; }
.delete-btn { background:#dc3545; color:#fff; }
.form-group input,
.form-group select {
  padding: 10px 12px;
  border-radius: 6px;
  border: 1px solid #ccc;
  font-size: 14px;
  background-color: #fff;
  transition: all 0.3s ease;
  appearance: none; /* hides browser default arrow for consistent styling */
  -webkit-appearance: none;
  -moz-appearance: none;
  background-image: linear-gradient(45deg, transparent 50%, #0078d4 50%),
    linear-gradient(135deg, #0078d4 50%, transparent 50%);
  background-position: calc(100% - 20px) calc(1em + 2px),
    calc(100% - 15px) calc(1em + 2px);
  background-size: 5px 5px, 5px 5px;
  background-repeat: no-repeat;
}

.form-group select:focus,
.form-group input:focus {
  border-color: #0078d4;
  outline: none;
  box-shadow: 0 0 3px rgba(0, 120, 212, 0.4);
}

.form-group select {
  cursor: pointer;
}

@keyframes fadeIn { from{opacity:0;} to{opacity:1;} }
@media (max-width:768px){ .form-group{flex:1 1 100%;} }
  `,
    script: () => {
        setTimeout(() => {
            const listName = "Project Managers List";
            // ✅ Universal site URL
            let siteUrl = "";
            if (typeof _spPageContextInfo !== "undefined" && (_spPageContextInfo === null || _spPageContextInfo === void 0 ? void 0 : _spPageContextInfo.webAbsoluteUrl)) {
                siteUrl = _spPageContextInfo.webAbsoluteUrl;
            }
            else {
                siteUrl = window.location.origin + "/sites/Kriya";
            }
            const apiBase = `${siteUrl}/_api/web/lists/getbytitle('${listName}')/items`;
            const form = document.querySelector("#managerForm");
            const formSection = document.querySelector(".form-section");
            const tableSection = document.querySelector(".table-section");
            const tableBody = document.querySelector(".manager-table-body");
            const showFormBtn = document.querySelector(".show-form-btn");
            const backBtn = document.querySelector(".back-to-table");
            const submitBtn = form.querySelector('button[type="submit"]');
            const searchInput = document.querySelector(".search-input");
            const exportBtn = document.querySelector(".export-btn");
            let managerData = [];
            let filteredData = [];
            let editingId = null;
            const pageSize = 5;
            let currentPage = 1;
            // ✅ Toast utility
            const showToast = (msg) => {
                const toast = document.createElement("div");
                toast.className = "toast";
                toast.textContent = msg;
                document.body.appendChild(toast);
                setTimeout(() => (toast.style.opacity = "1"), 100);
                setTimeout(() => {
                    toast.style.opacity = "0";
                    toast.remove();
                }, 2500);
            };
            // ✅ Get form digest dynamically
            const getFormDigest = async () => {
                if (typeof _spPageContextInfo !== "undefined" && (_spPageContextInfo === null || _spPageContextInfo === void 0 ? void 0 : _spPageContextInfo.formDigestValue)) {
                    return _spPageContextInfo.formDigestValue;
                }
                const res = await fetch(`${siteUrl}/_api/contextinfo`, {
                    method: "POST",
                    headers: { Accept: "application/json;odata=verbose" },
                });
                const data = await res.json();
                return data.d.GetContextWebInformation.FormDigestValue;
            };
            const getListItemType = async () => {
                const res = await fetch(`${siteUrl}/_api/web/lists/getbytitle('${listName}')?$select=ListItemEntityTypeFullName`, { headers: { Accept: "application/json;odata=verbose" } });
                const data = await res.json();
                return data.d.ListItemEntityTypeFullName;
            };
            // ✅ Fetch managers
            const fetchManagers = async () => {
                try {
                    const res = await fetch(apiBase, { headers: { Accept: "application/json;odata=verbose" } });
                    const data = await res.json();
                    managerData = data.d.results;
                    filteredData = managerData;
                    currentPage = 1;
                    renderTable();
                }
                catch (err) {
                    console.error(err);
                    showToast("⚠️ Error loading managers");
                }
            };
            // ✅ Add manager
            const addManager = async (manager) => {
                // 🔍 Prevent duplicate (Name + Email + Role)
                const duplicate = managerData.some((m) => m.Title.toLowerCase() === manager.Title.toLowerCase() ||
                    m.Email.toLowerCase() === manager.Email.toLowerCase() ||
                    (m.Title.toLowerCase() === manager.Title.toLowerCase() &&
                        m.Role.toLowerCase() === manager.Role.toLowerCase()));
                if (duplicate) {
                    showToast("⚠️ Manager already exists!");
                    return;
                }
                try {
                    const digest = await getFormDigest();
                    const itemType = await getListItemType();
                    const payload = {
                        __metadata: { type: itemType },
                        Title: manager.Title,
                        Department: manager.Department,
                        Email: manager.Email,
                        Phone: manager.Phone,
                        Role: manager.Role,
                    };
                    const res = await fetch(apiBase, {
                        method: "POST",
                        headers: {
                            Accept: "application/json;odata=verbose",
                            "Content-Type": "application/json;odata=verbose",
                            "X-RequestDigest": digest,
                        },
                        body: JSON.stringify(payload),
                    });
                    if (res.ok) {
                        showToast("✅ Manager added successfully!");
                        fetchManagers();
                    }
                    else
                        showToast("❌ Add failed");
                }
                catch (err) {
                    console.error(err);
                    showToast("❌ Add failed");
                }
            };
            // ✅ Update manager
            const updateManager = async (id, manager) => {
                // 🔍 Prevent duplicate for update
                const duplicate = managerData.some((m) => m.Id !== id &&
                    (m.Title.toLowerCase() === manager.Title.toLowerCase() ||
                        m.Email.toLowerCase() === manager.Email.toLowerCase() ||
                        (m.Title.toLowerCase() === manager.Title.toLowerCase() &&
                            m.Role.toLowerCase() === manager.Role.toLowerCase())));
                if (duplicate) {
                    showToast("⚠️ Duplicate record exists!");
                    return;
                }
                try {
                    const digest = await getFormDigest();
                    const itemType = await getListItemType();
                    const payload = {
                        __metadata: { type: itemType },
                        Title: manager.Title,
                        Department: manager.Department,
                        Email: manager.Email,
                        Phone: manager.Phone,
                        Role: manager.Role,
                    };
                    const res = await fetch(`${apiBase}(${id})`, {
                        method: "POST",
                        headers: {
                            Accept: "application/json;odata=verbose",
                            "Content-Type": "application/json;odata=verbose",
                            "X-RequestDigest": digest,
                            "X-HTTP-Method": "MERGE",
                            "IF-MATCH": "*",
                        },
                        body: JSON.stringify(payload),
                    });
                    if (res.ok) {
                        showToast("✅ Manager updated!");
                        fetchManagers();
                    }
                    else
                        showToast("❌ Update failed");
                }
                catch (err) {
                    console.error(err);
                    showToast("❌ Update failed");
                }
            };
            // ✅ Delete manager
            const deleteManager = async (id) => {
                Swal.fire({
                    title: "Are you sure?",
                    text: "You won't be able to revert this!",
                    icon: "warning",
                    showCancelButton: true,
                    confirmButtonColor: "#0078d4",
                    cancelButtonColor: "#d33",
                    confirmButtonText: "Yes, delete it!",
                }).then(async (result) => {
                    if (result.isConfirmed) {
                        try {
                            const digest = await getFormDigest();
                            const res = await fetch(`${apiBase}(${id})`, {
                                method: "POST",
                                headers: {
                                    Accept: "application/json;odata=verbose",
                                    "X-RequestDigest": digest,
                                    "X-HTTP-Method": "DELETE",
                                    "IF-MATCH": "*",
                                },
                            });
                            if (res.ok) {
                                showToast("🗑️ Manager deleted!");
                                fetchManagers();
                            }
                            else
                                showToast("❌ Delete failed");
                        }
                        catch (err) {
                            console.error(err);
                            showToast("❌ Delete request failed");
                        }
                    }
                });
            };
            // ✅ Render table
            const renderTable = () => {
                const start = (currentPage - 1) * pageSize;
                const paginatedData = filteredData.slice(start, start + pageSize);
                tableBody.innerHTML = paginatedData
                    .map((m) => `
      <tr>
        <td>${m.Title || ""}</td>
        <td>${m.Department || ""}</td>
        <td>${m.Email || ""}</td>
        <td>${m.Phone || ""}</td>
        <td>${m.Role || ""}</td>
        <td class="actions-cell">
          <button class="edit-btn" data-id="${m.Id}">✏️</button>
          <button class="delete-btn" data-id="${m.Id}">🗑️</button>
        </td>
      </tr>`)
                    .join("");
                document.querySelectorAll(".edit-btn").forEach((btn) => btn.addEventListener("click", () => {
                    const id = parseInt(btn.dataset.id);
                    const item = managerData.find((x) => x.Id === id);
                    if (!item)
                        return;
                    document.querySelector("#Title").value = item.Title;
                    document.querySelector("#Department").value = item.Department;
                    document.querySelector("#Email").value = item.Email;
                    document.querySelector("#Phone").value = item.Phone;
                    document.querySelector("#Role").value = item.Role;
                    editingId = id;
                    formSection.style.display = "block";
                    tableSection.style.display = "none";
                    showFormBtn.style.display = "none";
                    submitBtn.textContent = "Update Manager";
                }));
                document.querySelectorAll(".delete-btn").forEach((btn) => btn.addEventListener("click", () => {
                    const id = parseInt(btn.dataset.id);
                    deleteManager(id);
                }));
                renderPagination();
            };
            const renderPagination = () => {
                const totalPages = Math.ceil(filteredData.length / pageSize);
                const pagination = document.querySelector(".pagination");
                pagination.innerHTML = "";
                for (let i = 1; i <= totalPages; i++) {
                    const btn = document.createElement("button");
                    btn.textContent = i.toString();
                    if (i === currentPage)
                        btn.classList.add("active");
                    btn.addEventListener("click", () => {
                        currentPage = i;
                        renderTable();
                    });
                    pagination.appendChild(btn);
                }
            };
            // ✅ Search
            searchInput.addEventListener("input", () => {
                const val = searchInput.value.toLowerCase();
                filteredData = managerData.filter((m) => {
                    var _a, _b, _c, _d, _e;
                    return ((_a = m.Title) === null || _a === void 0 ? void 0 : _a.toLowerCase().includes(val)) ||
                        ((_b = m.Department) === null || _b === void 0 ? void 0 : _b.toLowerCase().includes(val)) ||
                        ((_c = m.Email) === null || _c === void 0 ? void 0 : _c.toLowerCase().includes(val)) ||
                        ((_d = m.Phone) === null || _d === void 0 ? void 0 : _d.toLowerCase().includes(val)) ||
                        ((_e = m.Role) === null || _e === void 0 ? void 0 : _e.toLowerCase().includes(val));
                });
                currentPage = 1;
                renderTable();
            });
            // ✅ Export CSV
            exportBtn.addEventListener("click", () => {
                let csv = "Title,Department,Email,Phone,Role\n";
                managerData.forEach((m) => (csv += `"${m.Title}","${m.Department}","${m.Email}","${m.Phone}","${m.Role}"\n`));
                const blob = new Blob([csv], { type: "text/csv" });
                const url = URL.createObjectURL(blob);
                const a = document.createElement("a");
                a.href = url;
                a.download = "managers.csv";
                a.click();
                URL.revokeObjectURL(url);
            });
            // ✅ Form submit
            form.addEventListener("submit", (e) => {
                e.preventDefault();
                const manager = {
                    Title: document.querySelector("#Title").value.trim(),
                    Department: document.querySelector("#Department").value.trim(),
                    Email: document.querySelector("#Email").value.trim(),
                    Phone: document.querySelector("#Phone").value.trim(),
                    Role: document.querySelector("#Role").value.trim(),
                };
                if (editingId)
                    updateManager(editingId, manager);
                else
                    addManager(manager);
                form.reset();
                formSection.style.display = "none";
                tableSection.style.display = "block";
                showFormBtn.style.display = "inline-block";
                editingId = null;
                submitBtn.textContent = "Save Manager";
            });
            showFormBtn.addEventListener("click", () => {
                form.reset();
                editingId = null;
                formSection.style.display = "block";
                tableSection.style.display = "none";
                showFormBtn.style.display = "none";
            });
            backBtn.addEventListener("click", () => {
                formSection.style.display = "none";
                tableSection.style.display = "block";
                showFormBtn.style.display = "inline-block";
            });
            // ✅ Inject dropdown options
            const roleField = document.querySelector("#Role");
            if (roleField && !(roleField instanceof HTMLSelectElement)) {
                const select = document.createElement("select");
                select.id = "Role";
                select.required = true;
                select.innerHTML = `
        <option value="">-- Select Role --</option>
        <option>Super User</option>
        <option>Group Head</option>
        <option>Engineering Manager</option>
        <option>Project Manager</option>
        <option>Engineering</option>
      `;
                roleField.parentElement.replaceChild(select, roleField);
            }
            fetchManagers();
        }, 300);
    },
};
//# sourceMappingURL=projectManager.js.map