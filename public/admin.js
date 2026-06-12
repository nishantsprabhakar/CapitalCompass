const adminIdentity = document.getElementById("adminIdentity");
const adminLocked = document.getElementById("adminLocked");
const adminDashboard = document.getElementById("adminDashboard");
const adminUsers = document.getElementById("adminUsers");
const promoList = document.getElementById("promoList");
const promoForm = document.getElementById("promoForm");
const settingsForm = document.getElementById("settingsForm");
const refreshAdmin = document.getElementById("refreshAdmin");
const adminLogout = document.getElementById("adminLogout");
const adminStatus = document.getElementById("adminStatus");
const kpiUsers = document.getElementById("kpiUsers");
const kpiPremium = document.getElementById("kpiPremium");
const kpiPromos = document.getElementById("kpiPromos");
const kpiSuspended = document.getElementById("kpiSuspended");
let csrfToken = "";
let currentUser = null;

initAdminEffects();
initAdmin();

refreshAdmin.addEventListener("click", () => loadAdmin());
adminLogout.addEventListener("click", async () => {
  await apiFetch("/api/logout", { method: "POST" });
  window.location.href = "/";
});

promoForm.addEventListener("submit", async (event) => {
  event.preventDefault();
  setStatus("Saving promo...");
  const body = Object.fromEntries(new FormData(promoForm).entries());
  body.discountPercent = Number(body.discountPercent || 0);
  const response = await apiFetch("/api/admin/promos", { method: "POST", body: JSON.stringify(body) });
  const data = await response.json();
  if (!response.ok) return setStatus(data.error || "Could not save promo.", true);
  promoForm.reset();
  setStatus("Promo saved.");
  await loadAdmin();
});

settingsForm.addEventListener("submit", async (event) => {
  event.preventDefault();
  setStatus("Saving payment link...");
  const body = Object.fromEntries(new FormData(settingsForm).entries());
  const response = await apiFetch("/api/admin/settings", { method: "PATCH", body: JSON.stringify(body) });
  const data = await response.json();
  if (!response.ok) return setStatus(data.error || "Could not save settings.", true);
  settingsForm.paymentLink.value = data.settings.paymentLink || "";
  setStatus("Payment link updated.");
});

async function initAdmin() {
  try {
    const response = await fetch("/api/me", { credentials: "same-origin" });
    const data = await response.json();
    csrfToken = data.csrfToken || "";
    currentUser = data.user || null;
    if (!currentUser || currentUser.role !== "admin") {
      adminIdentity.textContent = "No admin session detected.";
      adminLocked.classList.remove("hidden");
      adminDashboard.classList.add("hidden");
      return;
    }
    adminIdentity.textContent = `${currentUser.name || currentUser.email} | ${currentUser.email} | ${currentUser.plan.toUpperCase()} owner access`;
    adminLocked.classList.add("hidden");
    adminDashboard.classList.remove("hidden");
    await loadAdmin();
  } catch {
    adminIdentity.textContent = "Unable to verify admin session.";
    adminLocked.classList.remove("hidden");
  }
}

async function apiFetch(url, options = {}) {
  const headers = { ...(options.headers || {}) };
  const isFormData = options.body instanceof FormData;
  if (!isFormData && options.body && !headers["Content-Type"]) headers["Content-Type"] = "application/json";
  if ((options.method || "GET").toUpperCase() !== "GET" && csrfToken) headers["X-CSRF-Token"] = csrfToken;
  return fetch(url, { ...options, credentials: "same-origin", headers });
}

async function loadAdmin() {
  if (!currentUser || currentUser.role !== "admin") return;
  setStatus("Refreshing...");
  const [usersResponse, promosResponse, settingsResponse] = await Promise.all([
    apiFetch("/api/admin/users"),
    apiFetch("/api/admin/promos"),
    apiFetch("/api/admin/settings")
  ]);
  const usersData = await usersResponse.json();
  const promosData = await promosResponse.json();
  const settingsData = await settingsResponse.json();
  if (!usersResponse.ok) return setStatus(usersData.error || "Could not load users.", true);
  if (!promosResponse.ok) return setStatus(promosData.error || "Could not load promos.", true);
  if (!settingsResponse.ok) return setStatus(settingsData.error || "Could not load settings.", true);

  const users = usersData.users || [];
  const promos = promosData.promos || [];
  renderAdminUsers(users);
  renderPromos(promos);
  settingsForm.paymentLink.value = settingsData.settings.paymentLink || "";
  renderKpis(users, promos);
  setStatus("Ready");
}

function renderKpis(users, promos) {
  kpiUsers.textContent = users.length;
  kpiPremium.textContent = users.filter((u) => ["premium", "enterprise"].includes(u.plan)).length;
  kpiPromos.textContent = promos.filter((p) => p.active).length;
  kpiSuspended.textContent = users.filter((u) => u.status === "suspended").length;
}

function renderAdminUsers(users) {
  adminUsers.innerHTML = users.map((u) => `
    <div class="admin-user" data-user-id="${escapeHtml(u.id)}">
      <div class="admin-user-id">
        <strong>${escapeHtml(u.name || u.email)}</strong>
        <span>${escapeHtml(u.email)}</span>
        <small>${escapeHtml(u.lastLoginAt || "No login yet")}</small>
      </div>
      <select data-field="role" title="Role"><option value="user"${u.role === "user" ? " selected" : ""}>User</option><option value="admin"${u.role === "admin" ? " selected" : ""}>Admin</option></select>
      <select data-field="plan" title="Plan"><option value="free"${u.plan === "free" ? " selected" : ""}>Free</option><option value="premium"${u.plan === "premium" ? " selected" : ""}>Premium</option><option value="enterprise"${u.plan === "enterprise" ? " selected" : ""}>Enterprise</option></select>
      <select data-field="status" title="Status"><option value="active"${u.status === "active" ? " selected" : ""}>Active</option><option value="suspended"${u.status === "suspended" ? " selected" : ""}>Suspended</option></select>
      <label class="mini-check"><input type="checkbox" data-feature="deepDive"${u.featureAccess?.deepDive ? " checked" : ""}>Deep dive</label>
      <label class="mini-check"><input type="checkbox" data-feature="aiReview"${u.featureAccess?.aiReview ? " checked" : ""}>AI</label>
      <input data-field="discountPercent" title="Discount %" type="number" min="0" max="95" value="${escapeHtml(u.discountPercent || 0)}">
      <button type="button" class="secondary" data-action="save-user">Save</button>
      <button type="button" class="danger" data-action="delete-user">Delete</button>
    </div>
  `).join("");
  adminUsers.querySelectorAll("button").forEach((button) => {
    button.addEventListener("click", async () => {
      const row = button.closest(".admin-user");
      const id = row.dataset.userId;
      if (button.dataset.action === "delete-user") {
        const response = await apiFetch(`/api/admin/users/${encodeURIComponent(id)}`, { method: "DELETE" });
        const data = await response.json();
        if (!response.ok) return setStatus(data.error || "Could not delete user.", true);
        setStatus("User deleted.");
        return loadAdmin();
      }
      const body = {
        role: row.querySelector('[data-field="role"]').value,
        plan: row.querySelector('[data-field="plan"]').value,
        status: row.querySelector('[data-field="status"]').value,
        discountPercent: Number(row.querySelector('[data-field="discountPercent"]').value || 0),
        featureAccess: {
          deepDive: row.querySelector('[data-feature="deepDive"]').checked,
          aiReview: row.querySelector('[data-feature="aiReview"]').checked
        }
      };
      const response = await apiFetch(`/api/admin/users/${encodeURIComponent(id)}`, { method: "PATCH", body: JSON.stringify(body) });
      const data = await response.json();
      if (!response.ok) return setStatus(data.error || "Could not update user.", true);
      setStatus("User access updated.");
      await loadAdmin();
    });
  });
}

function renderPromos(promos) {
  promoList.innerHTML = promos.length ? promos.map((p) => `
    <div class="promo-pill">
      <strong>${escapeHtml(p.code)}</strong>
      <span>${escapeHtml(p.discountPercent)}% ${p.active ? "active" : "inactive"}${p.expiresAt ? ` | Expires ${escapeHtml(p.expiresAt)}` : ""}</span>
      <button type="button" class="danger" data-code="${escapeHtml(p.code)}">Delete</button>
    </div>
  `).join("") : `<p class="fine-print">No promo codes yet.</p>`;
  promoList.querySelectorAll("button").forEach((button) => {
    button.addEventListener("click", async () => {
      await apiFetch(`/api/admin/promos/${encodeURIComponent(button.dataset.code)}`, { method: "DELETE" });
      setStatus("Promo deleted.");
      await loadAdmin();
    });
  });
}

function setStatus(message, isError = false) {
  adminStatus.textContent = message;
  adminStatus.classList.toggle("error", Boolean(isError));
}

function escapeHtml(value) {
  return String(value).replace(/[<>&'"]/g, (c) => ({ "<": "&lt;", ">": "&gt;", "&": "&amp;", "'": "&#39;", '"': "&quot;" }[c]));
}

function initAdminEffects() {
  document.documentElement.classList.add("js-ready");
  const progress = document.querySelector(".scroll-progress");
  function updateProgress() {
    const max = document.documentElement.scrollHeight - window.innerHeight;
    const pct = max > 0 ? Math.min(100, Math.max(0, window.scrollY / max * 100)) : 0;
    if (progress) progress.style.width = `${pct}%`;
  }
  updateProgress();
  window.addEventListener("scroll", updateProgress, { passive: true });
  const revealObserver = new IntersectionObserver((entries) => {
    entries.forEach((entry) => {
      if (entry.isIntersecting) entry.target.classList.add("in-view");
    });
  }, { threshold: 0.14 });
  document.querySelectorAll(".reveal").forEach((el) => revealObserver.observe(el));
  requestAnimationFrame(() => {
    document.querySelectorAll(".reveal").forEach((el) => {
      const rect = el.getBoundingClientRect();
      if (rect.top < window.innerHeight * 0.92) el.classList.add("in-view");
    });
  });
}
