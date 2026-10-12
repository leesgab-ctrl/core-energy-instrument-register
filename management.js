const escapeHtml = value => String(value ?? "").replace(/[&<>"']/g, character => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" })[character]);
export function cycleLabel(months) {
  return months === "별도" || months === "별도참조" ? "별도" : Number(months) > 0 ? `${months}개월` : "미지정";
}
export function categoryMarkup(data, compare) {
  const categories = [...(data.categories || [])].sort((a, b) => compare(a.basis || a.code, b.basis || b.code));
  return categories.map(category => {
    const count = data.instruments.filter(item => item.categoryCode === category.code).length;
    return `<div class="public-category"><strong>${escapeHtml(category.name)}</strong><span class="category-basis">${escapeHtml(category.basis || category.code)}</span><span class="category-cycle">${cycleLabel(category.months)}</span><small>${count}대</small></div>`;
  }).join("") || '<p class="management-empty">등록된 장비분류가 없습니다.</p>';
}
export function ownerMarkup(data, compare) {
  const groups = new Map();
  for (const item of data.instruments) {
    const owner = String(item.owner || "").trim() || "미지정";
    if (!groups.has(owner)) groups.set(owner, []);
    groups.get(owner).push(item);
  }
  return [...groups.entries()].sort(([a], [b]) => a === b ? 0 : a === "미지정" ? 1 : b === "미지정" ? -1 : a.localeCompare(b, "ko"))
    .map(([owner, items]) => `<section class="owner-group"><h3>${escapeHtml(owner)} <small>${items.length}대</small></h3>${items.sort((a, b) => compare(a.no, b.no)).map(item => `<button type="button" class="owner-equipment" data-id="${escapeHtml(item.id)}"><span class="code${item.requiredInspection ? " required" : ""}">${escapeHtml(item.no)}</span><strong>${escapeHtml(item.name)}</strong><span class="owner-serial">SN ${escapeHtml(item.serial || "미입력")}</span></button>`).join("")}</section>`).join("") || '<p class="management-empty">등록된 장비가 없습니다.</p>';
}
export function setupManagement(getData, compare) {
  document.addEventListener("click", event => {
    const open = event.target.closest("[data-management-open]");
    if (open) {
      const owner = open.dataset.managementOpen === "owner";
      const dialog = document.getElementById(owner ? "ownerStatusDialog" : "categoryStatusDialog");
      document.getElementById(owner ? "ownerStatusList" : "categoryStatusList").innerHTML = owner ? ownerMarkup(getData(), compare) : categoryMarkup(getData(), compare);
      document.getElementById("publicManagementMenu").open = false;
      dialog.showModal();
    }
    const close = event.target.closest("[data-management-close]");
    if (close) close.closest("dialog").close();
    if (!event.target.closest("#publicManagementMenu")) document.getElementById("publicManagementMenu").open = false;
  });
  return () => document.querySelectorAll("[data-management-open]").forEach(button => { button.disabled = !Array.isArray(getData().categories); });
}
