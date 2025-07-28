let filteredApps = [...applications];
let currentSort = "name-asc";

function initializeFilters() {
  // Get unique categories
  const categories = [
    ...new Set(applications.map((app) => app.category)),
  ].sort();
  const categorySelect = document.getElementById("categorySelect");
  categories.forEach((cat) => {
    const option = document.createElement("option");
    option.value = cat;
    option.textContent = cat;
    categorySelect.appendChild(option);
  });

  // Get unique publishers
  const publishers = [
    ...new Set(applications.map((app) => app.publisher)),
  ].sort();
  const publisherSelect = document.getElementById("publisherSelect");
  publishers.forEach((pub) => {
    const option = document.createElement("option");
    option.value = pub;
    option.textContent = pub;
    publisherSelect.appendChild(option);
  });
}

function filterApplications() {
  const searchTerm = document.getElementById("searchInput").value.toLowerCase();
  const selectedCategory = document.getElementById("categorySelect").value;
  const selectedPublisher = document.getElementById("publisherSelect").value;
  currentSort = document.getElementById("sortSelect").value;

  filteredApps = applications.filter((app) => {
    const matchesSearch =
      app.name.toLowerCase().includes(searchTerm) ||
      app.publisher.toLowerCase().includes(searchTerm);
    const matchesCategory =
      selectedCategory === "all" || app.category === selectedCategory;
    const matchesPublisher =
      selectedPublisher === "all" || app.publisher === selectedPublisher;

    return matchesSearch && matchesCategory && matchesPublisher;
  });

  renderApplications();
  updateResultsCount();
}

function updateResultsCount() {
  document.getElementById("resultsCount").textContent =
    `Found ${filteredApps.length} applications`;
}

function resetFilters() {
  document.getElementById("searchInput").value = "";
  document.getElementById("categorySelect").value = "all";
  document.getElementById("publisherSelect").value = "all";
  document.getElementById("sortSelect").value = "name-asc";
  filterApplications();
}
function sortApps(arr) {
  switch (currentSort) {
    case "name-desc":
      return arr.sort((a, b) => b.name.localeCompare(a.name));
    case "publisher-asc":
      return arr.sort((a, b) => a.publisher.localeCompare(b.publisher));
    case "publisher-desc":
      return arr.sort((a, b) => b.publisher.localeCompare(a.publisher));
    default:
      return arr.sort((a, b) => a.name.localeCompare(b.name));
  }
}

function createCopyButton(url, label) {
  const btn = document.createElement("button");
  btn.type = "button";
  btn.className = "copy-btn";
  btn.setAttribute("aria-label", `Copy ${label} link`);
  btn.innerHTML = `
        <svg fill="none" stroke="currentColor" viewBox="0 0 24 24">
            <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M8 2h8a2 2 0 012 2v12a2 2 0 01-2 2H8a2 2 0 01-2-2V4a2 2 0 012-2zm-3 6H4a2 2 0 00-2 2v10a2 2 0 002 2h10a2 2 0 002-2v-1"/>
        </svg>
        Copy
    `;
  btn.addEventListener("click", async () => {
    try {
      await navigator.clipboard.writeText(url);
      const original = btn.textContent;
      btn.textContent = "Copied!";
      setTimeout(() => (btn.textContent = original), 2000);
    } catch (err) {
      alert("Failed to copy link");
    }
  });
  return btn;
}

function renderApplications() {
  const container = document.getElementById("appsContainer");
  container.innerHTML = "";

  if (filteredApps.length === 0) {
    container.innerHTML =
      '<div class="no-results">No applications found matching your criteria.</div>';
    return;
  }

  // Group by category
  const grouped = filteredApps.reduce((acc, app) => {
    if (!acc[app.category]) acc[app.category] = [];
    acc[app.category].push(app);
    return acc;
  }, {});

  // Render each category
  Object.keys(grouped)
    .sort()
    .forEach((category) => {
      const categorySection = document.createElement("div");
      categorySection.className = "category-section";

      const categoryTitle = document.createElement("h2");
      categoryTitle.className = "category-title";
      categoryTitle.textContent = category;
      categorySection.appendChild(categoryTitle);

      const appGrid = document.createElement("div");
      appGrid.className = "app-grid";
      sortApps(grouped[category]).forEach((app) => {
        const appCard = document.createElement("div");
        appCard.className = "app-card";

        const appName = document.createElement("h3");
        appName.className = "app-name";
        appName.textContent = app.name;

        const appPublisher = document.createElement("p");
        appPublisher.className = "app-publisher";
        appPublisher.textContent = app.publisher;

        const buttonGroup = document.createElement("div");
        buttonGroup.className = "button-group";

        // x86/x64 button
        const downloadBtn = document.createElement("a");
        downloadBtn.href = app.msiUrl;
        downloadBtn.target = "_blank";
        downloadBtn.rel = "noopener noreferrer";
        downloadBtn.className = "download-btn";
        downloadBtn.setAttribute(
          "aria-label",
          `Download ${app.name} for x86/x64`,
        );
        downloadBtn.innerHTML = `
            <svg fill="none" stroke="currentColor" viewBox="0 0 24 24">
                <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M7 16a4 4 0 01-.88-7.903A5 5 0 1115.9 6L16 6a5 5 0 011 9.9M9 19l3 3m0 0l3-3m-3 3V10"/>
            </svg>
            x86/x64
        `;
        buttonGroup.appendChild(downloadBtn);
        buttonGroup.appendChild(createCopyButton(app.msiUrl, `${app.name} x86/x64`));

        // ARM button (if available)
        if (app.armUrl) {
          const armBtn = document.createElement("a");
          armBtn.href = app.armUrl;
          armBtn.target = "_blank";
          armBtn.rel = "noopener noreferrer";
          armBtn.className = "download-btn arm";
          armBtn.setAttribute("aria-label", `Download ${app.name} for ARM64`);
          armBtn.innerHTML = `
                <svg fill="none" stroke="currentColor" viewBox="0 0 24 24">
                    <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M7 16a4 4 0 01-.88-7.903A5 5 0 1115.9 6L16 6a5 5 0 011 9.9M9 19l3 3m0 0l3-3m-3 3V10"/>
                </svg>
                ARM64
            `;
          buttonGroup.appendChild(armBtn);
          buttonGroup.appendChild(createCopyButton(app.armUrl, `${app.name} ARM64`));
        }

        appCard.appendChild(appName);
        appCard.appendChild(appPublisher);
        appCard.appendChild(buttonGroup);
        appGrid.appendChild(appCard);
      });

      categorySection.appendChild(appGrid);
      container.appendChild(categorySection);
    });
}

// Initialize the app
document.addEventListener("DOMContentLoaded", function () {
  initializeFilters();
  renderApplications();
  updateResultsCount();

  // Add event listeners
  document
    .getElementById("searchInput")
    .addEventListener("input", filterApplications);
  document
    .getElementById("categorySelect")
    .addEventListener("change", filterApplications);
  document
    .getElementById("publisherSelect")
    .addEventListener("change", filterApplications);
  document
    .getElementById("sortSelect")
    .addEventListener("change", filterApplications);

  document
    .getElementById("resetFilters")
    .addEventListener("click", resetFilters);
});
