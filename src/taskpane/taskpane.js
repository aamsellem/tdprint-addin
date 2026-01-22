import "./taskpane.css";
import { categories, getAllVariables, getTypeName, getTypeColor } from "../data/variables.js";

// Clé de stockage pour les favoris
const FAVORITES_KEY = "tdprint_favorites";

// État de l'application
let favorites = [];
let searchQuery = "";

/**
 * Initialisation de l'application
 */
Office.onReady((info) => {
  // Initialiser dans Word ou en mode standalone (navigateur)
  init();
});

// Fallback si Office.js n'est pas disponible (test navigateur)
if (typeof Office === "undefined") {
  document.addEventListener("DOMContentLoaded", init);
}

/**
 * Initialisation
 */
function init() {
  loadFavorites();
  renderCategories();
  updateFavoritesSection();
  setupEventListeners();
}

/**
 * Configuration des écouteurs d'événements
 */
function setupEventListeners() {
  const searchInput = document.getElementById("searchInput");
  const clearSearchBtn = document.getElementById("clearSearch");

  // Recherche en temps réel
  searchInput.addEventListener("input", (e) => {
    searchQuery = e.target.value.trim();
    clearSearchBtn.classList.toggle("visible", searchQuery.length > 0);
    filterVariables();
  });

  // Effacer la recherche
  clearSearchBtn.addEventListener("click", () => {
    searchInput.value = "";
    searchQuery = "";
    clearSearchBtn.classList.remove("visible");
    filterVariables();
  });

  // Raccourci clavier pour la recherche
  document.addEventListener("keydown", (e) => {
    if (e.key === "Escape" && searchQuery) {
      searchInput.value = "";
      searchQuery = "";
      clearSearchBtn.classList.remove("visible");
      filterVariables();
    }
    if ((e.ctrlKey || e.metaKey) && e.key === "f") {
      e.preventDefault();
      searchInput.focus();
    }
  });
}

/**
 * Charger les favoris depuis localStorage
 */
function loadFavorites() {
  try {
    const stored = localStorage.getItem(FAVORITES_KEY);
    favorites = stored ? JSON.parse(stored) : [];
  } catch (e) {
    console.error("Erreur de chargement des favoris:", e);
    favorites = [];
  }
}

/**
 * Sauvegarder les favoris dans localStorage
 */
function saveFavorites() {
  try {
    localStorage.setItem(FAVORITES_KEY, JSON.stringify(favorites));
  } catch (e) {
    console.error("Erreur de sauvegarde des favoris:", e);
  }
}

/**
 * Basculer le statut favori d'une variable
 */
function toggleFavorite(placeholder) {
  const index = favorites.indexOf(placeholder);
  if (index === -1) {
    favorites.push(placeholder);
    showToast("Ajouté aux favoris", "success");
  } else {
    favorites.splice(index, 1);
    showToast("Retiré des favoris");
  }
  saveFavorites();
  updateFavoritesSection();
  updateFavoriteButtons();
}

/**
 * Vérifier si une variable est en favori
 */
function isFavorite(placeholder) {
  return favorites.includes(placeholder);
}

/**
 * Mettre à jour les boutons favoris dans l'interface
 */
function updateFavoriteButtons() {
  document.querySelectorAll(".favorite-btn").forEach((btn) => {
    const placeholder = btn.dataset.placeholder;
    btn.classList.toggle("active", isFavorite(placeholder));
  });
}

/**
 * Mettre à jour la section favoris
 */
function updateFavoritesSection() {
  const section = document.getElementById("favoritesSection");
  const container = document.getElementById("favoritesContainer");
  const countBadge = document.getElementById("favoritesCount");

  if (favorites.length === 0) {
    section.classList.add("hidden");
    return;
  }

  section.classList.remove("hidden");
  countBadge.textContent = favorites.length;

  // Récupérer les variables favorites
  const allVariables = getAllVariables();
  const favoriteVariables = favorites
    .map((placeholder) => allVariables.find((v) => v.placeholder === placeholder))
    .filter(Boolean);

  container.innerHTML = favoriteVariables
    .map((variable) => createVariableItemHTML(variable, true))
    .join("");

  // Attacher les événements
  attachVariableEvents(container);
}

/**
 * Rendre les catégories
 */
function renderCategories() {
  const container = document.getElementById("categoriesContainer");

  container.innerHTML = categories
    .map(
      (category) => `
      <div class="category" data-category-id="${category.id}">
        <div class="category-header" onclick="toggleCategory('${category.id}')">
          <div class="category-header-left">
            <span class="category-icon">
              <svg width="16" height="16" viewBox="0 0 16 16" fill="currentColor">
                <path d="M6 4l4 4-4 4V4z"/>
              </svg>
            </span>
            <span class="category-name">${category.name}</span>
          </div>
          <span class="badge-count">${category.variables.length}</span>
        </div>
        <div class="category-content">
          <div class="variables-list">
            ${category.variables.map((v) => createVariableItemHTML({ ...v, categoryName: category.name })).join("")}
          </div>
        </div>
      </div>
    `
    )
    .join("");

  // Attacher les événements pour toutes les catégories
  categories.forEach((category) => {
    const categoryEl = container.querySelector(`[data-category-id="${category.id}"]`);
    const variablesList = categoryEl.querySelector(".variables-list");
    attachVariableEvents(variablesList);
  });
}

/**
 * Créer le HTML d'un élément variable
 */
function createVariableItemHTML(variable, showCategory = false) {
  const isFav = isFavorite(variable.placeholder);
  const typeClass = `type-${variable.type}`;
  const typeName = getTypeName(variable.type);

  return `
    <div class="variable-item"
         draggable="true"
         data-placeholder="${variable.placeholder}"
         data-type="${variable.type}">
      <div class="variable-main">
        <div class="variable-header">
          <span class="variable-placeholder">${highlightText(variable.placeholder, searchQuery)}</span>
          <span class="variable-type ${typeClass}" title="${typeName}">${variable.type}</span>
        </div>
        <div class="variable-description">${highlightText(variable.description, searchQuery)}</div>
        ${showCategory && variable.categoryName ? `<div class="variable-category">${highlightText(variable.categoryName, searchQuery)}</div>` : ""}
      </div>
      <button class="favorite-btn ${isFav ? "active" : ""}"
              data-placeholder="${variable.placeholder}"
              title="${isFav ? "Retirer des favoris" : "Ajouter aux favoris"}">
        <svg viewBox="0 0 16 16" fill="currentColor">
          ${isFav
            ? '<path d="M8 1.5l1.854 3.758 4.146.603-3 2.923.708 4.13L8 11.062l-3.708 1.852.708-4.13-3-2.923 4.146-.603L8 1.5z"/>'
            : '<path d="M8 1.5l1.854 3.758 4.146.603-3 2.923.708 4.13L8 11.062l-3.708 1.852.708-4.13-3-2.923 4.146-.603L8 1.5zm0 2.445L6.764 6.36l-2.682.39 1.941 1.891-.458 2.67L8 9.857l2.435 1.28-.458-2.67 1.941-1.891-2.682-.39L8 3.945z"/>'
          }
        </svg>
      </button>
      <button class="insert-btn" data-placeholder="${variable.placeholder}" title="Insérer">
        Insérer
      </button>
    </div>
  `;
}

/**
 * Attacher les événements aux éléments variables
 */
function attachVariableEvents(container) {
  // Événements de drag & drop
  container.querySelectorAll(".variable-item").forEach((item) => {
    item.addEventListener("dragstart", handleDragStart);
    item.addEventListener("dragend", handleDragEnd);
  });

  // Événements de clic sur favoris
  container.querySelectorAll(".favorite-btn").forEach((btn) => {
    btn.addEventListener("click", (e) => {
      e.stopPropagation();
      toggleFavorite(btn.dataset.placeholder);
    });
  });

  // Événements de clic sur Insérer
  container.querySelectorAll(".insert-btn").forEach((btn) => {
    btn.addEventListener("click", (e) => {
      e.stopPropagation();
      insertVariable(btn.dataset.placeholder);
    });
  });
}

/**
 * Basculer l'état d'une catégorie (ouverte/fermée)
 */
window.toggleCategory = function (categoryId) {
  const categoryEl = document.querySelector(`[data-category-id="${categoryId}"]`);
  if (categoryEl) {
    categoryEl.classList.toggle("expanded");
  }
};

/**
 * Filtrer les variables selon la recherche
 */
function filterVariables() {
  const noResults = document.getElementById("noResults");
  const query = searchQuery.toLowerCase();
  let hasVisibleItems = false;

  // Filtrer les catégories
  categories.forEach((category) => {
    const categoryEl = document.querySelector(`[data-category-id="${category.id}"]`);
    const variablesList = categoryEl.querySelector(".variables-list");
    let categoryHasVisibleItems = false;

    // Filtrer les variables dans cette catégorie
    const allVariables = category.variables;
    allVariables.forEach((variable) => {
      const itemEl = variablesList.querySelector(`[data-placeholder="${variable.placeholder}"]`);
      if (!itemEl) return;

      const matches =
        !query ||
        variable.placeholder.toLowerCase().includes(query) ||
        variable.description.toLowerCase().includes(query) ||
        category.name.toLowerCase().includes(query);

      itemEl.style.display = matches ? "" : "none";
      if (matches) {
        categoryHasVisibleItems = true;
        hasVisibleItems = true;
      }
    });

    // Afficher/masquer la catégorie
    categoryEl.style.display = categoryHasVisibleItems ? "" : "none";

    // Si recherche active et il y a des résultats, ouvrir la catégorie
    if (query && categoryHasVisibleItems) {
      categoryEl.classList.add("expanded");
    }

    // Re-rendre les variables pour mettre à jour la surbrillance
    if (query) {
      variablesList.innerHTML = category.variables
        .map((v) => createVariableItemHTML({ ...v, categoryName: category.name }))
        .join("");
      attachVariableEvents(variablesList);

      // Réappliquer le filtre de visibilité
      category.variables.forEach((variable) => {
        const itemEl = variablesList.querySelector(`[data-placeholder="${variable.placeholder}"]`);
        if (!itemEl) return;
        const matches =
          variable.placeholder.toLowerCase().includes(query) ||
          variable.description.toLowerCase().includes(query) ||
          category.name.toLowerCase().includes(query);
        itemEl.style.display = matches ? "" : "none";
      });
    }
  });

  // Filtrer les favoris
  updateFavoritesSection();
  if (query) {
    const favContainer = document.getElementById("favoritesContainer");
    const favItems = favContainer.querySelectorAll(".variable-item");
    favItems.forEach((item) => {
      const placeholder = item.dataset.placeholder;
      const variable = getAllVariables().find((v) => v.placeholder === placeholder);
      if (!variable) return;

      const matches =
        variable.placeholder.toLowerCase().includes(query) ||
        variable.description.toLowerCase().includes(query) ||
        (variable.categoryName && variable.categoryName.toLowerCase().includes(query));

      item.style.display = matches ? "" : "none";
      if (matches) hasVisibleItems = true;
    });
  }

  // Afficher/masquer le message "aucun résultat"
  noResults.classList.toggle("hidden", hasVisibleItems || !query);
}

/**
 * Surbrillance du texte recherché
 */
function highlightText(text, query) {
  if (!query) return text;
  const regex = new RegExp(`(${escapeRegExp(query)})`, "gi");
  return text.replace(regex, '<span class="highlight">$1</span>');
}

/**
 * Échapper les caractères spéciaux pour regex
 */
function escapeRegExp(string) {
  return string.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

/**
 * Gestion du drag start
 */
function handleDragStart(e) {
  const placeholder = e.target.dataset.placeholder || e.target.closest(".variable-item").dataset.placeholder;
  e.dataTransfer.setData("text/plain", placeholder);
  e.dataTransfer.effectAllowed = "copy";
  e.target.classList.add("dragging");
}

/**
 * Gestion du drag end
 */
function handleDragEnd(e) {
  e.target.classList.remove("dragging");
}

/**
 * Insérer une variable dans le document Word
 */
async function insertVariable(placeholder) {
  // Vérifier si Word est disponible
  if (typeof Word === "undefined" || !Word.run) {
    // Mode navigateur : copier dans le presse-papier
    try {
      await navigator.clipboard.writeText(placeholder);
      showToast(`Copié: ${placeholder}`, "success");
    } catch (e) {
      showToast(`Variable: ${placeholder}`, "success");
    }
    return;
  }

  try {
    await Word.run(async (context) => {
      const selection = context.document.getSelection();
      selection.insertText(placeholder, Word.InsertLocation.replace);
      await context.sync();
    });
    showToast(`Variable insérée: ${placeholder}`, "success");
  } catch (error) {
    console.error("Erreur d'insertion:", error);
    showToast("Erreur lors de l'insertion", "error");
  }
}

/**
 * Afficher une notification toast
 */
function showToast(message, type = "") {
  // Supprimer les toasts existants
  const existingToast = document.querySelector(".toast");
  if (existingToast) {
    existingToast.remove();
  }

  const toast = document.createElement("div");
  toast.className = `toast ${type}`;
  toast.textContent = message;
  document.body.appendChild(toast);

  // Afficher
  requestAnimationFrame(() => {
    toast.classList.add("show");
  });

  // Masquer après 2 secondes
  setTimeout(() => {
    toast.classList.remove("show");
    setTimeout(() => toast.remove(), 300);
  }, 2000);
}

// Gestion du drop dans Word (via Office.js)
// Note: Le glisser-déposer direct dans Word depuis le task pane
// utilise l'événement de drag & drop standard qui insère le texte
// Le texte est automatiquement inséré grâce à setData("text/plain", ...)
