const FIXED_SUPABASE = {
  projectRef: "qgcbgdhhequmuijbvshg",
  anonKey: "sb_publishable_WzEy7vCPL-ztzJqa_XzLfQ_36GfAd4a"
};

const THEME_KEY = "rivet_theme";
const VIEW_IDS = ["accueil", "bons", "bons-anciens", "clients", "projets", "factures", "depenses", "configuration"];
const CLIENT_TYPE_OPTIONS = ["Personnel", "Commercial"];
const ADDRESS_SUGGEST_MIN_CHARS = 3;
const ADDRESS_SUGGEST_LIMIT = 6;
const ADDRESS_SUGGEST_DEBOUNCE_MS = 220;
const FACTURE_DUE_DAYS = 30;
const FACTURES_STORAGE_BUCKET = "factures";
const EXPENSES_STORAGE_BUCKET = "expenses";
const FISCAL_REPORTS_BUCKET = "fiscal-reports";
const REPAIRS_PHOTOS_BUCKET = "photos";
const PROJECT_FILES_BUCKET = "projects";
const REPAIR_THUMBNAIL_OUTPUT_WIDTH = 1400;
const PDFJS_WORKER_SRC = "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js";
const DEFAULT_FACTURE_TPS_RATE = 0.05;
const DEFAULT_FACTURE_TVQ_RATE = 0.09975;
const PURCHASE_LINK_PRESETS = Object.freeze([
  {
    key: "digikey",
    title: "DigiKey",
    domain: "digikey.ca",
    configKey: "purchase_link_digikey",
    defaultUrl: "https://www.digikey.ca/",
  },
  {
    key: "mouser",
    title: "Mouser",
    domain: "mouser.ca",
    configKey: "purchase_link_mouser",
    defaultUrl: "https://www.mouser.ca/",
  },
  {
    key: "nextgen_guitars",
    title: "NextGen Guitars",
    domain: "nextgenguitars.ca",
    configKey: "purchase_link_nextgen_guitars",
    defaultUrl: "https://nextgenguitars.ca/",
  },
  {
    key: "thetubestore",
    title: "TheTubeStore",
    domain: "thetubestore.com",
    configKey: "purchase_link_thetubestore",
    defaultUrl: "https://www.thetubestore.com/",
  },
]);
const DEFAULT_PURCHASE_LINKS = Object.freeze(
  PURCHASE_LINK_PRESETS.map((preset) => ({
    key: preset.key,
    title: preset.title,
    url: preset.defaultUrl,
    note: "",
    domain: preset.domain,
  }))
);
const DEFAULT_APP_CONFIG = Object.freeze({
  id: 1,
  labor_rate_standard: 95,
  labor_rate_secondary: 125,
  labor_rate_diagnostic: 65,
  labor_rate_estimation: 65,
  developer_mode: false,
  podio_sync_enabled: true,
  company_name: "",
  company_phone: "",
  company_email: "",
  company_address_line1: "",
  company_address_line2: "",
  company_city: "",
  company_province: "",
  company_postal_code: "",
  company_country: "Canada",
  company_logo_path: "assets/facturation/rivet-logo-bk.png",
  purchase_links: DEFAULT_PURCHASE_LINKS,
  purchase_link_digikey: "https://www.digikey.ca/",
  purchase_link_mouser: "https://www.mouser.ca/",
  purchase_link_nextgen_guitars: "https://nextgenguitars.ca/",
  purchase_link_thetubestore: "https://www.thetubestore.com/",
  tax_tps_number: "",
  tax_tvq_number: "",
  tax_tps_rate: DEFAULT_FACTURE_TPS_RATE,
  tax_tvq_rate: DEFAULT_FACTURE_TVQ_RATE,
});

const BONS_COLUMNS = [
  { key: "dossier_ouvert", label: "Dossier ouvert", statusLabel: "Dossier ouvert" },
  { key: "en_test", label: "En test", statusLabel: "En test" },
  { key: "en_cours", label: "En cours", statusLabel: "En cours" },
  { key: "en_attente_pieces", label: "En attente de pièces", statusLabel: "En attente de pièces" },
  { key: "termine", label: "Terminée", statusLabel: "Terminée" },
  { key: "remis_client", label: "Remis au client", statusLabel: "Remis au client" },
  { key: "non_reparable", label: "Non réparable", statusLabel: "Non réparable" },
  { key: "annule", label: "Annulé", statusLabel: "Annulé" }
];

const BONS_DEFAULT_COLUMN_KEYS = new Set([
  "dossier_ouvert",
  "en_test",
  "en_cours",
  "en_attente_pieces",
  "termine",
]);

const REPAIR_STATUS_OPTIONS = [
  "Dossier ouvert",
  "En test",
  "En cours",
  "En attente de pièces",
  "Terminée",
  "Remis au client",
  "Non réparable",
  "Annulé"
];

const REPAIR_STATUS_TONE_COLORS = {
  "dossier-ouvert": "#d26d88",
  "en-test": "#9e7ad6",
  "en-cours": "#2ca9b7",
  "en-attente-pieces": "#ad8514",
  "terminee": "#2f9e58",
  "remis-client": "#74b984",
  "non-reparable": "#787d84",
  "annule": "#a8adb3",
  "default": "#586172"
};

const REPAIR_PARTS_SUPPLIERS = PURCHASE_LINK_PRESETS;

const FACTURATION_TABS = new Set(["factures", "estimations", "soumissions"]);
const BONS_PROJECTS_TABS = new Set(["bons", "projets"]);
const BONS_REPAIRS_TABS = new Set(["board", "archives"]);
const PROJECTS_TABS = new Set(["board", "archives"]);
const PROJECT_STAGE_OPENING_LABEL = "Ouverture";
const PROJECT_STAGE_EMIT_INVOICE_LABEL = "Émettre la facture";

const state = {
  authClient: null,
  session: null,
  sessionReady: false,
  theme: "light",
  profile: null,
  profileUserId: null,
  dashboardLoadedFor: null,
  repairs: [],
  projects: [],
  factures: [],
  expenses: [],
  estimations: [],
  taxReports: [],
  clients: [],
  appConfig: null,
  appConfigLoaded: false,
  configEditing: {
    honoraires: false,
    entreprise: false,
    achats: false,
    developpement: false,
  },
  configDeveloperDraft: null,
  configPodioSyncDraft: null,
  configPurchaseLinksDraft: null,
  configActiveTab: "honoraires",
  clientsByPodioItemId: {},
  factureByRepairItemId: {},
  estimationsLoadError: null,
  expensesLoadError: null,
  clientsLoadError: null,
  projectsLoadError: null,
  taxReportsLoadError: null,
  taxReportsLoading: false,
  taxReportsGeneratingYear: null,
  bonsProjectsActiveTab: "bons",
  bonsRepairsActiveTab: "board",
  projectsActiveTab: "board",
  facturationActiveTab: "factures",
  clientsSearchQuery: "",
  projectsSearchQuery: "",
  facturesSearchQuery: "",
  devisSearchQuery: "",
  facturesYearFilter: "all",
  facturesStatusFilter: "all",
  devisStatusFilter: "all",
  bonsOtherStatusFilter: "other",
  bonsOtherSearchQuery: "",
  bonsDrag: {
    activeRepairId: null,
    sourceColumnKey: null,
    overColumnKey: null,
    dropTargetKey: null,
    isSaving: false
  },
  bonsDragIgnoreClickUntil: 0
};

const $ = (id) => document.getElementById(id);

const clientModalState = {
  clientId: null,
  original: null,
  draft: null,
  editing: false,
  saving: false,
  message: "",
  isCreating: false,
  onCreated: null,
  addressParts: null,
  addressSuggestionsByValue: {},
  addressSuggestTimer: null,
  addressSuggestSeq: 0
};

const factureModalState = {
  open: false,
};

const projectModalState = {
  open: false,
};

const expenseModalState = {
  open: false,
};

const facturePdfModalState = {
  open: false,
  title: "Apercu PDF",
  pdfUrl: null,
  doc: null,
  page: 1,
  totalPages: 0,
  scale: 1.1,
  renderToken: 0,
  loading: false,
};

let repairUnsavedDialogPromise = null;
let projectProceedToFactureDialogPromise = null;
let repairFactureCreateDialogPromise = null;
let projectFactureCreateDialogPromise = null;
let repairPartsSuppliersDialogOpen = false;
let localEntitySequence = 0;
let configMessageTimer = null;
const repairThumbUrlCache = new Map();
const repairThumbLoadingKeys = new Set();
const projectThumbUrlCache = new Map();
const projectThumbLoadingKeys = new Set();

const EYE_OPEN_SVG = `
<svg width="18" height="18" viewBox="0 0 24 24" fill="none" aria-hidden="true">
  <path d="M2 12s3.5-6 10-6 10 6 10 6-3.5 6-10 6-10-6-10-6Z" stroke="currentColor" stroke-width="1.7"/>
  <circle cx="12" cy="12" r="3" stroke="currentColor" stroke-width="1.7"/>
</svg>`;

const EYE_CLOSED_SVG = `
<svg width="18" height="18" viewBox="0 0 24 24" fill="none" aria-hidden="true">
  <path d="M3 3 21 21" stroke="currentColor" stroke-width="1.7" stroke-linecap="round"/>
  <path d="M10.7 6.2A10.4 10.4 0 0 1 12 6c6.5 0 10 6 10 6a18.8 18.8 0 0 1-3.4 4.1" stroke="currentColor" stroke-width="1.7" stroke-linecap="round"/>
  <path d="M6.4 8.1A18.4 18.4 0 0 0 2 12s3.5 6 10 6c1.2 0 2.2-.2 3.2-.5" stroke="currentColor" stroke-width="1.7" stroke-linecap="round"/>
</svg>`;

const FACTURE_PREVIEW_EYE_SVG = `
<svg viewBox="0 0 24 24" aria-hidden="true" focusable="false">
  <path d="M2 12s3.5-6 10-6 10 6 10 6-3.5 6-10 6-10-6-10-6Z" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round"></path>
  <circle cx="12" cy="12" r="1.7" fill="currentColor"></circle>
</svg>`;

const FACTURE_EDIT_PENCIL_SVG = `
<svg viewBox="0 0 24 24" aria-hidden="true" focusable="false">
  <path d="M4.5 19.5h3.8l10-10-3.8-3.8-10 10v3.8Z" fill="none" stroke="currentColor" stroke-width="1.7" stroke-linecap="round" stroke-linejoin="round"></path>
  <path d="M12.9 6.1 17 10.2" fill="none" stroke="currentColor" stroke-width="1.7" stroke-linecap="round"></path>
</svg>`;

const REPAIR_THUMB_ICON_SVG = `
<svg viewBox="0 0 24 24" aria-hidden="true" focusable="false">
  <rect x="3.5" y="5.5" width="17" height="13" rx="2.2" ry="2.2" fill="none" stroke="currentColor" stroke-width="1.7"/>
  <path d="M6.5 15.5 10 12l2.2 2.1 2.8-3 2.5 4.4" fill="none" stroke="currentColor" stroke-width="1.7" stroke-linecap="round" stroke-linejoin="round"/>
  <circle cx="9" cy="9.5" r="1.1" fill="currentColor"/>
</svg>`;

const REPAIR_AGE_ALARM_SVG = `
<svg viewBox="0 0 24 24" aria-hidden="true" focusable="false">
  <path d="M7.1 5.7 5 7.8M16.9 5.7 19 7.8" fill="none" stroke="currentColor" stroke-width="1.7" stroke-linecap="round"/>
  <path d="M8.2 20 9.6 18.3M15.8 20 14.4 18.3" fill="none" stroke="currentColor" stroke-width="1.7" stroke-linecap="round"/>
  <circle cx="12" cy="12.6" r="5.9" fill="none" stroke="currentColor" stroke-width="1.7"/>
  <path d="M12 10v2.8l2 1.6" fill="none" stroke="currentColor" stroke-width="1.7" stroke-linecap="round" stroke-linejoin="round"/>
</svg>`;

const REPAIR_PARTS_CART_SVG = `
<svg viewBox="0 0 24 24" aria-hidden="true" focusable="false">
  <path d="M3.8 5.2h2.4l1.6 9.2h9.2l2-6.9H7.4" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"/>
  <circle cx="10.2" cy="18.3" r="1.5" fill="none" stroke="currentColor" stroke-width="1.7"/>
  <circle cx="16.7" cy="18.3" r="1.5" fill="none" stroke="currentColor" stroke-width="1.7"/>
</svg>`;

function escapeHtml(value) {
  return String(value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}

function normalizeText(value) {
  return String(value ?? "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toLowerCase()
    .trim();
}

function isMissingFacturesNumeroColumnError(error) {
  if (!error || typeof error !== "object") return false;
  const code = String(error.code ?? "").trim();
  const message = String(error.message ?? "").toLowerCase();
  const details = String(error.details ?? "").toLowerCase();
  const hint = String(error.hint ?? "").toLowerCase();
  const text = `${message} ${details} ${hint}`;
  if (code !== "42703" && !text.includes("does not exist")) return false;
  return text.includes("numero") && text.includes("factures");
}

function htmlToText(value) {
  if (value == null) return "";
  const container = document.createElement("div");
  container.innerHTML = String(value);
  return (container.textContent || container.innerText || "").trim();
}

function formatDateTime(value) {
  if (!value) return "";
  const d = new Date(value);
  if (Number.isNaN(d.getTime())) return "";
  return d.toLocaleString("fr-CA", {
    year: "numeric",
    month: "2-digit",
    day: "2-digit",
    hour: "2-digit",
    minute: "2-digit"
  });
}

function formatMoney(value) {
  const amount = Number(value || 0);
  return amount.toLocaleString("fr-CA", {
    style: "currency",
    currency: "CAD",
    minimumFractionDigits: 2
  });
}

function formatMoneyCompact(value) {
  if (value == null || String(value).trim() === "") return "";
  const amount = Number(value);
  if (!Number.isFinite(amount)) return "";
  const full = formatMoney(amount);
  return full.replace(/,00(?=\s*\$)/, "");
}

const PROJECT_COST_DEFAULT_CURRENCY = "CAD";
const PROJECT_COST_CURRENCIES = ["CAD", "USD", "GBP", "EUR", "JPY"];
const PROJECT_COST_FX_API_BASE = "https://api.frankfurter.app";
const PROJECT_COST_PAYPAL_SPREAD = 0.04;
const PROJECT_COST_DATE_SVG = `
  <svg viewBox="0 0 24 24" aria-hidden="true" focusable="false">
    <path d="M7 2a1 1 0 0 1 1 1v1h8V3a1 1 0 1 1 2 0v1h1a3 3 0 0 1 3 3v11a4 4 0 0 1-4 4H6a4 4 0 0 1-4-4V7a3 3 0 0 1 3-3h1V3a1 1 0 0 1 1-1Zm13 9H4v7a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2v-7ZM6 6a1 1 0 0 0-1 1v2h15V7a1 1 0 0 0-1-1H6Zm2 7h3v3H8v-3Z" fill="currentColor"/>
  </svg>
`;
const PROJECT_FILE_DOWNLOAD_SVG = `
  <svg viewBox="0 0 24 24" aria-hidden="true" focusable="false">
    <path d="M12 4v9m0 0 3.5-3.5M12 13l-3.5-3.5M5 18h14" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"/>
  </svg>
`;
const projectCostExchangeRateCache = new Map();

function formatBytes(value) {
  const bytes = Number(value);
  if (!Number.isFinite(bytes) || bytes <= 0) return "0 o";
  if (bytes >= 1024 * 1024) return `${(bytes / (1024 * 1024)).toFixed(1).replace(".", ",")} Mo`;
  if (bytes >= 1024) return `${Math.round(bytes / 1024)} Ko`;
  return `${Math.round(bytes)} o`;
}

function formatDateCell(value) {
  const raw = String(value ?? "").trim();
  if (!raw) return "-";
  if (/^\d{4}-\d{2}-\d{2}/.test(raw)) return raw.slice(0, 10);
  const d = new Date(raw);
  if (Number.isNaN(d.getTime())) return raw;
  return d.toISOString().slice(0, 10);
}

function formatDateYmd(date = new Date()) {
  const d = date instanceof Date ? date : new Date(date);
  if (Number.isNaN(d.getTime())) return "";
  const y = d.getFullYear();
  const m = String(d.getMonth() + 1).padStart(2, "0");
  const day = String(d.getDate()).padStart(2, "0");
  return `${y}-${m}-${day}`;
}

function addDaysToYmd(ymd, days) {
  const raw = String(ymd ?? "").trim();
  if (!/^\d{4}-\d{2}-\d{2}$/.test(raw)) return "";
  const d = new Date(`${raw}T00:00:00`);
  if (Number.isNaN(d.getTime())) return "";
  d.setDate(d.getDate() + Number(days || 0));
  return formatDateYmd(d);
}

function normalizeProjectCostCurrency(value) {
  const code = String(value ?? "").trim().toUpperCase();
  return PROJECT_COST_CURRENCIES.includes(code) ? code : PROJECT_COST_DEFAULT_CURRENCY;
}

function normalizeProjectCostPaypalPurchase(value) {
  return value === true || value === 1 || String(value ?? "").trim().toLowerCase() === "true";
}

function normalizeProjectCostExpenseDate(value) {
  const ymd = formatDateCell(value);
  return /^\d{4}-\d{2}-\d{2}$/.test(ymd) ? ymd : "";
}

function roundExchangeRate(value) {
  const n = Number(value);
  if (!Number.isFinite(n) || n <= 0) return 0;
  return Math.round(n * 1000000) / 1000000;
}

function normalizeProjectCostExchangeRate(value, currency = PROJECT_COST_DEFAULT_CURRENCY) {
  const normalizedCurrency = normalizeProjectCostCurrency(currency);
  if (normalizedCurrency === PROJECT_COST_DEFAULT_CURRENCY) return 1;
  return roundExchangeRate(value);
}

function formatExchangeRateValue(value) {
  const rate = roundExchangeRate(value);
  if (!(rate > 0)) return "";
  return rate.toLocaleString("fr-CA", {
    minimumFractionDigits: 4,
    maximumFractionDigits: 6,
  });
}

function getProjectCostLineBaseExchangeRate(line) {
  const currency = normalizeProjectCostCurrency(line?.currency);
  if (currency === PROJECT_COST_DEFAULT_CURRENCY) return 1;
  return normalizeProjectCostExchangeRate(line?.exchange_rate_to_cad, currency);
}

function getProjectCostLineExchangeRate(line) {
  const currency = normalizeProjectCostCurrency(line?.currency);
  if (currency === PROJECT_COST_DEFAULT_CURRENCY) return 1;
  const baseRate = getProjectCostLineBaseExchangeRate(line);
  if (!(baseRate > 0)) return 0;
  return normalizeProjectCostPaypalPurchase(line?.paypal_purchase)
    ? roundExchangeRate(baseRate * (1 + PROJECT_COST_PAYPAL_SPREAD))
    : baseRate;
}

function getProjectCostLineCadUnitPrice(line) {
  const unitPrice = Math.max(0, Number(line?.unit_price || 0));
  const exchangeRate = getProjectCostLineExchangeRate(line);
  return roundFactureMoney(unitPrice * exchangeRate);
}

function getProjectCostLineCadAmount(line) {
  const quantity = Math.max(0, Number(line?.quantity || 0));
  return roundFactureMoney(quantity * getProjectCostLineCadUnitPrice(line));
}

function clearProjectCostLineExchangeRate(line) {
  if (!line || typeof line !== "object") return line;
  const currency = normalizeProjectCostCurrency(line.currency);
  line.exchange_rate_to_cad = currency === PROJECT_COST_DEFAULT_CURRENCY ? 1 : 0;
  line.amount = getProjectCostLineCadAmount(line);
  return line;
}

function formatProjectCostFxStamp(currency, exchangeRate, expenseDate, { paypalPurchase = false } = {}) {
  const normalizedCurrency = normalizeProjectCostCurrency(currency);
  const normalizedDate = normalizeProjectCostExpenseDate(expenseDate);
  const rate = normalizeProjectCostExchangeRate(exchangeRate, normalizedCurrency);
  if (normalizedCurrency === PROJECT_COST_DEFAULT_CURRENCY || !(rate > 0)) return "";
  const rateText = formatExchangeRateValue(rate);
  const conversionLine = normalizedDate
    ? `1 ${normalizedCurrency} = ${rateText} CAD ${normalizedDate}`
    : `1 ${normalizedCurrency} = ${rateText} CAD`;
  return paypalPurchase
    ? `${conversionLine}\n4% de frais de conversion PayPal`
    : conversionLine;
}

function buildProjectCostPaypalToggleHtml(inputClass, checked, disabled) {
  return `
    <label class="project-cost-paypal-toggle">
      <input type="checkbox" class="${escapeHtml(inputClass)}" ${checked ? "checked" : ""} ${disabled ? "disabled" : ""}>
      <span>Achat PayPal</span>
    </label>
  `;
}

async function fetchHistoricalCadExchangeRate(currency, expenseDate) {
  const normalizedCurrency = normalizeProjectCostCurrency(currency);
  if (normalizedCurrency === PROJECT_COST_DEFAULT_CURRENCY) return 1;
  const normalizedDate = normalizeProjectCostExpenseDate(expenseDate);
  if (!normalizedDate) {
    throw new Error("Date de depense invalide pour le taux de change.");
  }
  const cacheKey = `${normalizedCurrency}:${normalizedDate}`;
  if (projectCostExchangeRateCache.has(cacheKey)) {
    return projectCostExchangeRateCache.get(cacheKey);
  }
  const fetchDirectRate = async () => {
    const url = `${PROJECT_COST_FX_API_BASE}/${encodeURIComponent(normalizedDate)}?base=${encodeURIComponent(normalizedCurrency)}&symbols=CAD`;
    let res;
    try {
      res = await fetch(url, {
        method: "GET",
        headers: { Accept: "application/json" },
      });
    } catch (_err) {
      throw new Error("Impossible de recuperer le taux de change. Verifiez la connexion reseau ou deployez la fonction exchange-rate-get.");
    }
    if (!res.ok) {
      throw new Error(`Impossible de recuperer le taux ${normalizedCurrency}/CAD du ${normalizedDate}.`);
    }
    const json = await res.json().catch(() => ({}));
    const rate = roundExchangeRate(json?.rates?.CAD);
    if (!(rate > 0)) {
      throw new Error(`Taux ${normalizedCurrency}/CAD introuvable pour le ${normalizedDate}.`);
    }
    return rate;
  };
  const client = state.authClient;
  let rate = 0;
  if (client) {
    try {
      const { data, error } = await client.functions.invoke("exchange-rate-get", {
        body: {
          currency: normalizedCurrency,
          expense_date: normalizedDate,
        },
      });
      if (error) {
        throw new Error(resolveEdgeFunctionInvokeError(error, "Erreur exchange-rate-get."));
      }
      rate = roundExchangeRate(data?.rate ?? data?.exchange_rate_to_cad);
      if (!(rate > 0)) {
        throw new Error(`Taux ${normalizedCurrency}/CAD introuvable pour le ${normalizedDate}.`);
      }
    } catch (_err) {
      rate = await fetchDirectRate();
    }
  } else {
    rate = await fetchDirectRate();
  }
  projectCostExchangeRateCache.set(cacheKey, rate);
  return rate;
}

function buildFactureNumeroFromSeqAndDate(seqValue, dateValue) {
  const seq = Number(seqValue);
  if (!Number.isFinite(seq) || seq <= 0) return "";
  if (Math.trunc(seq) < 39) {
    return `F${String(Math.trunc(seq)).padStart(5, "0")}`;
  }
  const seqToken = String(Math.trunc(seq)).padStart(4, "0");

  const rawDate = String(dateValue ?? "").trim();
  let ymd = "";
  if (/^\d{4}-\d{2}-\d{2}/.test(rawDate)) {
    ymd = rawDate.slice(0, 10);
  } else {
    const d = new Date(rawDate);
    if (!Number.isNaN(d.getTime())) ymd = formatDateYmd(d);
  }
  if (!/^\d{4}-\d{2}-\d{2}$/.test(ymd)) {
    ymd = formatDateYmd(new Date());
  }

  const yy = ymd.slice(2, 4);
  const mm = ymd.slice(5, 7);
  const dd = ymd.slice(8, 10);
  return `F${yy}${mm}${dd}${seqToken}`;
}

function normalizeFactureCodeCandidate(value) {
  const text = cleanNullableText(value);
  if (!text) return "";
  const upper = text.toUpperCase();
  return /^F\d{4,}(?:\.\d+)?$/.test(upper) ? upper : "";
}

function sanitizeFilenameToken(value, fallback = "fichier") {
  const text = String(value ?? "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/[^\w.-]+/g, "-")
    .replace(/-+/g, "-")
    .replace(/^-+|-+$/g, "");
  return text || fallback;
}

function normalizeRepairPhotoPhase(value, fallback = "before") {
  const text = normalizeText(value);
  if (text === "after" || text === "apres" || text === "photo apres") return "after";
  if (text === "before" || text === "avant" || text === "photo avant") return "before";
  return fallback === "after" ? "after" : "before";
}

function inferRepairPhotoPhase(row) {
  const direct = cleanNullableText(row?.phase);
  if (direct) return normalizeRepairPhotoPhase(direct, "before");
  const hint = normalizeText(`${row?.original_name || ""} ${row?.filename || ""}`);
  if (/(^|[^a-z])(after|apres)($|[^a-z])/.test(hint)) return "after";
  if (/(^|[^a-z])(before|avant)($|[^a-z])/.test(hint)) return "before";
  return "before";
}

function normalizeRepairPhotos(value) {
  if (!Array.isArray(value)) return [];
  const rows = value
    .map((row) => {
      if (!row || typeof row !== "object") return null;
      const filename = cleanNullableText(row.filename);
      if (!filename) return null;
      const bucket = cleanNullableText(row.bucket) || REPAIRS_PHOTOS_BUCKET;
      const originalName = cleanNullableText(row.original_name);
      const contentType = cleanNullableText(row.content_type);
      const createdAt = cleanNullableText(row.created_at);
      const rawSize = row.size == null ? null : Number(row.size);
      const size = Number.isFinite(rawSize) && rawSize >= 0 ? Math.trunc(rawSize) : null;
      const phase = inferRepairPhotoPhase(row);
      return {
        filename,
        bucket,
        original_name: originalName,
        content_type: contentType,
        size,
        created_at: createdAt,
        phase,
      };
    })
    .filter(Boolean);

  const scoreRow = (row) => {
    let score = 0;
    const bucket = String(row?.bucket || REPAIRS_PHOTOS_BUCKET);
    const filename = String(row?.filename || "");
    const contentType = String(row?.content_type || "").toLowerCase();
    if (bucket === REPAIRS_PHOTOS_BUCKET) score += 2;
    if (/^rep\d{4}_(av|ar)_\d+\.[a-z0-9]+$/i.test(filename)) score += 5;
    if (contentType.startsWith("image/")) score += 1;
    const createdAtMs = Date.parse(String(row?.created_at || "")) || 0;
    return { score, createdAtMs };
  };

  const byKey = new Map();
  for (const row of rows) {
    const phase = normalizeRepairPhotoPhase(row.phase, "before");
    const originalName = normalizeText(row.original_name);
    const filenameBase = normalizeText(String(row.filename || "").split("/").pop() || "");
    const dedupeName = originalName || filenameBase;
    const key = dedupeName
      ? `${phase}|name:${dedupeName}`
      : `${phase}|file:${String(row.bucket || REPAIRS_PHOTOS_BUCKET)}/${String(row.filename || "")}`;
    const prev = byKey.get(key);
    if (!prev) {
      byKey.set(key, row);
      continue;
    }
    const a = scoreRow(prev);
    const b = scoreRow(row);
    if (b.score > a.score || (b.score === a.score && b.createdAtMs >= a.createdAtMs)) {
      byKey.set(key, row);
    }
  }
  return Array.from(byKey.values());
}

function serializeRepairPhotos(value) {
  const rows = normalizeRepairPhotos(value)
    .map((row) => ({
      filename: row.filename,
      bucket: row.bucket || REPAIRS_PHOTOS_BUCKET,
      original_name: row.original_name || null,
      content_type: row.content_type || null,
      size: row.size == null ? null : Number(row.size),
      created_at: row.created_at || null,
      phase: normalizeRepairPhotoPhase(row.phase, "before"),
    }))
    .sort((a, b) => {
      const kA = `${a.bucket}/${a.filename}`;
      const kB = `${b.bucket}/${b.filename}`;
      return kA.localeCompare(kB, "fr-CA");
    });
  return JSON.stringify(rows);
}

function normalizeRepairPhotoThumbnail(value) {
  if (!value || typeof value !== "object") return null;
  const filename = cleanNullableText(value.filename);
  if (!filename) return null;
  const bucket = cleanNullableText(value.bucket) || REPAIRS_PHOTOS_BUCKET;
  const contentType = cleanNullableText(value.content_type) || "image/jpeg";
  const sourceFilename = cleanNullableText(value.source_filename);
  const createdAt = cleanNullableText(value.created_at);
  const rawSize = value.size == null ? null : Number(value.size);
  const size = Number.isFinite(rawSize) && rawSize >= 0 ? Math.trunc(rawSize) : null;
  const phase = normalizeRepairPhotoPhase(value.phase, "before");
  return {
    filename,
    bucket,
    content_type: contentType,
    source_filename: sourceFilename,
    created_at: createdAt,
    size,
    phase,
  };
}

function serializeRepairPhotoThumbnail(value) {
  const row = normalizeRepairPhotoThumbnail(value);
  if (!row) return "null";
  return JSON.stringify({
    filename: row.filename,
    bucket: row.bucket || REPAIRS_PHOTOS_BUCKET,
    content_type: row.content_type || "image/jpeg",
    source_filename: row.source_filename || null,
    created_at: row.created_at || null,
    size: row.size == null ? null : Number(row.size),
    phase: normalizeRepairPhotoPhase(row.phase, "before"),
  });
}

function getRepairStorageNumberToken(repairRow) {
  const podioAppItemId = Number(repairRow?.podio_app_item_id);
  if (Number.isInteger(podioAppItemId) && podioAppItemId > 0) {
    return `rep${String(podioAppItemId).padStart(4, "0")}`;
  }
  const podioItemId = Number(repairRow?.podio_item_id);
  if (Number.isInteger(podioItemId) && podioItemId > 0) {
    return `rep${String(podioItemId).slice(-4).padStart(4, "0")}`;
  }
  return "rep0000";
}

function resolveImageFileExtension(fileName, mimeType, fallback = "jpg") {
  const lowerName = String(fileName ?? "").toLowerCase().trim();
  const extMatch = lowerName.match(/\.([a-z0-9]{2,6})$/);
  if (extMatch?.[1]) {
    const ext = extMatch[1] === "jpeg" ? "jpg" : extMatch[1];
    if (["jpg", "png", "webp", "gif", "bmp", "heic", "heif", "avif"].includes(ext)) return ext;
  }
  const mime = String(mimeType ?? "").toLowerCase();
  if (mime.includes("png")) return "png";
  if (mime.includes("webp")) return "webp";
  if (mime.includes("gif")) return "gif";
  if (mime.includes("bmp")) return "bmp";
  if (mime.includes("heic")) return "heic";
  if (mime.includes("heif")) return "heif";
  if (mime.includes("avif")) return "avif";
  if (mime.includes("jpeg") || mime.includes("jpg")) return "jpg";
  return fallback;
}

function buildRepairPhotoObjectPath(
  repairRow,
  phase = "before",
  seq = 1,
  fileName = "photo.jpg",
  mimeType = "",
) {
  const rep = getRepairStorageNumberToken(repairRow);
  const ph = normalizeRepairPhotoPhase(phase, "before") === "after" ? "ar" : "av";
  const index = String(Math.max(1, Math.trunc(Number(seq) || 1))).padStart(2, "0");
  const ext = resolveImageFileExtension(fileName, mimeType, "jpg");
  return `${rep}_${ph}_${index}.${ext}`;
}

function buildRepairThumbnailObjectPath(repairRow) {
  return `${getRepairStorageNumberToken(repairRow)}_thumb.jpg`;
}

function buildRepairFileObjectPath(repairRow, fileName, seq = 1) {
  const token = sanitizeFilenameToken(getRepairStorageNumberToken(repairRow), "rep0000");
  const stamp = Date.now();
  const index = String(Math.max(1, Math.trunc(Number(seq) || 1))).padStart(2, "0");
  const safeName = sanitizeFilenameToken(fileName || `fichier-${stamp}-${index}`, `fichier-${stamp}-${index}`);
  return `${token}/fichiers/${stamp}_${index}_${safeName}`;
}

async function uploadRepairPhotoFile(client, repairRow, file, phase = "before", seq = 1) {
  if (!client) throw new Error("Connexion Supabase indisponible.");
  if (!(file instanceof File)) throw new Error("Fichier photo invalide.");
  const objectPath = buildRepairPhotoObjectPath(
    repairRow,
    phase,
    seq,
    file.name || "photo.jpg",
    file.type || "",
  );
  const { error } = await client.storage
    .from(REPAIRS_PHOTOS_BUCKET)
    .upload(objectPath, file, {
      upsert: true,
      contentType: file.type || "application/octet-stream",
      cacheControl: "3600",
    });
  if (error) {
    throw new Error(error.message || "Impossible de téléverser la photo.");
  }
  return {
    filename: objectPath,
    bucket: REPAIRS_PHOTOS_BUCKET,
    original_name: cleanNullableText(file.name),
    content_type: cleanNullableText(file.type),
    size: Number.isFinite(Number(file.size)) ? Number(file.size) : null,
    created_at: new Date().toISOString(),
    phase: normalizeRepairPhotoPhase(phase, "before"),
  };
}

async function uploadRepairFileAttachment(client, repairRow, file, seq = 1) {
  if (!client) throw new Error("Connexion Supabase indisponible.");
  if (!(file instanceof File)) throw new Error("Fichier invalide.");
  const objectPath = buildRepairFileObjectPath(repairRow, file.name || "fichier", seq);
  const { error } = await client.storage
    .from(PROJECT_FILES_BUCKET)
    .upload(objectPath, file, {
      upsert: false,
      contentType: file.type || "application/octet-stream",
      cacheControl: "3600",
    });
  if (error) throw new Error(error.message || "Impossible de televerser le fichier.");
  return {
    filename: objectPath,
    bucket: PROJECT_FILES_BUCKET,
    original_name: cleanNullableText(file.name),
    content_type: cleanNullableText(file.type),
    size: Number.isFinite(Number(file.size)) ? Number(file.size) : null,
    created_at: new Date().toISOString(),
  };
}

async function uploadRepairThumbnailBlob(client, repairRow, blob, sourcePhoto = null) {
  if (!client) throw new Error("Connexion Supabase indisponible.");
  if (!(blob instanceof Blob)) throw new Error("Miniature invalide.");
  const objectPath = buildRepairThumbnailObjectPath(repairRow);
  const { error } = await client.storage
    .from(REPAIRS_PHOTOS_BUCKET)
    .upload(objectPath, blob, {
      upsert: true,
      contentType: "image/jpeg",
      cacheControl: "3600",
    });
  if (error) {
    throw new Error(error.message || "Impossible de téléverser la miniature.");
  }
  return {
    filename: objectPath,
    bucket: REPAIRS_PHOTOS_BUCKET,
    content_type: "image/jpeg",
    source_filename: cleanNullableText(sourcePhoto?.filename),
    created_at: new Date().toISOString(),
    size: Number.isFinite(Number(blob.size)) ? Number(blob.size) : null,
    phase: normalizeRepairPhotoPhase(sourcePhoto?.phase, "before"),
  };
}

async function getRepairPhotoUrl(client, photo) {
  if (!client) throw new Error("Connexion Supabase indisponible.");
  const normalized = normalizeRepairPhotos([photo]);
  if (!normalized.length) throw new Error("Référence de photo invalide.");
  const row = normalized[0];
  const bucket = row.bucket || REPAIRS_PHOTOS_BUCKET;
  const objectPath = row.filename;
  if (!bucket || !objectPath) throw new Error("Référence de photo invalide.");

  const signed = await client.storage
    .from(bucket)
    .createSignedUrl(objectPath, 900, { download: false });
  if (!signed?.error && signed?.data?.signedUrl) {
    return signed.data.signedUrl;
  }

  throw new Error(
    signed?.error?.message
      || `Impossible d'ouvrir la photo (${bucket}/${objectPath}). Vérifie les policies Storage du bucket.`,
  );
}

function getRepairThumbnailKey(repair) {
  const thumb = normalizeRepairPhotoThumbnail(repair?.photo_thumbnail);
  if (!thumb) return null;
  return `${thumb.bucket || REPAIRS_PHOTOS_BUCKET}/${thumb.filename}`;
}

function getRepairThumbnailUrlFromCache(repair) {
  const key = getRepairThumbnailKey(repair);
  if (!key) return null;
  return repairThumbUrlCache.get(key) || null;
}

function dropRepairThumbnailCacheEntry(thumbnail) {
  const thumb = normalizeRepairPhotoThumbnail(thumbnail);
  if (!thumb) return;
  const key = `${thumb.bucket || REPAIRS_PHOTOS_BUCKET}/${thumb.filename}`;
  repairThumbUrlCache.delete(key);
  repairThumbLoadingKeys.delete(key);
}

async function hydrateRepairThumbnailUrls(repairs) {
  if (!state.authClient) return false;
  const rows = Array.isArray(repairs) ? repairs : [];
  const todo = rows
    .map((repair) => {
      const thumb = normalizeRepairPhotoThumbnail(repair?.photo_thumbnail);
      if (!thumb) return null;
      const key = `${thumb.bucket || REPAIRS_PHOTOS_BUCKET}/${thumb.filename}`;
      if (!key || repairThumbUrlCache.has(key) || repairThumbLoadingKeys.has(key)) return null;
      return { key, photo: thumb };
    })
    .filter(Boolean);

  if (!todo.length) return false;
  let didUpdate = false;
  await Promise.all(todo.map(async (entry) => {
    const { key, photo } = entry;
    repairThumbLoadingKeys.add(key);
    try {
      const url = await getRepairPhotoUrl(state.authClient, photo);
      if (url) {
        repairThumbUrlCache.set(key, url);
        didUpdate = true;
      }
    } catch {
      // ignore thumbnail hydration failures
    } finally {
      repairThumbLoadingKeys.delete(key);
    }
  }));
  return didUpdate;
}

function isProjectOpeningStageLabel(value) {
  return normalizeText(value) === normalizeText(PROJECT_STAGE_OPENING_LABEL);
}

function isProjectEmitInvoiceStageLabel(value) {
  return normalizeText(value) === normalizeText(PROJECT_STAGE_EMIT_INVOICE_LABEL);
}

function normalizeProjectStages(value) {
  const parsed = Array.isArray(value)
    ? value
      .map((row, index) => {
        if (!row || typeof row !== "object") return null;
        const label = cleanNullableText(row.label ?? row.name ?? row.title);
        if (!label) return null;
        const description = cleanNullableText(row.description ?? row.details ?? row.note ?? row.notes ?? row.desc);
        const done = row.done === true || normalizeText(row.done) === "true" || normalizeText(row.done) === "oui";
        const id = cleanNullableText(row.id) || `stage-${index + 1}`;
        return {
          id,
          label,
          description,
          done: Boolean(done),
        };
      })
      .filter(Boolean)
    : [];

  const opening = parsed.find((row) => isProjectOpeningStageLabel(row?.label)) || null;
  const emitInvoice = parsed.find((row) => isProjectEmitInvoiceStageLabel(row?.label)) || null;
  const middle = parsed.filter((row) => !isProjectOpeningStageLabel(row?.label) && !isProjectEmitInvoiceStageLabel(row?.label));
  const usedIds = new Set();
  const nextUniqueId = (preferred, fallbackBase) => {
    const preferredText = cleanNullableText(preferred);
    if (preferredText && !usedIds.has(preferredText)) {
      usedIds.add(preferredText);
      return preferredText;
    }
    let seq = 1;
    while (usedIds.has(`${fallbackBase}-${seq}`)) seq += 1;
    const id = `${fallbackBase}-${seq}`;
    usedIds.add(id);
    return id;
  };

  const out = [];
  out.push({
    id: nextUniqueId(opening?.id, "stage-opening"),
    label: PROJECT_STAGE_OPENING_LABEL,
    description: cleanNullableText(opening?.description) || "",
    done: Boolean(opening?.done),
  });
  for (const stage of middle) {
    out.push({
      id: nextUniqueId(stage?.id, "stage"),
      label: cleanNullableText(stage?.label) || "Étape",
      description: cleanNullableText(stage?.description) || "",
      done: Boolean(stage?.done),
    });
  }
  out.push({
    id: nextUniqueId(emitInvoice?.id, "stage-emit-invoice"),
    label: PROJECT_STAGE_EMIT_INVOICE_LABEL,
    description: cleanNullableText(emitInvoice?.description) || "",
    done: Boolean(emitInvoice?.done),
  });
  return out;
}

function serializeProjectStages(value) {
  return JSON.stringify(
    normalizeProjectStages(value).map((stage) => ({
      id: stage.id,
      label: stage.label,
      description: cleanNullableText(stage.description) || "",
      done: Boolean(stage.done),
    }))
  );
}

function getProjectProgressSummary(stagesValue) {
  const stages = normalizeProjectStages(stagesValue);
  const total = stages.length;
  const done = stages.filter((stage) => stage.done).length;
  const percent = total > 0 ? Math.round((done / total) * 100) : 0;
  let phaseLabel = "Ouverture";
  if (total > 0) {
    const nextPending = stages.find((stage) => !stage.done);
    phaseLabel = nextPending ? nextPending.label : stages[stages.length - 1].label;
  }
  return {
    stages,
    total,
    done,
    percent,
    phaseLabel,
  };
}

function normalizeProjectCostLines(value) {
  if (!Array.isArray(value)) return [];
  return value
    .map((row) => {
      if (!row || typeof row !== "object") return null;
      const description = cleanNullableText(row.description);
      const qtyRaw = Number(row.quantity ?? row.qty ?? 0);
      const unitRaw = Number(row.unit_price ?? row.unitPrice ?? 0);
      const quantity = Number.isFinite(qtyRaw) && qtyRaw >= 0 ? qtyRaw : 0;
      const unitPrice = Number.isFinite(unitRaw) && unitRaw >= 0 ? unitRaw : 0;
      const currency = normalizeProjectCostCurrency(row.currency);
      const nativeAmount = quantity * unitPrice;
      const storedAmountRaw = Number(row.amount);
      const derivedRateRaw = nativeAmount > 0 ? storedAmountRaw / nativeAmount : 0;
      const exchangeRate = currency === PROJECT_COST_DEFAULT_CURRENCY
        ? 1
        : (normalizeProjectCostExchangeRate(row.exchange_rate_to_cad, currency)
          || normalizeProjectCostExchangeRate(derivedRateRaw, currency));
      const amount = Number.isFinite(storedAmountRaw) && storedAmountRaw >= 0
        ? roundFactureMoney(storedAmountRaw)
        : roundFactureMoney(nativeAmount * (exchangeRate || 0));
      if (!description && amount <= 0 && nativeAmount <= 0) return null;
      const laborRateKey = cleanNullableText(row.labor_rate_key);
      const sourceKind = cleanNullableText(row.source_kind);
      const paypalPurchase = normalizeProjectCostPaypalPurchase(row.paypal_purchase);
      return {
        description: description || "",
        quantity,
        unit_price: unitPrice,
        amount,
        currency,
        exchange_rate_to_cad: exchangeRate,
        expense_date: normalizeProjectCostExpenseDate(row.expense_date),
        labor_rate_key: laborRateKey || null,
        paypal_purchase: paypalPurchase,
        source_kind: sourceKind || null,
      };
    })
    .filter(Boolean);
}

function serializeProjectCostLines(value) {
  return JSON.stringify(
    normalizeProjectCostLines(value).map((line) => ({
      description: line.description,
      quantity: Number(line.quantity || 0),
      unit_price: Number(line.unit_price || 0),
      amount: Number(line.amount || 0),
      currency: normalizeProjectCostCurrency(line.currency),
      exchange_rate_to_cad: getProjectCostLineBaseExchangeRate(line),
      expense_date: normalizeProjectCostExpenseDate(line.expense_date) || null,
      labor_rate_key: cleanNullableText(line.labor_rate_key),
      paypal_purchase: normalizeProjectCostPaypalPurchase(line.paypal_purchase),
      source_kind: cleanNullableText(line.source_kind),
    }))
  );
}

function normalizeRepairInvoiceLines(value) {
  if (!Array.isArray(value)) return [];
  return value
    .map((row) => {
      if (!row || typeof row !== "object") return null;
      const description = cleanNullableText(row.description);
      const qtyRaw = Number(row.quantity ?? row.qty ?? 0);
      const unitRaw = Number(row.unit_price ?? row.unitPrice ?? 0);
      const quantity = Number.isFinite(qtyRaw) && qtyRaw >= 0 ? qtyRaw : 0;
      const unitPrice = Number.isFinite(unitRaw) && unitRaw >= 0 ? unitRaw : 0;
      const currency = normalizeProjectCostCurrency(row.currency);
      const nativeAmount = quantity * unitPrice;
      const storedAmountRaw = Number(row.amount);
      const derivedRateRaw = nativeAmount > 0 ? storedAmountRaw / nativeAmount : 0;
      const exchangeRate = currency === PROJECT_COST_DEFAULT_CURRENCY
        ? 1
        : (normalizeProjectCostExchangeRate(row.exchange_rate_to_cad, currency)
          || normalizeProjectCostExchangeRate(derivedRateRaw, currency));
      const amount = Number.isFinite(storedAmountRaw) && storedAmountRaw >= 0
        ? roundFactureMoney(storedAmountRaw)
        : roundFactureMoney(nativeAmount * (exchangeRate || 0));
      const laborRateKey = cleanNullableText(row.labor_rate_key);
      const sourceKind = cleanNullableText(row.source_kind);
      const paypalPurchase = normalizeProjectCostPaypalPurchase(row.paypal_purchase);
      if (!description && amount <= 0 && nativeAmount <= 0 && sourceKind !== "repair_labor") return null;
      return {
        description: description || "",
        quantity,
        unit_price: unitPrice,
        amount,
        currency,
        exchange_rate_to_cad: exchangeRate,
        expense_date: normalizeProjectCostExpenseDate(row.expense_date),
        labor_rate_key: laborRateKey || null,
        paypal_purchase: paypalPurchase,
        source_kind: sourceKind || null,
      };
    })
    .filter(Boolean);
}

function serializeRepairInvoiceLines(value) {
  return JSON.stringify(
    normalizeRepairInvoiceLines(value).map((line) => ({
      description: line.description,
      quantity: Number(line.quantity || 0),
      unit_price: Number(line.unit_price || 0),
      amount: Number(line.amount || 0),
      currency: normalizeProjectCostCurrency(line.currency),
      exchange_rate_to_cad: getProjectCostLineBaseExchangeRate(line),
      expense_date: normalizeProjectCostExpenseDate(line.expense_date) || null,
      labor_rate_key: cleanNullableText(line.labor_rate_key),
      paypal_purchase: normalizeProjectCostPaypalPurchase(line.paypal_purchase),
      source_kind: cleanNullableText(line.source_kind),
    }))
  );
}

function isRepairLaborDescription(value) {
  const normalized = normalizeText(value)
    .replace(/[’'`]/g, "")
    .replace(/\s+/g, " ")
    .trim();
  return normalized.includes("main doeuvre") || normalized.includes("main d oeuvre");
}

function isRepairLaborLine(line) {
  if (!line || typeof line !== "object") return false;
  if (String(line.source_kind ?? "").trim() === "repair_labor") return true;
  return isRepairLaborDescription(line.description);
}

const DEFAULT_REPAIR_LABOR_RATE_KEY = "labor_rate_standard";

function ensureRepairInvoiceLines(
  linesValue,
  {
    durationMinutes = null,
    partsCost = null,
    defaultLaborRate = null,
    hourlyRateSnapshot = null,
    fallbackLaborQuantity = 0,
    includeLabor = true,
  } = {},
) {
  const lines = normalizeRepairInvoiceLines(linesValue);
  const fallbackRate = Number.isFinite(Number(defaultLaborRate)) && Number(defaultLaborRate) >= 0
    ? roundFactureMoney(defaultLaborRate)
    : roundFactureMoney(getConfiguredFactureLaborRate());
  const snapshotRate = Number.isFinite(Number(hourlyRateSnapshot)) && Number(hourlyRateSnapshot) >= 0
    ? roundFactureMoney(hourlyRateSnapshot)
    : null;
  const legacyHoursRaw = Number(durationMinutes);
  const legacyHours = Number.isFinite(legacyHoursRaw) && legacyHoursRaw > 0
    ? roundFactureMoney(legacyHoursRaw / 60)
    : null;
  const defaultQuantityRaw = Number(fallbackLaborQuantity);
  const defaultQuantity = Number.isFinite(defaultQuantityRaw) && defaultQuantityRaw >= 0
    ? roundFactureMoney(defaultQuantityRaw)
    : 0;

  if (!includeLabor) {
    const personalLines = lines.filter((line) => !isRepairLaborLine(line));
    const normalizedPartsCost = Number(partsCost);
    const hasPositivePartsLine = personalLines.some((line) => roundFactureMoney(line.amount) > 0);
    if (!hasPositivePartsLine && Number.isFinite(normalizedPartsCost) && normalizedPartsCost > 0) {
      personalLines.push({
        description: "Pièces",
        quantity: 1,
        unit_price: roundFactureMoney(normalizedPartsCost),
        amount: roundFactureMoney(normalizedPartsCost),
        currency: PROJECT_COST_DEFAULT_CURRENCY,
        exchange_rate_to_cad: 1,
        expense_date: null,
        labor_rate_key: null,
        paypal_purchase: false,
        source_kind: "repair_parts",
      });
    }
    return personalLines.map((line) => {
      const normalizedLine = {
        ...line,
        currency: normalizeProjectCostCurrency(line.currency),
        exchange_rate_to_cad: getProjectCostLineBaseExchangeRate(line),
        expense_date: normalizeProjectCostExpenseDate(line.expense_date),
        paypal_purchase: normalizeProjectCostPaypalPurchase(line.paypal_purchase),
      };
      return {
        ...normalizedLine,
        amount: getProjectCostLineCadAmount(normalizedLine),
      };
    });
  }

  let laborIndex = lines.findIndex((line) => isRepairLaborLine(line));
  const laborQuantity = legacyHours != null ? legacyHours : defaultQuantity;
  if (laborIndex < 0) {
    lines.unshift({
      description: "Main d'oeuvre",
      quantity: laborQuantity,
      unit_price: fallbackRate,
      amount: roundFactureMoney(laborQuantity * fallbackRate),
      currency: PROJECT_COST_DEFAULT_CURRENCY,
      exchange_rate_to_cad: 1,
      expense_date: null,
      labor_rate_key: DEFAULT_REPAIR_LABOR_RATE_KEY,
      paypal_purchase: false,
      source_kind: "repair_labor",
    });
  } else {
    const laborLine = { ...lines[laborIndex] };
    const qtyRaw = Number(laborLine.quantity);
    const unitRaw = Number(laborLine.unit_price);
    const laborRateKey = cleanNullableText(laborLine.labor_rate_key);
    const quantity = Number.isFinite(qtyRaw) && qtyRaw >= 0
      ? roundFactureMoney(qtyRaw)
      : laborQuantity;
    const normalizedUnitRaw = Number.isFinite(unitRaw) && unitRaw >= 0
      ? roundFactureMoney(unitRaw)
      : null;
    const matchesLegacySnapshot = normalizedUnitRaw != null
      && snapshotRate != null
      && Math.abs(normalizedUnitRaw - snapshotRate) < 0.005;
    const useConfiguredDefaultRate = laborRateKey === DEFAULT_REPAIR_LABOR_RATE_KEY
      || (!laborRateKey && (matchesLegacySnapshot || normalizedUnitRaw == null));
    const unitPrice = useConfiguredDefaultRate
      ? fallbackRate
      : (normalizedUnitRaw == null ? fallbackRate : normalizedUnitRaw);
    laborLine.description = "Main d'oeuvre";
    laborLine.quantity = quantity;
    laborLine.unit_price = unitPrice;
    laborLine.amount = roundFactureMoney(quantity * unitPrice);
    laborLine.currency = PROJECT_COST_DEFAULT_CURRENCY;
    laborLine.exchange_rate_to_cad = 1;
    laborLine.expense_date = null;
    laborLine.paypal_purchase = false;
    laborLine.labor_rate_key = useConfiguredDefaultRate
      ? DEFAULT_REPAIR_LABOR_RATE_KEY
      : (laborRateKey || null);
    laborLine.source_kind = "repair_labor";
    lines[laborIndex] = laborLine;
    if (laborIndex > 0) {
      lines.splice(laborIndex, 1);
      lines.unshift(laborLine);
    }
  }

  const normalizedPartsCost = Number(partsCost);
  const hasPositivePartsLine = lines.some((line, index) => index > 0 && roundFactureMoney(line.amount) > 0);
  if (!hasPositivePartsLine && Number.isFinite(normalizedPartsCost) && normalizedPartsCost > 0) {
    lines.push({
      description: "Pièces",
      quantity: 1,
      unit_price: roundFactureMoney(normalizedPartsCost),
      amount: roundFactureMoney(normalizedPartsCost),
      currency: PROJECT_COST_DEFAULT_CURRENCY,
      exchange_rate_to_cad: 1,
      expense_date: null,
      labor_rate_key: null,
      paypal_purchase: false,
      source_kind: "repair_parts",
    });
  }

  return lines.map((line) => {
    const normalizedLine = {
      ...line,
      currency: normalizeProjectCostCurrency(line.currency),
      exchange_rate_to_cad: getProjectCostLineBaseExchangeRate(line),
      expense_date: normalizeProjectCostExpenseDate(line.expense_date),
      paypal_purchase: normalizeProjectCostPaypalPurchase(line.paypal_purchase),
    };
    if (isRepairLaborLine(normalizedLine)) {
      normalizedLine.currency = PROJECT_COST_DEFAULT_CURRENCY;
      normalizedLine.exchange_rate_to_cad = 1;
      normalizedLine.expense_date = null;
      normalizedLine.paypal_purchase = false;
    }
    return {
      ...normalizedLine,
      amount: getProjectCostLineCadAmount(normalizedLine),
    };
  });
}

function deriveRepairLegacyValuesFromLines(linesValue) {
  const lines = normalizeRepairInvoiceLines(linesValue);
  let laborHours = 0;
  let hasLaborLine = false;
  let partsSubtotal = 0;
  let hasPartsLine = false;

  for (const line of lines) {
    const quantity = roundFactureMoney(Number(line.quantity || 0));
    const amount = Number.isFinite(Number(line.amount))
      ? roundFactureMoney(Number(line.amount))
      : getProjectCostLineCadAmount(line);
    if (isRepairLaborLine(line)) {
      hasLaborLine = true;
      laborHours += quantity;
      continue;
    }
    hasPartsLine = true;
    partsSubtotal += amount;
  }

  return {
    duration_minutes: hasLaborLine ? Math.round(laborHours * 60) : null,
    parts_cost: hasPartsLine ? roundFactureMoney(partsSubtotal) : null,
  };
}

function normalizeProjectPhotos(value) {
  if (!Array.isArray(value)) return [];
  const rows = value
    .map((row) => {
      if (!row || typeof row !== "object") return null;
      const filename = cleanNullableText(row.filename);
      if (!filename) return null;
      const bucket = cleanNullableText(row.bucket) || REPAIRS_PHOTOS_BUCKET;
      return {
        filename,
        bucket,
        original_name: cleanNullableText(row.original_name),
        content_type: cleanNullableText(row.content_type),
        size: Number.isFinite(Number(row.size)) ? Number(row.size) : null,
        created_at: cleanNullableText(row.created_at),
        preview_url: cleanNullableText(row.preview_url),
      };
    })
    .filter(Boolean);

  const byKey = new Map();
  for (const row of rows) {
    const key = `${row.bucket || REPAIRS_PHOTOS_BUCKET}/${row.filename}`;
    if (!byKey.has(key)) {
      byKey.set(key, row);
    }
  }
  return Array.from(byKey.values());
}

function serializeProjectPhotos(value) {
  return JSON.stringify(
    normalizeProjectPhotos(value).map((row) => ({
      filename: row.filename,
      bucket: row.bucket || REPAIRS_PHOTOS_BUCKET,
      original_name: row.original_name || null,
      content_type: row.content_type || null,
      size: row.size == null ? null : Number(row.size),
      created_at: row.created_at || null,
    }))
  );
}

function normalizeProjectPhotoThumbnail(value) {
  if (!value || typeof value !== "object") return null;
  const filename = cleanNullableText(value.filename);
  if (!filename) return null;
  return {
    filename,
    bucket: cleanNullableText(value.bucket) || REPAIRS_PHOTOS_BUCKET,
    original_name: cleanNullableText(value.original_name),
    content_type: cleanNullableText(value.content_type),
    size: Number.isFinite(Number(value.size)) ? Number(value.size) : null,
    created_at: cleanNullableText(value.created_at),
    preview_url: cleanNullableText(value.preview_url),
  };
}

function serializeProjectPhotoThumbnail(value) {
  const row = normalizeProjectPhotoThumbnail(value);
  if (!row) return "null";
  return JSON.stringify({
    filename: row.filename,
    bucket: row.bucket || REPAIRS_PHOTOS_BUCKET,
    original_name: row.original_name || null,
    content_type: row.content_type || null,
    size: row.size == null ? null : Number(row.size),
    created_at: row.created_at || null,
  });
}

function normalizeProjectFiles(value) {
  if (!Array.isArray(value)) return [];
  const rows = value
    .map((row) => {
      if (!row || typeof row !== "object") return null;
      const filename = cleanNullableText(row.filename);
      if (!filename) return null;
      const bucket = cleanNullableText(row.bucket) || PROJECT_FILES_BUCKET;
      return {
        filename,
        bucket,
        original_name: cleanNullableText(row.original_name),
        content_type: cleanNullableText(row.content_type),
        size: Number.isFinite(Number(row.size)) ? Number(row.size) : null,
        created_at: cleanNullableText(row.created_at),
        preview_url: cleanNullableText(row.preview_url),
      };
    })
    .filter(Boolean);
  const byKey = new Map();
  for (const row of rows) {
    const key = `${row.bucket || PROJECT_FILES_BUCKET}/${row.filename}`;
    if (!byKey.has(key)) byKey.set(key, row);
  }
  return Array.from(byKey.values());
}

function serializeProjectFiles(value) {
  return JSON.stringify(
    normalizeProjectFiles(value).map((row) => ({
      filename: row.filename,
      bucket: row.bucket || PROJECT_FILES_BUCKET,
      original_name: row.original_name || null,
      content_type: row.content_type || null,
      size: row.size == null ? null : Number(row.size),
      created_at: row.created_at || null,
    }))
  );
}

function getBonsCardAspectRatio() {
  const sample = document.querySelector(".bons-repair-btn");
  if (!sample) return 2.45;
  const box = sample.getBoundingClientRect();
  if (!box.width || !box.height) return 2.45;
  const ratio = box.width / box.height;
  if (!Number.isFinite(ratio) || ratio < 1.2 || ratio > 4) return 2.45;
  return ratio;
}

function formatProjectCode(project) {
  const numero = cleanNullableText(project?.numero);
  if (numero) return numero.toUpperCase();
  const fallback = Number(project?.id);
  if (Number.isInteger(fallback) && fallback > 0) {
    return `P${String(Math.trunc(fallback)).padStart(4, "0")}`;
  }
  return "P0000";
}

function parseProjectCodeSortValue(code) {
  const text = String(code ?? "").trim().toUpperCase();
  const match = text.match(/^P(\d{1,12})$/);
  if (!match?.[1]) return 0;
  return Number(match[1]) || 0;
}

function nextProjectNumeroFromState() {
  const maxSeq = (state.projects || []).reduce((acc, project) => {
    const value = parseProjectCodeSortValue(formatProjectCode(project));
    return value > acc ? value : acc;
  }, 0);
  return `P${String(maxSeq + 1).padStart(4, "0")}`;
}

function matchProjectSearch(project, query) {
  const q = normalizeText(query);
  if (!q) return true;
  const progress = getProjectProgressSummary(project?.stages);
  const haystack = normalizeText([
    formatProjectCode(project),
    project?.client_title,
    project?.title,
    project?.description,
    progress.phaseLabel,
    `${progress.percent}%`,
  ].join(" "));
  return haystack.includes(q);
}

function isProjectArchived(project) {
  const progress = getProjectProgressSummary(project?.stages);
  return progress.total > 0 && progress.done >= progress.total;
}

async function getStorageSignedObjectUrl(client, bucketValue, objectPathValue) {
  const bucket = cleanNullableText(bucketValue);
  const objectPath = cleanNullableText(objectPathValue);
  if (!bucket || !objectPath) throw new Error("Reference de fichier invalide.");

  const signed = await client.storage
    .from(bucket)
    .createSignedUrl(objectPath, 900, { download: false });
  if (!signed?.error && signed?.data?.signedUrl) return signed.data.signedUrl;

  const publicUrl = cleanNullableText(client.storage.from(bucket).getPublicUrl(objectPath)?.data?.publicUrl);
  if (publicUrl) return publicUrl;

  throw new Error(signed?.error?.message || `Impossible d'ouvrir le fichier (${bucket}/${objectPath}).`);
}

async function getProjectPhotoUrl(client, photo) {
  const normalized = normalizeProjectPhotos([photo]);
  if (!normalized.length) throw new Error("Reference de photo invalide.");
  const row = normalized[0];
  return getStorageSignedObjectUrl(client, row.bucket || REPAIRS_PHOTOS_BUCKET, row.filename);
}

async function getProjectFileUrl(client, fileRow) {
  const normalized = normalizeProjectFiles([fileRow]);
  if (!normalized.length) throw new Error("Reference de fichier invalide.");
  const row = normalized[0];
  if (row.preview_url && String(row.preview_url).startsWith("blob:")) return row.preview_url;
  return getStorageSignedObjectUrl(client, row.bucket || PROJECT_FILES_BUCKET, row.filename);
}

async function getProjectFileDownloadUrl(client, fileRow) {
  const normalized = normalizeProjectFiles([fileRow]);
  if (!normalized.length) throw new Error("Reference de fichier invalide.");
  const row = normalized[0];
  if (row.preview_url && String(row.preview_url).startsWith("blob:")) return row.preview_url;
  const bucket = row.bucket || PROJECT_FILES_BUCKET;
  const fileName = cleanNullableText(row.original_name) || row.filename.split("/").pop() || "fichier";
  const signed = await client.storage
    .from(bucket)
    .createSignedUrl(row.filename, 900, { download: fileName });
  if (!signed?.error && signed?.data?.signedUrl) return signed.data.signedUrl;
  const publicUrl = cleanNullableText(client.storage.from(bucket).getPublicUrl(row.filename)?.data?.publicUrl);
  if (publicUrl) return publicUrl;
  throw new Error(signed?.error?.message || `Impossible de télécharger le fichier (${bucket}/${row.filename}).`);
}

function getProjectThumbnailKey(project) {
  const thumb = normalizeProjectPhotoThumbnail(project?.photo_thumbnail);
  if (!thumb) return null;
  return `${thumb.bucket || REPAIRS_PHOTOS_BUCKET}/${thumb.filename}`;
}

function getProjectThumbnailUrlFromCache(project) {
  const thumb = normalizeProjectPhotoThumbnail(project?.photo_thumbnail);
  if (thumb?.preview_url) return thumb.preview_url;
  const key = getProjectThumbnailKey(project);
  if (!key) return null;
  return projectThumbUrlCache.get(key) || null;
}

function dropProjectThumbnailCacheEntry(thumbnail) {
  const thumb = normalizeProjectPhotoThumbnail(thumbnail);
  if (!thumb) return;
  if (thumb.preview_url && thumb.preview_url.startsWith("blob:")) {
    try {
      URL.revokeObjectURL(thumb.preview_url);
    } catch {
      // ignore
    }
  }
  const key = `${thumb.bucket || REPAIRS_PHOTOS_BUCKET}/${thumb.filename}`;
  projectThumbUrlCache.delete(key);
  projectThumbLoadingKeys.delete(key);
}

async function hydrateProjectThumbnailUrls(projects) {
  if (!state.authClient) return false;
  const rows = Array.isArray(projects) ? projects : [];
  const todo = rows
    .map((project) => {
      const thumb = normalizeProjectPhotoThumbnail(project?.photo_thumbnail);
      if (!thumb || thumb.preview_url) return null;
      const key = `${thumb.bucket || REPAIRS_PHOTOS_BUCKET}/${thumb.filename}`;
      if (!key || projectThumbUrlCache.has(key) || projectThumbLoadingKeys.has(key)) return null;
      return { key, photo: thumb };
    })
    .filter(Boolean);

  if (!todo.length) return false;
  let didUpdate = false;
  await Promise.all(todo.map(async (entry) => {
    const { key, photo } = entry;
    projectThumbLoadingKeys.add(key);
    try {
      const url = await getProjectPhotoUrl(state.authClient, photo);
      if (url) {
        projectThumbUrlCache.set(key, url);
        didUpdate = true;
      }
    } catch {
      // ignore thumbnail hydration failures
    } finally {
      projectThumbLoadingKeys.delete(key);
    }
  }));
  return didUpdate;
}

function buildProjectPhotoObjectPath(projectCode, seq, fileName, mimeType) {
  const token = sanitizeFilenameToken(String(projectCode || "p0000").toLowerCase(), "p0000");
  const stamp = Date.now();
  const index = String(Math.max(1, Math.trunc(Number(seq) || 1))).padStart(2, "0");
  const ext = resolveImageFileExtension(fileName, mimeType, "jpg");
  return `${token}_ph_${stamp}_${index}.${ext}`;
}

function buildProjectThumbnailObjectPath(projectCode) {
  const token = sanitizeFilenameToken(String(projectCode || "p0000").toLowerCase(), "p0000");
  return `${token}_thumb.jpg`;
}

function buildProjectFileObjectPath(projectCode, fileName, seq = 1) {
  const token = sanitizeFilenameToken(String(projectCode || "p0000").toLowerCase(), "p0000");
  const stamp = Date.now();
  const index = String(Math.max(1, Math.trunc(Number(seq) || 1))).padStart(2, "0");
  const safeName = sanitizeFilenameToken(fileName || `fichier-${stamp}-${index}`, `fichier-${stamp}-${index}`);
  return `${token}/fichiers/${stamp}_${index}_${safeName}`;
}

async function uploadProjectPhotoFile(client, projectCode, file, seq = 1) {
  if (!client) throw new Error("Connexion Supabase indisponible.");
  if (!(file instanceof File)) throw new Error("Fichier photo invalide.");
  const objectPath = buildProjectPhotoObjectPath(projectCode, seq, file.name || "photo.jpg", file.type || "");
  const { error } = await client.storage
    .from(REPAIRS_PHOTOS_BUCKET)
    .upload(objectPath, file, {
      upsert: false,
      contentType: file.type || "application/octet-stream",
      cacheControl: "3600",
    });
  if (error) throw new Error(error.message || "Impossible de televerser la photo.");
  return {
    filename: objectPath,
    bucket: REPAIRS_PHOTOS_BUCKET,
    original_name: cleanNullableText(file.name),
    content_type: cleanNullableText(file.type),
    size: Number.isFinite(Number(file.size)) ? Number(file.size) : null,
    created_at: new Date().toISOString(),
  };
}

async function uploadProjectFileAttachment(client, projectCode, file, seq = 1) {
  if (!client) throw new Error("Connexion Supabase indisponible.");
  if (!(file instanceof File)) throw new Error("Fichier invalide.");
  const objectPath = buildProjectFileObjectPath(projectCode, file.name || "fichier", seq);
  const { error } = await client.storage
    .from(PROJECT_FILES_BUCKET)
    .upload(objectPath, file, {
      upsert: false,
      contentType: file.type || "application/octet-stream",
      cacheControl: "3600",
    });
  if (error) throw new Error(error.message || "Impossible de televerser le fichier.");
  return {
    filename: objectPath,
    bucket: PROJECT_FILES_BUCKET,
    original_name: cleanNullableText(file.name),
    content_type: cleanNullableText(file.type),
    size: Number.isFinite(Number(file.size)) ? Number(file.size) : null,
    created_at: new Date().toISOString(),
  };
}

async function uploadProjectThumbnailBlob(client, projectCode, blob, sourcePhoto = null) {
  if (!client) throw new Error("Connexion Supabase indisponible.");
  if (!(blob instanceof Blob)) throw new Error("Miniature invalide.");
  const objectPath = buildProjectThumbnailObjectPath(projectCode);
  const { error } = await client.storage
    .from(REPAIRS_PHOTOS_BUCKET)
    .upload(objectPath, blob, {
      upsert: true,
      contentType: "image/jpeg",
      cacheControl: "3600",
    });
  if (error) throw new Error(error.message || "Impossible de televerser la miniature.");
  return {
    filename: objectPath,
    bucket: REPAIRS_PHOTOS_BUCKET,
    original_name: cleanNullableText(sourcePhoto?.original_name) || "Miniature",
    content_type: "image/jpeg",
    source_filename: cleanNullableText(sourcePhoto?.filename),
    size: Number.isFinite(Number(blob.size)) ? Number(blob.size) : null,
    created_at: new Date().toISOString(),
  };
}

function normalizeNonNegativeNumber(value, fallback) {
  const n = Number(value);
  if (!Number.isFinite(n) || n < 0) return Number(fallback) || 0;
  return Math.round(n * 100) / 100;
}

function normalizeTaxRateNumber(value, fallback) {
  const fallbackNumber = Number(fallback);
  let n = Number(value);
  if (!Number.isFinite(n) || n < 0) {
    n = Number.isFinite(fallbackNumber) && fallbackNumber >= 0 ? fallbackNumber : 0;
  }
  if (n > 1) n /= 100;
  if (Number.isFinite(fallbackNumber) && fallbackNumber > 0) {
    const roundedFallback = Math.round(fallbackNumber * 100) / 100;
    if (Math.abs(n - roundedFallback) < 0.000001 && Math.abs(roundedFallback - fallbackNumber) > 0.000001) {
      n = fallbackNumber;
    }
  }
  return Math.round(n * 1000000) / 1000000;
}

function normalizePurchaseLinkKey(value, fallback = "purchase_link") {
  const raw = String(value ?? "")
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, "_")
    .replace(/^_+|_+$/g, "");
  return raw || fallback;
}

function extractDomainFromUrl(url, fallback = "") {
  const normalized = sanitizeExternalUrl(url);
  if (!normalized) return String(fallback || "").trim().toLowerCase();
  try {
    const parsed = new URL(normalized);
    return String(parsed.hostname || "")
      .replace(/^www\./i, "")
      .trim()
      .toLowerCase();
  } catch (_err) {
    return String(fallback || "").trim().toLowerCase();
  }
}

function normalizePurchaseLinkEntry(value, fallbackPreset = null, index = 0) {
  const src = value && typeof value === "object" ? value : {};
  const fallback = fallbackPreset || {};
  const key = normalizePurchaseLinkKey(
    src.key || fallback.key || src.title || src.name || `purchase_link_${index + 1}`,
    `purchase_link_${index + 1}`
  );
  const title = cleanNullableText(src.title || src.name || fallback.title) || `Lien ${index + 1}`;
  const url = sanitizeExternalUrl(src.url || src.href || src.link || fallback.defaultUrl || fallback.url) || "";
  const note = cleanNullableText(src.note) || "";
  const domain = extractDomainFromUrl(url, src.domain || fallback.domain || "");
  return {
    key,
    title,
    url,
    note,
    domain,
  };
}

function normalizePurchaseLinksCollection(value, legacyRaw = null) {
  const hasArrayInput = Array.isArray(value);
  const source = hasArrayInput ? value : [];
  const normalizedFromArray = source
    .map((row, index) => normalizePurchaseLinkEntry(row, null, index))
    .filter((row) => cleanNullableText(row.title) || cleanNullableText(row.url));

  if (hasArrayInput) {
    const unique = [];
    const usedKeys = new Set();
    normalizedFromArray.forEach((row, index) => {
      let key = normalizePurchaseLinkKey(row.key, `purchase_link_${index + 1}`);
      if (usedKeys.has(key)) {
        let seq = 2;
        while (usedKeys.has(`${key}_${seq}`)) seq += 1;
        key = `${key}_${seq}`;
      }
      usedKeys.add(key);
      unique.push({ ...row, key });
    });
    return unique;
  }

  const raw = legacyRaw && typeof legacyRaw === "object" ? legacyRaw : {};
  return PURCHASE_LINK_PRESETS.map((preset, index) => {
    const legacyUrl = cleanNullableText(raw?.[preset.configKey]);
    return normalizePurchaseLinkEntry(
      {
        key: preset.key,
        title: preset.title,
        url: legacyUrl || preset.defaultUrl,
        note: "",
        domain: preset.domain,
      },
      preset,
      index
    );
  });
}

function clonePurchaseLinks(links) {
  return normalizePurchaseLinksCollection(links).map((row) => ({ ...row }));
}

function openImagePreviewGallery(items, startIndex = 0, fallbackTitle = "Image") {
  const rows = Array.isArray(items) ? items.filter(Boolean) : [];
  if (!rows.length) return;
  let currentIndex = Math.max(0, Math.min(rows.length - 1, Number(startIndex) || 0));
  let isClosed = false;
  let renderToken = 0;

  const root = document.createElement("div");
  root.className = "repair-image-preview-overlay";
  root.innerHTML = `
    <div class="repair-image-preview-shell">
      <button type="button" class="repair-image-preview-nav repair-image-preview-nav-prev" data-preview-nav="prev" aria-label="Photo précédente">‹</button>
      <div class="repair-image-preview-panel" role="dialog" aria-modal="true" aria-label="${escapeHtml(fallbackTitle)}">
        <div class="repair-image-preview-head">
          <div class="repair-image-preview-head-main">
            <p class="repair-image-preview-title"></p>
            <p class="repair-image-preview-counter"></p>
          </div>
          <button type="button" class="repair-image-preview-close" aria-label="Fermer">×</button>
        </div>
        <div class="repair-image-preview-body">
          <div class="repair-image-preview-stage">
            <img class="repair-image-preview-img" alt="">
            <p class="repair-image-preview-status" hidden></p>
          </div>
        </div>
      </div>
      <button type="button" class="repair-image-preview-nav repair-image-preview-nav-next" data-preview-nav="next" aria-label="Photo suivante">›</button>
    </div>
  `;

  const titleEl = root.querySelector(".repair-image-preview-title");
  const counterEl = root.querySelector(".repair-image-preview-counter");
  const imgEl = root.querySelector(".repair-image-preview-img");
  const statusEl = root.querySelector(".repair-image-preview-status");
  const prevBtn = root.querySelector('[data-preview-nav="prev"]');
  const nextBtn = root.querySelector('[data-preview-nav="next"]');

  const close = () => {
    if (isClosed) return;
    isClosed = true;
    document.removeEventListener("keydown", onKeyDown, true);
    root.remove();
  };

  const updateNav = () => {
    const hasMany = rows.length > 1;
    if (prevBtn instanceof HTMLButtonElement) {
      prevBtn.hidden = !hasMany;
      prevBtn.disabled = !hasMany || currentIndex <= 0;
    }
    if (nextBtn instanceof HTMLButtonElement) {
      nextBtn.hidden = !hasMany;
      nextBtn.disabled = !hasMany || currentIndex >= (rows.length - 1);
    }
    if (counterEl instanceof HTMLElement) {
      counterEl.hidden = !hasMany;
      counterEl.textContent = hasMany ? `${currentIndex + 1} / ${rows.length}` : "";
    }
  };

  const renderCurrent = async () => {
    const item = rows[currentIndex];
    if (!item || !(imgEl instanceof HTMLImageElement)) return;
    const nextTitle = cleanNullableText(item.title) || fallbackTitle;
    if (titleEl instanceof HTMLElement) titleEl.textContent = nextTitle;
    updateNav();
    if (statusEl instanceof HTMLElement) {
      statusEl.hidden = false;
      statusEl.textContent = "Chargement...";
    }
    imgEl.hidden = true;
    imgEl.removeAttribute("src");
    imgEl.alt = nextTitle;
    const token = ++renderToken;
    try {
      const nextUrl = typeof item.getUrl === "function"
        ? await item.getUrl()
        : cleanNullableText(item.url);
      if (isClosed || token !== renderToken) return;
      const safeUrl = cleanNullableText(nextUrl);
      if (!safeUrl) throw new Error("Image introuvable.");
      imgEl.src = safeUrl;
      imgEl.alt = nextTitle;
      imgEl.hidden = false;
      if (statusEl instanceof HTMLElement) {
        statusEl.hidden = true;
        statusEl.textContent = "";
      }
    } catch (err) {
      if (isClosed || token !== renderToken) return;
      if (statusEl instanceof HTMLElement) {
        statusEl.hidden = false;
        statusEl.textContent = err?.message || "Impossible d'ouvrir cette image.";
      }
    }
  };

  const move = (delta) => {
    const nextIndex = currentIndex + delta;
    if (nextIndex < 0 || nextIndex >= rows.length) return;
    currentIndex = nextIndex;
    renderCurrent();
  };

  const onKeyDown = (ev) => {
    if (ev.key === "Escape") {
      ev.preventDefault();
      ev.stopPropagation();
      close();
      return;
    }
    if (ev.key === "ArrowLeft") {
      ev.preventDefault();
      ev.stopPropagation();
      move(-1);
      return;
    }
    if (ev.key === "ArrowRight") {
      ev.preventDefault();
      ev.stopPropagation();
      move(1);
    }
  };

  root.addEventListener("click", (ev) => {
    if (ev.target === root || ev.target?.closest?.(".repair-image-preview-close")) {
      close();
      return;
    }
    const navBtn = ev.target?.closest?.("[data-preview-nav]");
    if (!navBtn) return;
    const dir = String(navBtn.getAttribute("data-preview-nav") || "");
    if (dir === "prev") move(-1);
    if (dir === "next") move(1);
  });

  document.addEventListener("keydown", onKeyDown, true);
  document.body.appendChild(root);
  renderCurrent();
}

function normalizeAppConfigRow(row) {
  const raw = row || {};
  const purchaseLinks = normalizePurchaseLinksCollection(raw.purchase_links, raw);
  const purchaseLinksByKey = new Map(purchaseLinks.map((link) => [normalizePurchaseLinkKey(link.key), link]));
  return {
    id: 1,
    labor_rate_standard: normalizeNonNegativeNumber(raw.labor_rate_standard, DEFAULT_APP_CONFIG.labor_rate_standard),
    labor_rate_secondary: normalizeNonNegativeNumber(raw.labor_rate_secondary, DEFAULT_APP_CONFIG.labor_rate_secondary),
    labor_rate_diagnostic: normalizeNonNegativeNumber(raw.labor_rate_diagnostic, DEFAULT_APP_CONFIG.labor_rate_diagnostic),
    labor_rate_estimation: normalizeNonNegativeNumber(raw.labor_rate_estimation, DEFAULT_APP_CONFIG.labor_rate_estimation),
    developer_mode: Boolean(raw.developer_mode),
    podio_sync_enabled: raw.podio_sync_enabled !== false,
    company_name: cleanNullableText(raw.company_name) || "",
    company_phone: cleanNullableText(raw.company_phone) || "",
    company_email: cleanNullableText(raw.company_email) || "",
    company_address_line1: cleanNullableText(raw.company_address_line1) || "",
    company_address_line2: cleanNullableText(raw.company_address_line2) || "",
    company_city: cleanNullableText(raw.company_city) || "",
    company_province: cleanNullableText(raw.company_province) || "",
    company_postal_code: cleanNullableText(raw.company_postal_code) || "",
    company_country: cleanNullableText(raw.company_country) || DEFAULT_APP_CONFIG.company_country,
    company_logo_path: cleanNullableText(raw.company_logo_path) || DEFAULT_APP_CONFIG.company_logo_path,
    purchase_links: purchaseLinks,
    purchase_link_digikey: cleanNullableText(
      purchaseLinksByKey.get("digikey")?.url || raw.purchase_link_digikey
    ) || DEFAULT_APP_CONFIG.purchase_link_digikey,
    purchase_link_mouser: cleanNullableText(
      purchaseLinksByKey.get("mouser")?.url || raw.purchase_link_mouser
    ) || DEFAULT_APP_CONFIG.purchase_link_mouser,
    purchase_link_nextgen_guitars: cleanNullableText(
      purchaseLinksByKey.get("nextgen_guitars")?.url || raw.purchase_link_nextgen_guitars
    ) || DEFAULT_APP_CONFIG.purchase_link_nextgen_guitars,
    purchase_link_thetubestore: cleanNullableText(
      purchaseLinksByKey.get("thetubestore")?.url || raw.purchase_link_thetubestore
    ) || DEFAULT_APP_CONFIG.purchase_link_thetubestore,
    tax_tps_number: cleanNullableText(raw.tax_tps_number) || "",
    tax_tvq_number: cleanNullableText(raw.tax_tvq_number) || "",
    tax_tps_rate: normalizeTaxRateNumber(raw.tax_tps_rate, DEFAULT_APP_CONFIG.tax_tps_rate),
    tax_tvq_rate: normalizeTaxRateNumber(raw.tax_tvq_rate, DEFAULT_APP_CONFIG.tax_tvq_rate),
  };
}

function isDeveloperModeEnabled() {
  return Boolean(state.appConfig?.developer_mode);
}

function isPodioSyncEnabled() {
  return state.appConfig?.podio_sync_enabled !== false;
}

function getDeveloperModeDraftValue() {
  if (typeof state.configDeveloperDraft === "boolean") return state.configDeveloperDraft;
  return Boolean(state.appConfig?.developer_mode);
}

function getPodioSyncDraftValue() {
  if (typeof state.configPodioSyncDraft === "boolean") return state.configPodioSyncDraft;
  return state.appConfig?.podio_sync_enabled !== false;
}

function isDeveloperModeDirty() {
  if (!state.appConfigLoaded) return false;
  return getDeveloperModeDraftValue() !== Boolean(state.appConfig?.developer_mode);
}

function isPodioSyncDirty() {
  if (!state.appConfigLoaded) return false;
  return getPodioSyncDraftValue() !== (state.appConfig?.podio_sync_enabled !== false);
}

function setConfigFeedback(messageEl, text, { clearAfterMs = 0 } = {}) {
  if (!messageEl) return;
  if (configMessageTimer) {
    clearTimeout(configMessageTimer);
    configMessageTimer = null;
  }
  messageEl.textContent = text || "";
  if (clearAfterMs > 0 && text) {
    configMessageTimer = setTimeout(() => {
      messageEl.textContent = "";
      configMessageTimer = null;
    }, clearAfterMs);
  }
}

function getConfigPurchaseLinksDraftValue() {
  if (Array.isArray(state.configPurchaseLinksDraft)) {
    return clonePurchaseLinks(state.configPurchaseLinksDraft);
  }
  const cfg = normalizeAppConfigRow(state.appConfig || DEFAULT_APP_CONFIG);
  return clonePurchaseLinks(cfg.purchase_links);
}

function renderConfigPurchaseLinksList(container, links, editing = false) {
  if (!container) return;
  const safeLinks = editing
    ? (Array.isArray(links) ? links : []).map((row, index) => {
      const src = row && typeof row === "object" ? row : {};
      return {
        key: normalizePurchaseLinkKey(src.key || `purchase_link_${index + 1}`, `purchase_link_${index + 1}`),
        title: cleanNullableText(src.title) || "",
        url: cleanNullableText(src.url) || "",
        note: cleanNullableText(src.note) || "",
        domain: cleanNullableText(src.domain) || "",
      };
    })
    : normalizePurchaseLinksCollection(links);
  if (!safeLinks.length) {
    container.innerHTML = editing
      ? `
        <div class="configuration-purchase-links-empty">Aucun lien. Ajoutez votre premier lien d'achat.</div>
        <button type="button" class="configuration-link-add-btn" data-config-purchase-link-add="1">Ajouter un lien</button>
      `
      : `<div class="configuration-purchase-links-empty">Aucun lien défini.</div>`;
    return;
  }

  if (editing) {
    const rows = safeLinks.map((link, index) => `
      <div class="configuration-purchase-link-row is-editing" data-config-purchase-link-row="${index}">
        <div class="configuration-purchase-link-edit-grid">
          <input
            class="configuration-rate-input configuration-purchase-link-input"
            type="text"
            maxlength="80"
            autocomplete="off"
            data-config-purchase-link-field="title"
            data-config-purchase-link-index="${index}"
            value="${escapeHtml(link.title || "")}"
            placeholder="Titre"
          >
          <input
            class="configuration-rate-input configuration-purchase-link-input"
            type="url"
            maxlength="280"
            autocomplete="off"
            data-config-purchase-link-field="url"
            data-config-purchase-link-index="${index}"
            value="${escapeHtml(link.url || "")}"
            placeholder="Lien https://..."
          >
          <input
            class="configuration-rate-input configuration-purchase-link-input"
            type="text"
            maxlength="220"
            autocomplete="off"
            data-config-purchase-link-field="note"
            data-config-purchase-link-index="${index}"
            value="${escapeHtml(link.note || "")}"
            placeholder="Note (optionnelle)"
          >
        </div>
        <button
          type="button"
          class="configuration-link-remove-btn"
          data-config-purchase-link-remove="${index}"
          aria-label="Supprimer ce lien"
          title="Supprimer"
        >&times;</button>
      </div>
    `).join("");

    container.innerHTML = `
      <div class="configuration-purchase-links-grid">${rows}</div>
      <button type="button" class="configuration-link-add-btn" data-config-purchase-link-add="1">Ajouter un lien</button>
    `;
    return;
  }

  const rows = safeLinks.map((link, index) => {
    const href = sanitizeExternalUrl(link.url);
    const title = cleanNullableText(link.title) || `Lien ${index + 1}`;
    const note = cleanNullableText(link.note);
    const info = note
      ? `
        <span class="configuration-label-with-info configuration-purchase-link-info-wrap">
          <button type="button" class="configuration-info-btn configuration-purchase-link-info-btn" aria-label="Information sur ${escapeHtml(title)}">i</button>
          <span class="configuration-info-tooltip" role="tooltip">${escapeHtml(note)}</span>
        </span>
      `
      : "";

    return `
      <div class="configuration-purchase-link-row">
        <div class="configuration-purchase-link-view-main">
          <span class="configuration-purchase-link-title">${escapeHtml(title)}</span>
          ${info}
        </div>
        ${href
          ? `<a class="configuration-purchase-link-url" href="${escapeHtml(href)}" target="_blank" rel="noopener noreferrer">${escapeHtml(href)}</a>`
          : `<span class="configuration-purchase-link-url is-empty">Non défini</span>`}
      </div>
    `;
  }).join("");
  container.innerHTML = `<div class="configuration-purchase-links-grid">${rows}</div>`;
}

function normalizeTaxReportRow(row) {
  const src = row && typeof row === "object" ? row : {};
  const reportYear = Number(src.report_year);
  if (!Number.isInteger(reportYear) || reportYear <= 0) return null;
  const fileSize = Number(src.file_size ?? src.size ?? null);
  return {
    id: Number(src.id) || reportYear,
    report_year: reportYear,
    pdf_filename: cleanNullableText(src.pdf_filename),
    storage_bucket: cleanNullableText(src.storage_bucket) || FISCAL_REPORTS_BUCKET,
    storage_path: cleanNullableText(src.storage_path),
    file_size: Number.isFinite(fileSize) && fileSize > 0 ? fileSize : null,
    facture_count: Math.max(0, Number(src.facture_count) || 0),
    montant_sans_taxes: Number(src.montant_sans_taxes) || 0,
    tps: Number(src.tps) || 0,
    tvq: Number(src.tvq) || 0,
    total: Number(src.total) || 0,
    generated_at: cleanNullableText(src.generated_at),
    created_at: cleanNullableText(src.created_at),
    updated_at: cleanNullableText(src.updated_at),
  };
}

function normalizeExpenseRow(row) {
  const src = row && typeof row === "object" ? row : {};
  const id = Number(src.id);
  if (!Number.isInteger(id) || id <= 0) return null;
  const amount = Number(src.amount ?? 0);
  return {
    id,
    title: cleanNullableText(src.title) || "Dépense sans titre",
    description: cleanNullableText(src.description),
    purchase_date: formatDateCell(src.purchase_date),
    amount: Number.isFinite(amount) ? Math.max(0, amount) : 0,
    invoice_filename: cleanNullableText(src.invoice_filename),
    invoice_bucket: cleanNullableText(src.invoice_bucket) || EXPENSES_STORAGE_BUCKET,
    invoice_path: cleanNullableText(src.invoice_path),
    invoice_content_type: cleanNullableText(src.invoice_content_type),
    preview_url: cleanNullableText(src.preview_url),
    created_at: cleanNullableText(src.created_at),
    updated_at: cleanNullableText(src.updated_at),
  };
}

function getTaxReportFileName(report, reportYear = null) {
  const fileName = cleanNullableText(report?.pdf_filename)
    || cleanNullableText(report?.storage_path)?.split("/").pop()
    || null;
  if (fileName) return fileName;
  const year = Number(reportYear ?? report?.report_year);
  return Number.isInteger(year) && year > 0 ? `rapport-fiscal-${year}.pdf` : "rapport-fiscal.pdf";
}

async function hydrateTaxReportStorageMeta(client, reports) {
  const normalizedReports = Array.isArray(reports) ? reports.filter(Boolean) : [];
  if (!client || !normalizedReports.length) return normalizedReports;

  const pathsByBucket = new Map();
  for (const report of normalizedReports) {
    const bucket = cleanNullableText(report?.storage_bucket) || FISCAL_REPORTS_BUCKET;
    const objectPath = cleanNullableText(report?.storage_path);
    if (!bucket || !objectPath) continue;
    if (!pathsByBucket.has(bucket)) pathsByBucket.set(bucket, new Set());
    pathsByBucket.get(bucket).add(objectPath);
  }

  if (!pathsByBucket.size) return normalizedReports;

  const sizeByObjectKey = new Map();
  for (const [bucket, pathSet] of pathsByBucket.entries()) {
    const objectPaths = Array.from(pathSet).filter(Boolean);
    if (!objectPaths.length) continue;
    try {
      const { data, error } = await client
        .schema("storage")
        .from("objects")
        .select("name,metadata")
        .eq("bucket_id", bucket)
        .in("name", objectPaths);
      if (error) {
        console.error("tax-reports: storage objects query error:", error);
        continue;
      }
      for (const row of Array.isArray(data) ? data : []) {
        const objectName = cleanNullableText(row?.name);
        const metadata = row?.metadata && typeof row.metadata === "object" ? row.metadata : {};
        const rawSize = Number(metadata.size ?? metadata.Size ?? metadata.length ?? null);
        if (!objectName) continue;
        sizeByObjectKey.set(
          `${bucket}:${objectName}`,
          Number.isFinite(rawSize) && rawSize > 0 ? rawSize : null,
        );
      }
    } catch (error) {
      console.error("tax-reports: storage metadata fetch failed:", error);
    }
  }

  return normalizedReports.map((report) => {
    const bucket = cleanNullableText(report?.storage_bucket) || FISCAL_REPORTS_BUCKET;
    const objectPath = cleanNullableText(report?.storage_path);
    if (!objectPath) return report;
    const size = sizeByObjectKey.get(`${bucket}:${objectPath}`);
    return size != null ? { ...report, file_size: size } : report;
  });
}

function getFiscalReportYears() {
  const years = new Set();
  for (const facture of state.factures || []) {
    const year = resolveFactureYear(facture);
    if (Number.isInteger(year) && year > 0) years.add(year);
  }
  for (const report of state.taxReports || []) {
    const year = Number(report?.report_year);
    if (Number.isInteger(year) && year > 0) years.add(year);
  }
  return Array.from(years).sort((a, b) => b - a);
}

function renderConfigTaxReportsList(container) {
  if (!container) return;
  if (state.taxReportsLoading) {
    container.innerHTML = `<div class="configuration-tax-reports-empty">Chargement des rapports fiscaux...</div>`;
    return;
  }

  const reportMap = new Map(
    (state.taxReports || [])
      .map((row) => normalizeTaxReportRow(row))
      .filter(Boolean)
      .map((row) => [row.report_year, row])
  );
  const years = getFiscalReportYears();
  if (!years.length) {
    const errorText = cleanNullableText(state.taxReportsLoadError);
    container.innerHTML = `<div class="configuration-tax-reports-empty">${escapeHtml(errorText || "Aucun document fiscal pour l'instant.")}</div>`;
    return;
  }

  const header = `
    <div class="configuration-tax-report-table-head" aria-hidden="true">
      <span>Année</span>
      <span>Montant sans taxe</span>
      <span>TPS</span>
      <span>TVQ</span>
      <span>Total</span>
      <span class="configuration-tax-report-table-head-actions">Actions</span>
    </div>
  `;

  const rows = years.map((year) => {
    const report = reportMap.get(year) || null;
    const isGenerating = Number(state.taxReportsGeneratingYear) === Number(year);
    const canOpen = !!cleanNullableText(report?.storage_path);
    const fileName = getTaxReportFileName(report, year);
    const sizeText = Number.isFinite(Number(report?.file_size)) && Number(report.file_size) > 0
      ? formatBytes(Number(report.file_size))
      : "";
    const generatedText = report?.generated_at ? formatDateCell(report.generated_at) : "";
    const detailParts = canOpen
      ? [fileName, sizeText, generatedText].filter(Boolean)
      : ["Aucun PDF généré"];
    return `
      <div class="configuration-tax-report-row ${canOpen ? "is-ready" : "is-empty"}">
        <div class="configuration-tax-report-row-main">
          <span class="configuration-tax-report-cell configuration-tax-report-cell-year">${year}</span>
          <span class="configuration-tax-report-cell">${escapeHtml(formatMoney(report?.montant_sans_taxes || 0))}</span>
          <span class="configuration-tax-report-cell">${escapeHtml(formatMoney(report?.tps || 0))}</span>
          <span class="configuration-tax-report-cell">${escapeHtml(formatMoney(report?.tvq || 0))}</span>
          <span class="configuration-tax-report-cell configuration-tax-report-cell-total">${escapeHtml(formatMoney(report?.total || 0))}</span>
          <span class="configuration-tax-report-cell configuration-tax-report-actions">
            <button
              type="button"
              class="configuration-tax-report-icon-btn"
              data-tax-report-open="${year}"
              title="Voir le PDF"
              aria-label="Voir le PDF"
              ${canOpen ? "" : "disabled"}
            >${FACTURE_PREVIEW_EYE_SVG}</button>
            <button
              type="button"
              class="configuration-tax-report-generate-btn"
              data-tax-report-generate="${year}"
              ${isGenerating ? "disabled" : ""}
            >${isGenerating ? "Génération..." : (report ? "Régénérer" : "Générer")}</button>
            <button
              type="button"
              class="configuration-tax-report-icon-btn"
              data-tax-report-download="${year}"
              title="Télécharger le PDF"
              aria-label="Télécharger le PDF"
              ${canOpen ? "" : "disabled"}
            >${PROJECT_FILE_DOWNLOAD_SVG}</button>
          </span>
        </div>
        <div class="configuration-tax-report-row-sub">
          ${detailParts.map((part) => `<span class="configuration-tax-report-detail-part">${escapeHtml(part)}</span>`).join("")}
        </div>
      </div>
    `;
  }).join("");

  const errorHtml = cleanNullableText(state.taxReportsLoadError)
    ? `<div class="configuration-tax-reports-empty">${escapeHtml(state.taxReportsLoadError)}</div>`
    : "";
  container.innerHTML = `${errorHtml}<div class="configuration-tax-report-table">${header}${rows}</div>`;
}

function nextLocalEntityId() {
  localEntitySequence += 1;
  return Date.now() + localEntitySequence;
}

function nextLocalPodioItemId() {
  return -nextLocalEntityId();
}

function nextLocalAppItemId(items) {
  const max = (items || []).reduce((acc, item) => {
    const value = Number(item?.podio_app_item_id);
    return Number.isFinite(value) && value > acc ? value : acc;
  }, 0);
  return max + 1;
}

function getConfiguredFactureLaborRate() {
  const cfg = state.appConfig || DEFAULT_APP_CONFIG;
  return normalizeNonNegativeNumber(cfg.labor_rate_standard, DEFAULT_APP_CONFIG.labor_rate_standard);
}

function getConfiguredFactureTaxRates() {
  const cfg = state.appConfig || DEFAULT_APP_CONFIG;
  return {
    tps: normalizeTaxRateNumber(cfg.tax_tps_rate, DEFAULT_APP_CONFIG.tax_tps_rate),
    tvq: normalizeTaxRateNumber(cfg.tax_tvq_rate, DEFAULT_APP_CONFIG.tax_tvq_rate),
  };
}

function getCompanySnapshotFromConfig() {
  const cfg = normalizeAppConfigRow(state.appConfig || DEFAULT_APP_CONFIG);
  return {
    name: cfg.company_name || null,
    phone: cfg.company_phone || null,
    email: cfg.company_email || null,
    address_line1: cfg.company_address_line1 || null,
    address_line2: cfg.company_address_line2 || null,
    city: cfg.company_city || null,
    province: cfg.company_province || null,
    postal_code: cfg.company_postal_code || null,
    country: cfg.company_country || null,
    logo_path: cfg.company_logo_path || null,
    tps_number: cfg.tax_tps_number || null,
    tvq_number: cfg.tax_tvq_number || null,
  };
}

function sanitizeExternalUrl(url) {
  const value = cleanNullableText(url);
  if (!value) return null;
  if (/^https?:\/\//i.test(value)) return value;
  return `https://${value.replace(/^\/+/, "")}`;
}

function getRepairPartsSupplierLinks() {
  const cfg = normalizeAppConfigRow(state.appConfig || DEFAULT_APP_CONFIG);
  const links = normalizePurchaseLinksCollection(cfg.purchase_links, cfg);
  return links.map((supplier, index) => {
    const fallbackPreset = REPAIR_PARTS_SUPPLIERS[index] || null;
    const href = sanitizeExternalUrl(supplier.url);
    const domain = extractDomainFromUrl(href, supplier.domain || fallbackPreset?.domain || "");
    const label = cleanNullableText(supplier.title) || cleanNullableText(fallbackPreset?.title) || `Lien ${index + 1}`;
    return {
      key: normalizePurchaseLinkKey(supplier.key, `purchase_link_${index + 1}`),
      name: label,
      domain,
      note: cleanNullableText(supplier.note),
      href,
      faviconUrl: domain
        ? `https://www.google.com/s2/favicons?domain=${encodeURIComponent(domain)}&sz=64`
        : "",
    };
  });
}

async function saveAppConfig(payload) {
  const client = state.authClient;
  if (!client) throw new Error("Connexion Supabase indisponible.");
  const hasPurchaseLinksPayload = Object.prototype.hasOwnProperty.call(payload || {}, "purchase_links");
  const cleaned = normalizeAppConfigRow({
    ...(state.appConfig || DEFAULT_APP_CONFIG),
    ...(payload || {}),
  });
  const purchaseLinksByKey = new Map(
    normalizePurchaseLinksCollection(cleaned.purchase_links).map((link) => [normalizePurchaseLinkKey(link.key), link])
  );
  const legacyUrlFromLinks = (key) => cleanNullableText(purchaseLinksByKey.get(key)?.url);
  const row = {
    id: 1,
    labor_rate_standard: cleaned.labor_rate_standard,
    labor_rate_secondary: cleaned.labor_rate_secondary,
    labor_rate_diagnostic: cleaned.labor_rate_diagnostic,
    labor_rate_estimation: cleaned.labor_rate_estimation,
    developer_mode: Boolean(cleaned.developer_mode),
    podio_sync_enabled: cleaned.podio_sync_enabled !== false,
    company_name: cleaned.company_name || null,
    company_phone: cleaned.company_phone || null,
    company_email: cleaned.company_email || null,
    company_address_line1: cleaned.company_address_line1 || null,
    company_address_line2: cleaned.company_address_line2 || null,
    company_city: cleaned.company_city || null,
    company_province: cleaned.company_province || null,
    company_postal_code: cleaned.company_postal_code || null,
    company_country: cleaned.company_country || null,
    company_logo_path: cleaned.company_logo_path || null,
    purchase_links: clonePurchaseLinks(cleaned.purchase_links),
    purchase_link_digikey: hasPurchaseLinksPayload
      ? (legacyUrlFromLinks("digikey") || null)
      : (cleaned.purchase_link_digikey || null),
    purchase_link_mouser: hasPurchaseLinksPayload
      ? (legacyUrlFromLinks("mouser") || null)
      : (cleaned.purchase_link_mouser || null),
    purchase_link_nextgen_guitars: hasPurchaseLinksPayload
      ? (legacyUrlFromLinks("nextgen_guitars") || null)
      : (cleaned.purchase_link_nextgen_guitars || null),
    purchase_link_thetubestore: hasPurchaseLinksPayload
      ? (legacyUrlFromLinks("thetubestore") || null)
      : (cleaned.purchase_link_thetubestore || null),
    tax_tps_number: cleaned.tax_tps_number || null,
    tax_tvq_number: cleaned.tax_tvq_number || null,
    tax_tps_rate: cleaned.tax_tps_rate,
    tax_tvq_rate: cleaned.tax_tvq_rate,
  };
  const { data, error } = await client
    .from("app_config")
    .upsert(row, { onConflict: "id" })
    .select("*")
    .maybeSingle();
  if (error) throw error;
  state.appConfig = normalizeAppConfigRow(data || row);
  state.appConfigLoaded = true;
  void updateMenuCompanyLogo();
  return state.appConfig;
}

function renderConfigurationView(keepMessage = false) {
  const standardInput = $("configLaborRateStandardInput");
  const secondaryInput = $("configLaborRateSecondaryInput");
  const diagnosticInput = $("configLaborRateDiagnosticInput");
  const estimationInput = $("configLaborRateEstimationInput");
  const developerModeOffBtn = $("configDeveloperModeOffBtn");
  const developerModeSwitchBtn = $("configDeveloperModeSwitchBtn");
  const developerModeOnBtn = $("configDeveloperModeOnBtn");
  const companyNameInput = $("configCompanyNameInput");
  const companyPhoneInput = $("configCompanyPhoneInput");
  const companyEmailInput = $("configCompanyEmailInput");
  const companyAddress1Input = $("configCompanyAddress1Input");
  const companyAddress2Input = $("configCompanyAddress2Input");
  const companyCityInput = $("configCompanyCityInput");
  const companyProvinceInput = $("configCompanyProvinceInput");
  const companyPostalCodeInput = $("configCompanyPostalCodeInput");
  const companyCountryInput = $("configCompanyCountryInput");
  const companyLogoPathInput = $("configCompanyLogoPathInput");
  const purchaseLinksList = $("configPurchaseLinksList");
  const tpsNumberInput = $("configTpsNumberInput");
  const tpsRateInput = $("configTpsRateInput");
  const tvqNumberInput = $("configTvqNumberInput");
  const tvqRateInput = $("configTvqRateInput");
  const standardText = $("configLaborRateStandardText");
  const secondaryText = $("configLaborRateSecondaryText");
  const diagnosticText = $("configLaborRateDiagnosticText");
  const estimationText = $("configLaborRateEstimationText");
  const companyNameText = $("configCompanyNameText");
  const companyPhoneText = $("configCompanyPhoneText");
  const companyEmailText = $("configCompanyEmailText");
  const companyAddress1Text = $("configCompanyAddress1Text");
  const companyAddress2Text = $("configCompanyAddress2Text");
  const companyCityText = $("configCompanyCityText");
  const companyProvinceText = $("configCompanyProvinceText");
  const companyPostalCodeText = $("configCompanyPostalCodeText");
  const companyCountryText = $("configCompanyCountryText");
  const companyLogoPathText = $("configCompanyLogoPathText");
  const tpsNumberText = $("configTpsNumberText");
  const tpsRateText = $("configTpsRateText");
  const tvqNumberText = $("configTvqNumberText");
  const tvqRateText = $("configTvqRateText");
  const honorairesEditBtn = $("configHonorairesEditBtn");
  const entrepriseEditBtn = $("configEntrepriseEditBtn");
  const achatsEditBtn = $("configAchatsEditBtn");
  const honorairesControls = $("configHonorairesControls");
  const entrepriseControls = $("configEntrepriseControls");
  const achatsControls = $("configAchatsControls");
  const developpementControls = $("configDeveloppementControls");
  const taxReportsList = $("configTaxReportsList");
  const messageEl = $("configMessage");
  if (
    !standardInput
    || !secondaryInput
    || !diagnosticInput
    || !estimationInput
    || !developerModeOffBtn
    || !developerModeSwitchBtn
    || !developerModeOnBtn
    || !companyNameInput
    || !companyPhoneInput
    || !companyEmailInput
    || !companyAddress1Input
    || !companyAddress2Input
    || !companyCityInput
    || !companyProvinceInput
    || !companyPostalCodeInput
    || !companyCountryInput
    || !companyLogoPathInput
    || !purchaseLinksList
    || !tpsNumberInput
    || !tpsRateInput
    || !tvqNumberInput
    || !tvqRateInput
    || !standardText
    || !secondaryText
    || !diagnosticText
    || !estimationText
    || !companyNameText
    || !companyPhoneText
    || !companyEmailText
    || !companyAddress1Text
    || !companyAddress2Text
    || !companyCityText
    || !companyProvinceText
    || !companyPostalCodeText
    || !companyCountryText
    || !companyLogoPathText
    || !tpsNumberText
    || !tpsRateText
    || !tvqNumberText
    || !tvqRateText
    || !taxReportsList
  ) return;

  const setFieldMode = (inputEl, textEl, editing) => {
    inputEl.hidden = !editing;
    inputEl.disabled = !editing;
    textEl.hidden = editing;
  };
  const formatRateLine = (rate) => `${formatMoneyCompact(rate)}/h`;
  const formatTaxRateLine = (rate) => `${(normalizeTaxRateNumber(rate, 0) * 100).toLocaleString("fr-CA", { minimumFractionDigits: 0, maximumFractionDigits: 3 })} %`;
  const editingHonoraires = !!state.configEditing?.honoraires;
  const editingEntreprise = !!state.configEditing?.entreprise;
  const editingAchats = !!state.configEditing?.achats;
  const developerModeValue = getDeveloperModeDraftValue();
  const developerModeDirty = isDeveloperModeDirty();

  if (!state.appConfigLoaded) {
    standardInput.value = "";
    secondaryInput.value = "";
    diagnosticInput.value = "";
    estimationInput.value = "";
    companyNameInput.value = "";
    companyPhoneInput.value = "";
    companyEmailInput.value = "";
    companyAddress1Input.value = "";
    companyAddress2Input.value = "";
    companyCityInput.value = "";
    companyProvinceInput.value = "";
    companyPostalCodeInput.value = "";
    companyCountryInput.value = "";
    companyLogoPathInput.value = "";
    tpsNumberInput.value = "";
    tpsRateInput.value = "";
    tvqNumberInput.value = "";
    tvqRateInput.value = "";
    standardText.textContent = "-";
    secondaryText.textContent = "-";
    diagnosticText.textContent = "-";
    estimationText.textContent = "-";
    companyNameText.textContent = "-";
    companyPhoneText.textContent = "-";
    companyEmailText.textContent = "-";
    companyAddress1Text.textContent = "-";
    companyAddress2Text.textContent = "-";
    companyCityText.textContent = "-";
    companyProvinceText.textContent = "-";
    companyPostalCodeText.textContent = "-";
    companyCountryText.textContent = "-";
    companyLogoPathText.textContent = "-";
    tpsNumberText.textContent = "-";
    tpsRateText.textContent = "-";
    tvqNumberText.textContent = "-";
    tvqRateText.textContent = "-";
    setFieldMode(standardInput, standardText, false);
    setFieldMode(secondaryInput, secondaryText, false);
    setFieldMode(diagnosticInput, diagnosticText, false);
    setFieldMode(estimationInput, estimationText, false);
    setFieldMode(companyNameInput, companyNameText, false);
    setFieldMode(companyPhoneInput, companyPhoneText, false);
    setFieldMode(companyEmailInput, companyEmailText, false);
    setFieldMode(companyAddress1Input, companyAddress1Text, false);
    setFieldMode(companyAddress2Input, companyAddress2Text, false);
    setFieldMode(companyCityInput, companyCityText, false);
    setFieldMode(companyProvinceInput, companyProvinceText, false);
    setFieldMode(companyPostalCodeInput, companyPostalCodeText, false);
    setFieldMode(companyCountryInput, companyCountryText, false);
    setFieldMode(companyLogoPathInput, companyLogoPathText, false);
    setFieldMode(tpsNumberInput, tpsNumberText, false);
    setFieldMode(tpsRateInput, tpsRateText, false);
    setFieldMode(tvqNumberInput, tvqNumberText, false);
    setFieldMode(tvqRateInput, tvqRateText, false);
    renderConfigPurchaseLinksList(purchaseLinksList, [], false);
    renderConfigTaxReportsList(taxReportsList);
    if (honorairesEditBtn) honorairesEditBtn.disabled = true;
    if (entrepriseEditBtn) entrepriseEditBtn.disabled = true;
    if (achatsEditBtn) achatsEditBtn.disabled = true;
    if (honorairesControls) honorairesControls.hidden = true;
    if (entrepriseControls) entrepriseControls.hidden = true;
    if (achatsControls) achatsControls.hidden = true;
    if (developpementControls) developpementControls.hidden = true;
    developerModeSwitchBtn.classList.remove("is-on");
    developerModeSwitchBtn.setAttribute("aria-checked", "false");
    developerModeSwitchBtn.disabled = true;
    developerModeOffBtn.disabled = true;
    developerModeOnBtn.disabled = true;
    developerModeOffBtn.classList.remove("is-active");
    developerModeOnBtn.classList.remove("is-active");
    setConfigurationTab(state.configActiveTab || "honoraires");
    if (!keepMessage) setConfigFeedback(messageEl, "Chargement de la configuration...");
    renderDeveloperModeBanner();
    return;
  }

  const cfg = state.appConfig ? normalizeAppConfigRow(state.appConfig) : null;
  standardInput.value = cfg ? String(cfg.labor_rate_standard) : "";
  secondaryInput.value = cfg ? String(cfg.labor_rate_secondary) : "";
  diagnosticInput.value = cfg ? String(cfg.labor_rate_diagnostic) : "";
  estimationInput.value = cfg ? String(cfg.labor_rate_estimation) : "";
  companyNameInput.value = cfg?.company_name || "";
  companyPhoneInput.value = cfg?.company_phone || "";
  companyEmailInput.value = cfg?.company_email || "";
  companyAddress1Input.value = cfg?.company_address_line1 || "";
  companyAddress2Input.value = cfg?.company_address_line2 || "";
  companyCityInput.value = cfg?.company_city || "";
  companyProvinceInput.value = cfg?.company_province || "";
  companyPostalCodeInput.value = cfg?.company_postal_code || "";
  companyCountryInput.value = cfg?.company_country || "";
  companyLogoPathInput.value = cfg?.company_logo_path || "";
  tpsNumberInput.value = cfg?.tax_tps_number || "";
  tpsRateInput.value = cfg ? String(normalizeTaxRateNumber(cfg.tax_tps_rate, DEFAULT_APP_CONFIG.tax_tps_rate) * 100) : "";
  tvqNumberInput.value = cfg?.tax_tvq_number || "";
  tvqRateInput.value = cfg ? String(normalizeTaxRateNumber(cfg.tax_tvq_rate, DEFAULT_APP_CONFIG.tax_tvq_rate) * 100) : "";
  standardText.textContent = cfg ? formatRateLine(cfg.labor_rate_standard) : "-";
  secondaryText.textContent = cfg ? formatRateLine(cfg.labor_rate_secondary) : "-";
  diagnosticText.textContent = cfg ? formatRateLine(cfg.labor_rate_diagnostic) : "-";
  estimationText.textContent = cfg ? formatRateLine(cfg.labor_rate_estimation) : "-";
  companyNameText.textContent = cfg?.company_name || "Non défini";
  companyPhoneText.textContent = cfg?.company_phone || "Non défini";
  companyEmailText.textContent = cfg?.company_email || "Non défini";
  companyAddress1Text.textContent = cfg?.company_address_line1 || "Non défini";
  companyAddress2Text.textContent = cfg?.company_address_line2 || "Non défini";
  companyCityText.textContent = cfg?.company_city || "Non défini";
  companyProvinceText.textContent = cfg?.company_province || "Non défini";
  companyPostalCodeText.textContent = cfg?.company_postal_code || "Non défini";
  companyCountryText.textContent = cfg?.company_country || "Non défini";
  companyLogoPathText.textContent = cfg?.company_logo_path || "Non défini";
  tpsNumberText.textContent = cfg?.tax_tps_number || "Non defini";
  tpsRateText.textContent = cfg ? formatTaxRateLine(cfg.tax_tps_rate) : "-";
  tvqNumberText.textContent = cfg?.tax_tvq_number || "Non defini";
  tvqRateText.textContent = cfg ? formatTaxRateLine(cfg.tax_tvq_rate) : "-";
  const purchaseLinksToRender = editingAchats
    ? getConfigPurchaseLinksDraftValue()
    : cfg?.purchase_links;
  renderConfigPurchaseLinksList(purchaseLinksList, purchaseLinksToRender, editingAchats);
  renderConfigTaxReportsList(taxReportsList);

  setFieldMode(standardInput, standardText, editingHonoraires);
  setFieldMode(secondaryInput, secondaryText, editingHonoraires);
  setFieldMode(diagnosticInput, diagnosticText, editingHonoraires);
  setFieldMode(estimationInput, estimationText, editingHonoraires);
  setFieldMode(companyNameInput, companyNameText, editingEntreprise);
  setFieldMode(companyPhoneInput, companyPhoneText, editingEntreprise);
  setFieldMode(companyEmailInput, companyEmailText, editingEntreprise);
  setFieldMode(companyAddress1Input, companyAddress1Text, editingEntreprise);
  setFieldMode(companyAddress2Input, companyAddress2Text, editingEntreprise);
  setFieldMode(companyCityInput, companyCityText, editingEntreprise);
  setFieldMode(companyProvinceInput, companyProvinceText, editingEntreprise);
  setFieldMode(companyPostalCodeInput, companyPostalCodeText, editingEntreprise);
  setFieldMode(companyCountryInput, companyCountryText, editingEntreprise);
  setFieldMode(companyLogoPathInput, companyLogoPathText, editingEntreprise);
  setFieldMode(tpsNumberInput, tpsNumberText, editingEntreprise);
  setFieldMode(tpsRateInput, tpsRateText, editingEntreprise);
  setFieldMode(tvqNumberInput, tvqNumberText, editingEntreprise);
  setFieldMode(tvqRateInput, tvqRateText, editingEntreprise);

  if (honorairesEditBtn) {
    honorairesEditBtn.disabled = false;
    honorairesEditBtn.classList.toggle("is-active", editingHonoraires);
    honorairesEditBtn.setAttribute("aria-pressed", String(editingHonoraires));
  }
  if (entrepriseEditBtn) {
    entrepriseEditBtn.disabled = false;
    entrepriseEditBtn.classList.toggle("is-active", editingEntreprise);
    entrepriseEditBtn.setAttribute("aria-pressed", String(editingEntreprise));
  }
  if (achatsEditBtn) {
    achatsEditBtn.disabled = false;
    achatsEditBtn.classList.toggle("is-active", editingAchats);
    achatsEditBtn.setAttribute("aria-pressed", String(editingAchats));
  }
  if (honorairesControls) honorairesControls.hidden = !editingHonoraires;
  if (entrepriseControls) entrepriseControls.hidden = !editingEntreprise;
  if (achatsControls) achatsControls.hidden = !editingAchats;
  if (developpementControls) developpementControls.hidden = !developerModeDirty;
  developerModeSwitchBtn.disabled = false;
  developerModeOffBtn.disabled = false;
  developerModeOnBtn.disabled = false;
  developerModeSwitchBtn.classList.toggle("is-on", developerModeValue);
  developerModeSwitchBtn.setAttribute("aria-checked", String(developerModeValue));
  developerModeOffBtn.classList.toggle("is-active", !developerModeValue);
  developerModeOnBtn.classList.toggle("is-active", developerModeValue);
  setConfigurationTab(state.configActiveTab || "honoraires");
  if (!keepMessage) setConfigFeedback(messageEl, "");
  renderDeveloperModeBanner();
}

function setConfigurationTab(tabKey) {
  const raw = String(tabKey ?? "").trim().toLowerCase();
  const wanted = raw === "entreprise" || raw === "achats" || raw === "developpement" || raw === "documents-fiscaux"
    ? raw
    : "honoraires";
  state.configActiveTab = wanted;

  const tabButtons = document.querySelectorAll("[data-config-tab]");
  const tabPanels = document.querySelectorAll("[data-config-tab-panel]");
  if (!tabButtons.length && !tabPanels.length) return;

  tabButtons.forEach((btn) => {
    const key = String(btn.getAttribute("data-config-tab") ?? "").trim().toLowerCase();
    const isActive = key === wanted;
    btn.classList.toggle("is-active", isActive);
    btn.setAttribute("aria-selected", String(isActive));
    btn.setAttribute("tabindex", isActive ? "0" : "-1");
  });

  tabPanels.forEach((panel) => {
    const key = String(panel.getAttribute("data-config-tab-panel") ?? "").trim().toLowerCase();
    const isActive = key === wanted;
    panel.classList.toggle("is-active", isActive);
    panel.hidden = !isActive;
  });
}

function setBonsRepairsTab(tabKey, options = {}) {
  const raw = String(tabKey ?? "").trim().toLowerCase();
  const wanted = BONS_REPAIRS_TABS.has(raw) ? raw : "board";
  state.bonsRepairsActiveTab = wanted;

  const tabButtons = document.querySelectorAll("[data-bons-repairs-tab]");
  tabButtons.forEach((btn) => {
    const key = String(btn.getAttribute("data-bons-repairs-tab") ?? "").trim().toLowerCase();
    const isActive = key === wanted;
    btn.classList.toggle("is-active", isActive);
    btn.setAttribute("aria-selected", String(isActive));
    btn.setAttribute("tabindex", isActive ? "0" : "-1");
  });

  const tabPanels = document.querySelectorAll("[data-bons-repairs-tab-panel]");
  tabPanels.forEach((panel) => {
    const key = String(panel.getAttribute("data-bons-repairs-tab-panel") ?? "").trim().toLowerCase();
    const isActive = key === wanted;
    panel.classList.toggle("is-active", isActive);
    panel.hidden = !isActive;
  });

  const boardMeta = $("bonsMeta");
  const archivesMeta = $("bonsOtherMeta");
  if (boardMeta) boardMeta.hidden = wanted !== "board";
  if (archivesMeta) archivesMeta.hidden = wanted !== "archives";

  if (options?.render === false) return;
  if (wanted === "archives") {
    renderBonsOtherList();
  } else {
    renderBonsBoard();
  }
}

function setProjectsTab(tabKey, options = {}) {
  const raw = String(tabKey ?? "").trim().toLowerCase();
  const wanted = PROJECTS_TABS.has(raw) ? raw : "board";
  state.projectsActiveTab = wanted;

  const tabButtons = document.querySelectorAll("[data-projects-tab]");
  tabButtons.forEach((btn) => {
    const key = String(btn.getAttribute("data-projects-tab") ?? "").trim().toLowerCase();
    const isActive = key === wanted;
    btn.classList.toggle("is-active", isActive);
    btn.setAttribute("aria-selected", String(isActive));
    btn.setAttribute("tabindex", isActive ? "0" : "-1");
  });

  const tabPanels = document.querySelectorAll("[data-projects-tab-panel]");
  tabPanels.forEach((panel) => {
    const key = String(panel.getAttribute("data-projects-tab-panel") ?? "").trim().toLowerCase();
    const isActive = key === wanted;
    panel.classList.toggle("is-active", isActive);
    panel.hidden = !isActive;
  });

  if (options?.render === false) return;
  renderProjectsList();
}

function setBonsProjectsTab(tabKey, options = {}) {
  const raw = String(tabKey ?? "").trim().toLowerCase();
  const wanted = BONS_PROJECTS_TABS.has(raw) ? raw : "bons";
  const previous = state.bonsProjectsActiveTab;
  state.bonsProjectsActiveTab = wanted;

  const tabButtons = document.querySelectorAll("[data-bons-projects-tab]");
  tabButtons.forEach((btn) => {
    const key = String(btn.getAttribute("data-bons-projects-tab") ?? "").trim().toLowerCase();
    const isActive = key === wanted;
    btn.classList.toggle("is-active", isActive);
    btn.setAttribute("aria-selected", String(isActive));
    btn.setAttribute("tabindex", isActive ? "0" : "-1");
  });

  const tabPanels = document.querySelectorAll("[data-bons-projects-tab-panel]");
  tabPanels.forEach((panel) => {
    const key = String(panel.getAttribute("data-bons-projects-tab-panel") ?? "").trim().toLowerCase();
    const isActive = key === wanted;
    panel.classList.toggle("is-active", isActive);
    panel.hidden = !isActive;
  });

  const bonsHeadActions = $("bonsHeadActions");
  const projectsHeadActions = $("projectsHeadActions");
  if (bonsHeadActions) bonsHeadActions.hidden = wanted !== "bons";
  if (projectsHeadActions) projectsHeadActions.hidden = wanted !== "projets";

  if (wanted === "projets") {
    setProjectsTab(state.projectsActiveTab, options);
  } else {
    setBonsRepairsTab(state.bonsRepairsActiveTab, options);
  }
}

function setFacturationTab(tabKey, options = {}) {
  const raw = String(tabKey ?? "").trim().toLowerCase();
  const wanted = FACTURATION_TABS.has(raw) ? raw : "factures";
  const previous = state.facturationActiveTab;
  state.facturationActiveTab = wanted;

  const tabButtons = document.querySelectorAll("[data-facturation-tab]");
  tabButtons.forEach((btn) => {
    const key = String(btn.getAttribute("data-facturation-tab") ?? "").trim().toLowerCase();
    const isActive = key === wanted;
    btn.classList.toggle("is-active", isActive);
    btn.setAttribute("aria-selected", String(isActive));
    btn.setAttribute("tabindex", isActive ? "0" : "-1");
  });

  const createBtn = $("facturesCreateBtn");
  if (createBtn) {
    createBtn.hidden = false;
    if (wanted === "estimations") {
      createBtn.textContent = "Cr\u00e9er une nouvelle estimation";
    } else if (wanted === "soumissions") {
      createBtn.textContent = "Cr\u00e9er une nouvelle soumission";
    } else {
      createBtn.textContent = "Cr\u00e9er une nouvelle facture";
    }
  }

  const searchInput = $("facturesSearchInput");
  if (searchInput) {
    if (wanted === "estimations") {
      searchInput.placeholder = "Rechercher une estimation (num\u00e9ro, client, statut...)";
      searchInput.value = state.devisSearchQuery || "";
    } else if (wanted === "soumissions") {
      searchInput.placeholder = "Rechercher une soumission (num\u00e9ro, client, statut...)";
      searchInput.value = state.devisSearchQuery || "";
    } else {
      searchInput.placeholder = "Rechercher une facture (num\u00e9ro, client, REP, \u00e9tat...)";
      searchInput.value = state.facturesSearchQuery || "";
    }
  }

  const invoiceFilters = $("facturesToolbarFiltersFactures");
  const docsFilters = $("facturesToolbarFiltersDevis");
  if (invoiceFilters) invoiceFilters.hidden = wanted !== "factures";
  if (docsFilters) docsFilters.hidden = wanted === "factures";

  const statusDocs = $("devisStatusFilterSelect");
  if (statusDocs) {
    const allowed = new Set(["all", "en_redaction", "emise", "accepte", "refuse", "annule", "echoue"]);
    const current = String(state.devisStatusFilter ?? "all");
    const resolved = allowed.has(current) ? current : "all";
    state.devisStatusFilter = resolved;
    statusDocs.value = resolved;
  }

  const shouldRender = options?.render !== false;
  if (shouldRender && previous !== wanted) {
    renderFacturesList();
  } else if (shouldRender && previous === wanted) {
    renderFacturesList();
  }
}

function inferSupabaseUrl(projectRef) {
  const ref = (projectRef || "").trim();
  if (!ref) return "";
  return `https://${ref}.supabase.co`;
}

function currentPageName() {
  const page = (window.location.pathname.split("/").pop() || "index.html").toLowerCase();
  return page || "index.html";
}

function isLoginPage() {
  return currentPageName() === "login.html";
}

function isAppPage() {
  return currentPageName() === "index.html" || currentPageName() === "";
}

function goToLogin() {
  if (!isLoginPage()) window.location.href = "./login.html";
}

function goToApp() {
  if (!isAppPage()) window.location.href = "./index.html";
}

function setLoginStatus(text, kind = "muted") {
  const el = $("loginStatus");
  if (!el) return;
  el.textContent = text;
  el.classList.remove("ok", "warn", "bad", "muted");
  el.classList.add(kind);
}

function setPasswordVisibility(visible) {
  const input = $("loginPassword");
  const eye = $("passwordEyeIcon");
  if (!input || !eye) return;
  input.type = visible ? "text" : "password";
  eye.innerHTML = visible ? EYE_OPEN_SVG : EYE_CLOSED_SVG;
}

function ensureAuthClient() {
  if (!window.supabase || !window.supabase.createClient) {
    setLoginStatus("Librairie Supabase absente.", "bad");
    return null;
  }
  if (!FIXED_SUPABASE.anonKey) {
    setLoginStatus("Anon key manquante dans app.js.", "bad");
    return null;
  }
  if (state.authClient) return state.authClient;

  state.authClient = window.supabase.createClient(
    inferSupabaseUrl(FIXED_SUPABASE.projectRef),
    FIXED_SUPABASE.anonKey,
    {
      auth: {
        persistSession: true,
        autoRefreshToken: true,
        detectSessionInUrl: true
      }
    }
  );
  return state.authClient;
}

function resolveUsername(user) {
  if (!user) return "Utilisateur";
  const meta = user.user_metadata || {};
  return (
    state.profile?.display_name ||
    meta.display_name ||
    meta.full_name ||
    meta.name ||
    (user.email ? user.email.split("@")[0] : "Utilisateur")
  );
}

function formatRole(roleValue) {
  const role = normalizeText(roleValue);
  if (role === "admin") return "Admin";
  if (role === "employe" || role === "employes" || role === "technicien" || role === "manager") return "Employés";
  return "";
}

function updateMenuUser() {
  const hello = $("menuHello");
  const role = $("menuRole");
  if (!hello) return;
  hello.textContent = `Bonjour ${resolveUsername(state.session?.user)}`;
  if (role) role.textContent = formatRole(state.profile?.role);
}

function updateMenuCompanyLogo() {
  const wrap = $("menuBrandLogoWrap");
  const img = $("menuBrandLogo");
  if (!wrap || !img) return;
  img.src = state.theme === "dark" ? "./ressources/logo-wh.png" : "./ressources/logo-bk.png";
  wrap.hidden = false;
}

function canonicalRepairStatus(statusLabel) {
  const s = normalizeText(statusLabel);
  if (s === "dossier ouvert") return "dossier_ouvert";
  if (s === "en test") return "en_test";
  if (s === "en cours") return "en_cours";
  if (s === "en attente de piece" || s === "en attente de pieces") return "en_attente_pieces";
  if (s === "termine" || s === "terminee") return "termine";
  if (s === "remis au client" || s === "remis client") return "remis_client";
  if (s === "non reparable") return "non_reparable";
  if (s === "annule" || s === "annulee") return "annule";
  return "other";
}

function getBonsColumnByKey(columnKey) {
  return BONS_COLUMNS.find((col) => col.key === columnKey) || null;
}

function isBonsDefaultColumn(columnKey) {
  return BONS_DEFAULT_COLUMN_KEYS.has(String(columnKey || ""));
}

function isBonsColumnKey(columnKey) {
  return Boolean(getBonsColumnByKey(columnKey));
}

function getBonsColumnKeyFromDragEvent(ev, bonsBoard) {
  if (!bonsBoard) return null;
  const targetEl = ev?.target instanceof Element ? ev.target : null;
  let columnEl = targetEl?.closest?.("[data-column-key]") || null;
  if (!columnEl && Number.isFinite(ev?.clientX) && Number.isFinite(ev?.clientY)) {
    const hovered = document.elementFromPoint(ev.clientX, ev.clientY);
    columnEl = hovered?.closest?.("[data-column-key]") || null;
  }
  const key = columnEl?.getAttribute?.("data-column-key") || "";
  return isBonsColumnKey(key) ? key : null;
}

function getRepairStatusLabelForColumnKey(columnKey) {
  return getBonsColumnByKey(columnKey)?.statusLabel || null;
}

function repairStatusToneKey(statusLabel) {
  const s = normalizeText(statusLabel);
  if (s === "dossier ouvert") return "dossier-ouvert";
  if (s === "en test") return "en-test";
  if (s === "en cours") return "en-cours";
  if (s === "en attente de piece" || s === "en attente de pieces") return "en-attente-pieces";
  if (s === "termine" || s === "terminee") return "terminee";
  if (s === "remis au client") return "remis-client";
  if (s === "non reparable") return "non-reparable";
  if (s === "annule") return "annule";
  return "default";
}

function buildRepairStatusOptions(selectedStatusLabel) {
  const selected = normalizeText(selectedStatusLabel);
  const available = [...REPAIR_STATUS_OPTIONS];
  const hasSelected = available.some((label) => normalizeText(label) === selected);
  if (selected && !hasSelected) {
    available.unshift(selectedStatusLabel);
  }

  return available.map((label) => {
    const tone = repairStatusToneKey(label);
    const color = REPAIR_STATUS_TONE_COLORS[tone] || REPAIR_STATUS_TONE_COLORS.default;
    const isSelected = normalizeText(label) === selected ? " selected" : "";
    return `<option value="${escapeHtml(label)}" style="color:${color}"${isSelected}>${escapeHtml(label)}</option>`;
  }).join("");
}

function formatRepairCode(repair) {
  const seq = Number(repair?.podio_app_item_id);
  if (Number.isFinite(seq) && seq > 0) {
    return `REP${String(Math.trunc(seq)).padStart(4, "0")}`;
  }
  const fallback = Number(repair?.id);
  if (Number.isFinite(fallback) && fallback > 0) {
    return `REP${String(Math.trunc(fallback)).padStart(4, "0")}`;
  }
  return "REP0000";
}

function getRepairCardDeviceLabel(repair) {
  const manufacturer = cleanNullableText(repair?.manufacturer);
  const model = cleanNullableText(repair?.model);
  const deviceLabel = [manufacturer, model].filter(Boolean).join(" ").trim();
  if (deviceLabel) return deviceLabel;

  const fallbackTitle = cleanNullableText(repair?.title);
  if (!fallbackTitle) return "";

  const repCode = formatRepairCode(repair);
  const prefixedTitle = `${repCode} | `;
  if (fallbackTitle.startsWith(prefixedTitle)) {
    return fallbackTitle.slice(prefixedTitle.length).trim();
  }

  return fallbackTitle;
}

function formatDurationMinutesDisplay(totalMinutes) {
  if (totalMinutes == null || String(totalMinutes).trim() === "") return "";
  const raw = Number(totalMinutes);
  if (!Number.isFinite(raw) || raw <= 0) return "";
  const rounded = Math.trunc(raw);
  const hours = Math.floor(rounded / 60);
  const minutes = rounded % 60;
  if (hours <= 0) return `${minutes}min`;
  if (minutes === 0) return `${hours}h`;
  return `${hours}h${String(minutes).padStart(2, "0")}`;
}

function isYesValue(value) {
  const s = normalizeText(value);
  return s === "oui" || s === "yes" || s === "true" || s === "1";
}

function isPaidValue(value) {
  const s = normalizeText(value);
  return s === "paye" || s === "payee" || s === "paid" || isYesValue(value);
}

function isOverdueLabel(value) {
  return normalizeText(value).includes("retard");
}

function formatFactureCode(facture) {
  const numero = normalizeFactureCodeCandidate(facture?.numero);
  if (numero) return numero;
  const rawTitle = normalizeFactureCodeCandidate(facture?.raw_title);
  if (rawTitle) return rawTitle;

  const seq = Number(facture?.podio_app_item_id);
  const issueDate = String(facture?.date_emission ?? "").trim();
  if (/^\d{4}-\d{2}-\d{2}/.test(issueDate)) {
    const seqCode = buildFactureNumeroFromSeqAndDate(seq, issueDate);
    if (seqCode) return seqCode;
  }
  if (Number.isFinite(seq) && seq > 0) return `F${String(Math.trunc(seq)).padStart(5, "0")}`;
  const fallback = Number(facture?.id);
  if (Number.isFinite(fallback) && fallback > 0) return `F${String(Math.trunc(fallback)).padStart(5, "0")}`;
  return "";
}

function formatFactureNumber(facture) {
  return formatFactureCode(facture);
}

function buildStorageReferenceFromFilename(filenameValue) {
  const filename = cleanNullableText(filenameValue);
  if (!filename) return null;
  const normalized = filename.replace(/^\/+/, "").split("?")[0].split("#")[0];
  if (!normalized) return null;
  return {
    bucket: FACTURES_STORAGE_BUCKET,
    objectPath: normalized,
  };
}

async function getFacturePdfUrl(client, storageRef) {
  let bucket = String(storageRef?.bucket ?? "").trim();
  let objectPath = String(storageRef?.objectPath ?? "").trim();
  if (normalizeText(bucket) === "buckets" && objectPath.includes("/")) {
    const firstSlash = objectPath.indexOf("/");
    bucket = objectPath.slice(0, firstSlash).trim();
    objectPath = objectPath.slice(firstSlash + 1).trim();
  }
  if (!bucket || !objectPath) {
    throw new Error("Référence de fichier PDF invalide.");
  }

  const signed = await client.storage
    .from(bucket)
    .createSignedUrl(objectPath, 900, { download: false });
  if (!signed?.error && signed?.data?.signedUrl) {
    return signed.data.signedUrl;
  }

  // Fallback public URL (utile seulement si bucket public)
  try {
    const pub = client.storage
      .from(bucket)
      .getPublicUrl(objectPath);
    const publicUrl = cleanNullableText(pub?.data?.publicUrl);
    if (publicUrl) return publicUrl;
  } catch {
    // ignore
  }

  throw new Error(
    signed?.error?.message
      || `Impossible d'ouvrir le PDF (${bucket}/${objectPath}).`,
  );
}

function getPdfJsLib() {
  const lib = window?.pdfjsLib;
  if (!lib || typeof lib.getDocument !== "function") return null;
  if (lib.GlobalWorkerOptions && !lib.GlobalWorkerOptions.workerSrc) {
    lib.GlobalWorkerOptions.workerSrc = PDFJS_WORKER_SRC;
  }
  return lib;
}

function resetFacturePdfModalState() {
  facturePdfModalState.open = false;
  facturePdfModalState.title = "Apercu PDF";
  facturePdfModalState.pdfUrl = null;
  facturePdfModalState.page = 1;
  facturePdfModalState.totalPages = 0;
  facturePdfModalState.scale = 1.1;
  facturePdfModalState.renderToken += 1;
  facturePdfModalState.loading = false;

  if (facturePdfModalState.doc) {
    try {
      facturePdfModalState.doc.cleanup();
    } catch {
      // ignore
    }
    try {
      facturePdfModalState.doc.destroy();
    } catch {
      // ignore
    }
  }
  facturePdfModalState.doc = null;
}

function setFacturePdfModalStatus(message, isError = false) {
  const statusEl = $("facturePdfModalStatus");
  const canvasWrap = $("facturePdfCanvasWrap");
  if (!statusEl || !canvasWrap) return;
  statusEl.textContent = String(message ?? "");
  statusEl.hidden = false;
  statusEl.classList.toggle("is-error", Boolean(isError));
  canvasWrap.hidden = true;
}

function updateFacturePdfModalControls() {
  const pageInfo = $("facturePdfPageInfo");
  const prevBtn = $("facturePdfPrevBtn");
  const nextBtn = $("facturePdfNextBtn");
  const zoomOutBtn = $("facturePdfZoomOutBtn");
  const zoomInBtn = $("facturePdfZoomInBtn");

  const hasDoc = Boolean(facturePdfModalState.doc);
  const totalPages = Number(facturePdfModalState.totalPages || 0);
  const page = Number(facturePdfModalState.page || 1);
  const loading = Boolean(facturePdfModalState.loading);

  if (pageInfo) {
    pageInfo.textContent = totalPages > 0 ? `${page} / ${totalPages}` : "- / -";
  }
  if (prevBtn) prevBtn.disabled = !hasDoc || loading || page <= 1;
  if (nextBtn) nextBtn.disabled = !hasDoc || loading || page >= totalPages;
  if (zoomOutBtn) zoomOutBtn.disabled = !hasDoc || loading || facturePdfModalState.scale <= 0.55;
  if (zoomInBtn) zoomInBtn.disabled = !hasDoc || loading || facturePdfModalState.scale >= 2.8;
}

async function renderFacturePdfPage() {
  const doc = facturePdfModalState.doc;
  if (!doc) return;

  const canvas = $("facturePdfCanvas");
  const canvasWrap = $("facturePdfCanvasWrap");
  const statusEl = $("facturePdfModalStatus");
  if (!(canvas instanceof HTMLCanvasElement) || !canvasWrap || !statusEl) return;

  const safePage = Math.min(Math.max(1, Math.trunc(facturePdfModalState.page || 1)), Math.max(1, doc.numPages || 1));
  facturePdfModalState.page = safePage;
  facturePdfModalState.loading = true;
  updateFacturePdfModalControls();
  setFacturePdfModalStatus("Chargement de la page...");

  const token = ++facturePdfModalState.renderToken;
  try {
    const page = await doc.getPage(safePage);
    if (token !== facturePdfModalState.renderToken) return;

    const ratio = window.devicePixelRatio || 1;
    const logicalViewport = page.getViewport({ scale: facturePdfModalState.scale });
    const renderViewport = page.getViewport({ scale: facturePdfModalState.scale * ratio });
    const ctx = canvas.getContext("2d", { alpha: false });
    if (!ctx) throw new Error("Contexte canvas indisponible.");

    canvas.width = Math.max(1, Math.floor(renderViewport.width));
    canvas.height = Math.max(1, Math.floor(renderViewport.height));
    canvas.style.width = `${Math.max(1, Math.floor(logicalViewport.width))}px`;
    canvas.style.height = `${Math.max(1, Math.floor(logicalViewport.height))}px`;

    const renderTask = page.render({
      canvasContext: ctx,
      viewport: renderViewport,
    });
    await renderTask.promise;
    if (token !== facturePdfModalState.renderToken) return;

    canvasWrap.hidden = false;
    statusEl.hidden = true;
    statusEl.classList.remove("is-error");
  } catch (err) {
    if (token !== facturePdfModalState.renderToken) return;
    const msg = err?.message || "Impossible d'afficher ce PDF dans le modal.";
    setFacturePdfModalStatus(msg, true);
  } finally {
    if (token === facturePdfModalState.renderToken) {
      facturePdfModalState.loading = false;
      updateFacturePdfModalControls();
    }
  }
}

function openPdfInNewTab(pdfUrl) {
  const url = cleanNullableText(pdfUrl);
  if (!url) return false;
  const opened = window.open(url, "_blank", "noopener");
  if (opened) return true;
  window.location.href = url;
  return true;
}

function closeFacturePdfModal() {
  const modal = $("facturePdfModal");
  if (!modal) return;
  modal.hidden = true;
  resetFacturePdfModalState();
}

async function openFacturePdfModal(pdfUrl, titleText) {
  const modal = $("facturePdfModal");
  const title = $("facturePdfModalTitle");
  if (!modal || !title) return false;

  const pdfjs = getPdfJsLib();
  if (!pdfjs) return false;

  resetFacturePdfModalState();
  facturePdfModalState.open = true;
  facturePdfModalState.pdfUrl = pdfUrl;
  facturePdfModalState.title = cleanNullableText(titleText) || "Apercu PDF";
  title.textContent = facturePdfModalState.title;

  modal.hidden = false;
  setFacturePdfModalStatus("Chargement du document...");
  updateFacturePdfModalControls();

  try {
    const loadingTask = pdfjs.getDocument({ url: pdfUrl });
    const doc = await loadingTask.promise;
    facturePdfModalState.doc = doc;
    facturePdfModalState.totalPages = Number(doc?.numPages || 0);
    facturePdfModalState.page = 1;
    updateFacturePdfModalControls();
    await renderFacturePdfPage();
    return true;
  } catch (err) {
    closeFacturePdfModal();
    return false;
  }
}

async function openFacturePdfPreview(facture) {
  const client = state.authClient || ensureAuthClient();
  if (!client) throw new Error("Connexion Supabase indisponible.");

  const storageRef = buildStorageReferenceFromFilename(facture?.pdf_filename);

  if (!storageRef || !storageRef.objectPath) {
    throw new Error("Aucun fichier PDF associé à cette facture.");
  }

  const pdfUrl = await getFacturePdfUrl(client, storageRef);
  const factureNumber = cleanNullableText(formatFactureListNumber(facture)) || "Facture";
  const openedInModal = await openFacturePdfModal(pdfUrl, `Apercu PDF - ${factureNumber}`);
  if (!openedInModal) {
    const opened = openPdfInNewTab(pdfUrl);
    if (!opened) throw new Error("Impossible d'ouvrir ce PDF.");
  }
}

async function openDevisPdfPreview(doc) {
  const client = state.authClient || ensureAuthClient();
  if (!client) throw new Error("Connexion Supabase indisponible.");

  let workingDoc = doc && typeof doc === "object" ? doc : null;
  let storageRef = buildStorageReferenceFromFilename(workingDoc?.pdf_filename);
  if (!storageRef || !storageRef.objectPath) {
    if (isDeveloperModeEnabled()) {
      throw new Error("En mode développeur, le PDF n'est pas généré automatiquement.");
    }
    const docId = Number(workingDoc?.id);
    if (!Number.isInteger(docId) || docId <= 0) {
      throw new Error("Aucun fichier PDF associé à ce document.");
    }

    const { data, error } = await client.functions.invoke("estimations-submit", {
      body: {
        id: docId,
        numero: cleanNullableText(workingDoc?.numero),
        type_document: cleanNullableText(workingDoc?.type_document),
        contexte: cleanNullableText(workingDoc?.contexte),
        client_title: cleanNullableText(workingDoc?.client_title),
        date_emission: cleanNullableText(workingDoc?.date_emission),
        date_valid_until: cleanNullableText(workingDoc?.date_valid_until),
        etat_label: cleanNullableText(workingDoc?.etat_label),
        total: Number(workingDoc?.total ?? 0),
        description: cleanNullableText(workingDoc?.description),
        notes: cleanNullableText(workingDoc?.notes),
        generate_pdf: true,
      },
    });
    if (error) throw new Error(resolveEdgeFunctionInvokeError(error, "Erreur estimations-submit."));
    if (!data?.ok || !data?.item || typeof data.item !== "object") {
      throw new Error(data?.detail || data?.error || "Impossible de générer le PDF.");
    }

    workingDoc = data.item;
    const existing = Array.isArray(state.estimations) ? state.estimations : [];
    state.estimations = existing.map((row) => (
      String(row?.id) === String(workingDoc?.id) ? workingDoc : row
    ));
    renderFacturesList();
    storageRef = buildStorageReferenceFromFilename(workingDoc?.pdf_filename);
  }

  if (!storageRef || !storageRef.objectPath) {
    throw new Error("Aucun fichier PDF associé à ce document.");
  }

  const pdfUrl = await getFacturePdfUrl(client, storageRef);
  const numero = cleanNullableText(workingDoc?.numero) || "Document";
  const openedInModal = await openFacturePdfModal(pdfUrl, `Apercu PDF - ${numero}`);
  if (!openedInModal) {
    const opened = openPdfInNewTab(pdfUrl);
    if (!opened) throw new Error("Impossible d'ouvrir ce PDF.");
  }
}

function buildFactureIndex(factures) {
  const index = {};
  for (const facture of factures || []) {
    const reparations = Array.isArray(facture?.reparations) ? facture.reparations : [];
    for (const rep of reparations) {
      const refKind = normalizeText(rep?.kind);
      if (refKind === "project") continue;
      const itemId = rep?.item_id;
      if (itemId == null) continue;
      const key = String(itemId);
      if (!index[key]) index[key] = facture;
    }
  }
  return index;
}

function getClientReferenceId(client) {
  const dbId = Number(client?.id);
  if (Number.isInteger(dbId) && dbId > 0) return dbId;
  const legacyPodioItemId = Number(client?.podio_item_id);
  if (Number.isInteger(legacyPodioItemId) && legacyPodioItemId !== 0) return legacyPodioItemId;
  return null;
}

function buildClientIndex(clients) {
  const index = {};
  for (const client of clients || []) {
    const referenceId = getClientReferenceId(client);
    if (referenceId != null) index[String(referenceId)] = client;
    if (client?.podio_item_id != null) index[String(client.podio_item_id)] = client;
  }
  return index;
}

function formatClientPrimaryName(client) {
  const name = String(client?.name ?? "").trim();
  if (name) return name;
  const title = String(client?.title ?? "").trim();
  if (title) return title;
  const company = String(client?.company ?? "").trim();
  if (company) return company;
  return "Sans client";
}

function isPersonalRepair(repair) {
  return repair?.is_personal === true;
}

function isPersonalProject(project) {
  return project?.is_personal === true;
}

function getPersonalOwnerSelection() {
  const preferredPersonalClientId = 21;
  const preferredPersonalName = "Nicholas Rivet";
  const sessionUser = state.session?.user || null;
  const preferredName = cleanNullableText(state.profile?.display_name)
    || cleanNullableText(sessionUser?.user_metadata?.display_name)
    || cleanNullableText(sessionUser?.user_metadata?.full_name)
    || cleanNullableText(sessionUser?.user_metadata?.name)
    || (sessionUser?.email ? cleanNullableText(String(sessionUser.email).split("@")[0]) : null)
    || "Personnel";
  const preferredEmail = cleanNullableText(state.profile?.email)
    || cleanNullableText(sessionUser?.email);

  const explicitClient = (state.clients || []).find((client) => Number(client?.id) === preferredPersonalClientId) || null;
  const personalClients = (state.clients || []).filter((client) => normalizeClientType(client?.type) === "Personnel");
  const nicholasClient = personalClients.find((client) => normalizeText(formatClientPrimaryName(client)) === normalizeText(preferredPersonalName));
  const exactMatch = personalClients.find((client) => {
    const clientEmail = cleanNullableText(client?.email);
    if (preferredEmail && clientEmail && normalizeText(clientEmail) === normalizeText(preferredEmail)) return true;
    const clientName = formatClientPrimaryName(client);
    return preferredName && normalizeText(clientName) === normalizeText(preferredName);
  });
  const selectedClient = explicitClient || nicholasClient || exactMatch || personalClients[0] || null;
  if (selectedClient) {
    const itemId = Number(selectedClient?.podio_item_id);
    return {
      itemId: Number.isInteger(itemId) && itemId > 0 ? itemId : null,
      title: formatClientPrimaryName(selectedClient),
      company: cleanNullableText(selectedClient?.company) || "",
    };
  }

  return {
    itemId: null,
    title: preferredPersonalName,
    company: cleanNullableText(state.appConfig?.company_name) || "",
  };
}

function formatClientSecondaryLine(client) {
  const parts = [client?.company, client?.phone, client?.email]
    .map((v) => String(v ?? "").trim())
    .filter(Boolean);
  return parts.join(" | ");
}

function formatPhoneForDisplay(value) {
  const raw = String(value ?? "").trim();
  if (!raw) return "-";
  const digits = raw.replace(/\D/g, "");
  if (digits.length === 10) {
    return `${digits.slice(0, 3)}-${digits.slice(3, 6)}-${digits.slice(6, 10)}`;
  }
  return raw;
}

function normalizePhoneForInput(value) {
  const raw = String(value ?? "").trim();
  if (!raw) return "";
  const digits = raw.replace(/\D/g, "").slice(0, 10);
  if (!digits) return "";
  if (digits.length <= 3) return digits;
  if (digits.length <= 6) return `${digits.slice(0, 3)}-${digits.slice(3)}`;
  return `${digits.slice(0, 3)}-${digits.slice(3, 6)}-${digits.slice(6, 10)}`;
}

function extractAddressLinesFromRaw(rawAddress) {
  if (!rawAddress || typeof rawAddress !== "object") {
    return { street: null, cityProvince: null, postal: null };
  }

  const streetNumber = cleanNullableText(rawAddress.street_number);
  const streetName = cleanNullableText(rawAddress.street_name);
  const streetText = cleanNullableText(rawAddress.street);
  const street = streetText || [streetNumber, streetName].filter(Boolean).join(" ").trim() || null;
  const city = cleanNullableText(rawAddress.city);
  const state = cleanNullableText(rawAddress.state);
  const postal = cleanNullableText(rawAddress.postal_code);
  const cityProvince = [city, state].filter(Boolean).join(", ").trim() || city || state || null;

  return { street, cityProvince, postal };
}

function extractAddressLinesFromText(addressText) {
  const full = cleanNullableText(addressText);
  if (!full) return { street: null, cityProvince: null, postal: null };

  const parts = full.split(",").map((part) => part.trim()).filter(Boolean);
  if (!parts.length) return { street: null, cityProvince: null, postal: null };

  const countryRegex = /^(canada|united states|usa|etats-unis)$/i;
  const cleaned = countryRegex.test(parts[parts.length - 1]) ? parts.slice(0, -1) : parts;
  if (!cleaned.length) return { street: null, cityProvince: null, postal: null };

  const street = cleanNullableText(cleaned[0]);
  const postalRegex = /([A-Za-z]\d[A-Za-z]\s?\d[A-Za-z]\d|\d{5}(?:-\d{4})?)/;
  let postal = null;
  let state = null;
  let postalIndex = -1;

  for (let i = cleaned.length - 1; i >= 1; i -= 1) {
    const chunk = cleaned[i];
    const match = chunk.match(postalRegex);
    if (!match) continue;
    postal = cleanNullableText(match[1]?.toUpperCase());
    const statePart = cleanNullableText(chunk.replace(postalRegex, " ").replace(/\s+/g, " "));
    if (statePart) state = statePart;
    postalIndex = i;
    break;
  }

  if (!state && cleaned.length >= 3) {
    const maybeState = cleanNullableText(cleaned[cleaned.length - 1]);
    if (maybeState && (/^[A-Za-z]{2,3}$/.test(maybeState) || /quebec|ontario|alberta|british columbia|manitoba|saskatchewan|nova scotia|new brunswick|newfoundland|prince edward|pei|yukon|nunavut|northwest/i.test(maybeState))) {
      state = maybeState;
    }
  }

  const end = postalIndex >= 0 ? postalIndex : (state ? cleaned.length - 1 : cleaned.length);
  const city = cleanNullableText(cleaned.slice(1, Math.max(1, end)).pop());
  const cityProvince = [city, state].filter(Boolean).join(", ").trim() || city || state || null;

  return { street, cityProvince, postal };
}

function buildRepairClientInfoHtml(linkedClient, fallbackAddressText, company, phone, email) {
  const fromRaw = extractAddressLinesFromRaw(linkedClient?.raw?.address ?? null);
  const fromText = extractAddressLinesFromText(fallbackAddressText);

  const street = fromRaw.street || fromText.street;
  const cityProvince = fromRaw.cityProvince || fromText.cityProvince;
  const postal = fromRaw.postal || fromText.postal;

  const left = [street, cityProvince, postal].filter(Boolean);
  const right = [
    cleanNullableText(company),
    cleanNullableText(phone) ? formatPhoneForDisplay(phone) : null,
    cleanNullableText(email),
  ].filter(Boolean);

  if (!left.length && !right.length) {
    return `<div class="repair-client-quick-grid is-single-col"><p>-</p></div>`;
  }

  if (!left.length) {
    return `<div class="repair-client-quick-grid is-single-col">${right.map((value) => `<p>${escapeHtml(value)}</p>`).join("")}</div>`;
  }

  const rowCount = Math.max(left.length, right.length, 1);
  const cells = [];
  for (let i = 0; i < rowCount; i += 1) {
    const leftValue = left[i] ?? "";
    const rightValue = right[i] ?? "";
    cells.push(`<p${leftValue ? "" : ' class="is-empty"'}>${leftValue ? escapeHtml(leftValue) : "&nbsp;"}</p>`);
    cells.push(`<p${rightValue ? "" : ' class="is-empty"'}>${rightValue ? escapeHtml(rightValue) : "&nbsp;"}</p>`);
  }

  return `<div class="repair-client-quick-grid">${cells.join("")}</div>`;
}

function clearClientAddressSuggestions() {
  const list = $("clientAddressSuggestions");
  if (list) list.innerHTML = "";
  clientModalState.addressSuggestionsByValue = {};
  clientModalState.addressParts = null;
}

function resetClientAddressSuggestState() {
  if (clientModalState.addressSuggestTimer) {
    clearTimeout(clientModalState.addressSuggestTimer);
  }
  clientModalState.addressSuggestTimer = null;
  clientModalState.addressSuggestSeq = 0;
  clearClientAddressSuggestions();
}

function mapAddressSuggestions(items) {
  const byValue = {};
  for (const item of (items || [])) {
    const value = String(item?.display_name ?? "").trim();
    if (!value) continue;
    byValue[value] = item;
  }
  clientModalState.addressSuggestionsByValue = byValue;
}

function renderAddressSuggestions(items) {
  const list = $("clientAddressSuggestions");
  if (!list) return;
  list.innerHTML = (items || []).map((item) => {
    const value = String(item?.display_name ?? "").trim();
    if (!value) return "";
    return `<option value="${escapeHtml(value)}"></option>`;
  }).join("");
}

function addressPartsFromSuggestion(suggestion) {
  if (!suggestion || typeof suggestion !== "object") return null;

  const parts = {
    formatted: cleanNullableText(suggestion.formatted ?? suggestion.display_name),
    street: cleanNullableText(suggestion.street),
    city: cleanNullableText(suggestion.city),
    state: cleanNullableText(suggestion.state),
    postal_code: cleanNullableText(suggestion.postal_code),
    country: cleanNullableText(suggestion.country)
  };

  const hasData = Object.values(parts).some(Boolean);
  return hasData ? parts : null;
}

function syncClientAddressPartsFromInput() {
  const value = String($("clientAddressInput")?.value ?? "").trim();
  const selected = value ? clientModalState.addressSuggestionsByValue[value] : null;
  clientModalState.addressParts = addressPartsFromSuggestion(selected);
}

async function loadAddressSuggestions(query) {
  const q = String(query ?? "").trim();
  if (!clientModalState.editing) return;

  if (q.length < ADDRESS_SUGGEST_MIN_CHARS) {
    clearClientAddressSuggestions();
    return;
  }

  const client = state.authClient;
  if (!client) return;
  const seq = ++clientModalState.addressSuggestSeq;

  try {
    const { data, error } = await client.functions.invoke("address-suggest", {
      body: { query: q, limit: ADDRESS_SUGGEST_LIMIT }
    });
    if (seq !== clientModalState.addressSuggestSeq) return;
    if (error) throw error;

    const items = Array.isArray(data?.suggestions) ? data.suggestions : [];
    mapAddressSuggestions(items);
    renderAddressSuggestions(items);
    syncClientAddressPartsFromInput();
  } catch (_err) {
    if (seq !== clientModalState.addressSuggestSeq) return;
    clearClientAddressSuggestions();
  }
}

function scheduleAddressSuggestions(query) {
  if (clientModalState.addressSuggestTimer) {
    clearTimeout(clientModalState.addressSuggestTimer);
  }
  clientModalState.addressSuggestTimer = setTimeout(() => {
    loadAddressSuggestions(query);
  }, ADDRESS_SUGGEST_DEBOUNCE_MS);
}

function normalizeTextWithMap(value) {
  const raw = String(value ?? "");
  let normalized = "";
  const map = [];
  for (let i = 0; i < raw.length; i += 1) {
    const folded = raw[i]
      .normalize("NFD")
      .replace(/[\u0300-\u036f]/g, "")
      .toLowerCase();
    for (let j = 0; j < folded.length; j += 1) {
      normalized += folded[j];
      map.push(i);
    }
  }
  return { raw, normalized, map };
}

function mergeRanges(ranges) {
  if (!ranges.length) return [];
  const sorted = [...ranges].sort((a, b) => (a[0] - b[0]) || (a[1] - b[1]));
  const merged = [sorted[0]];
  for (let i = 1; i < sorted.length; i += 1) {
    const [start, end] = sorted[i];
    const last = merged[merged.length - 1];
    if (start <= last[1]) {
      last[1] = Math.max(last[1], end);
    } else {
      merged.push([start, end]);
    }
  }
  return merged;
}

function highlightSearchText(value, query) {
  const raw = String(value ?? "");
  if (!raw) return "-";
  const q = String(query ?? "").trim();
  if (!q) return escapeHtml(raw);

  const terms = [...new Set(q
    .split(/\s+/)
    .map((part) => normalizeText(part))
    .filter(Boolean))];
  if (!terms.length) return escapeHtml(raw);

  const folded = normalizeTextWithMap(raw);
  const ranges = [];
  for (const term of terms) {
    let startAt = 0;
    while (startAt < folded.normalized.length) {
      const idx = folded.normalized.indexOf(term, startAt);
      if (idx === -1) break;
      const rawStart = folded.map[idx];
      const rawEnd = (folded.map[idx + term.length - 1] ?? rawStart) + 1;
      if (Number.isInteger(rawStart) && Number.isInteger(rawEnd) && rawEnd > rawStart) {
        ranges.push([rawStart, rawEnd]);
      }
      startAt = idx + term.length;
    }
  }

  if (!ranges.length) return escapeHtml(raw);
  const merged = mergeRanges(ranges);
  let cursor = 0;
  let html = "";
  for (const [start, end] of merged) {
    html += escapeHtml(raw.slice(cursor, start));
    html += `<mark class="clients-highlight">${escapeHtml(raw.slice(start, end))}</mark>`;
    cursor = end;
  }
  html += escapeHtml(raw.slice(cursor));
  return html;
}

function matchesClientSearch(client, query) {
  const q = normalizeText(query);
  if (!q) return true;
  const haystack = [
    client?.title,
    client?.name,
    client?.company,
    client?.phone,
    client?.email,
    client?.address,
    client?.type,
    client?.podio_item_id,
    client?.podio_app_item_id
  ].map((v) => normalizeText(v)).join(" ");
  return haystack.includes(q);
}

function resolveClientForRepair(repair) {
  const clientItemId = repair?.client_item_id;
  if (clientItemId != null) {
    const byId = state.clientsByPodioItemId[String(clientItemId)];
    if (byId) return byId;
  }

  const wanted = normalizeText(repair?.client_title);
  if (!wanted) return null;
  return state.clients.find((client) => {
    return [
      client?.title,
      client?.name,
      client?.company
    ].some((field) => normalizeText(field) === wanted);
  }) || null;
}

function getFactureForRepair(repair) {
  const itemId = repair?.podio_item_id;
  if (itemId == null) return null;
  return state.factureByRepairItemId[String(itemId)] || null;
}

function extractProjectCode(text) {
  const m = String(text ?? "").match(/P\d{1,}/i);
  return m ? m[0].toUpperCase() : "";
}

function getFactureForProject(project) {
  if (!project) return null;
  const projectId = Number(project?.id);
  const projectCode = cleanNullableText(project?.numero)?.toUpperCase() || "";
  for (const facture of state.factures || []) {
    const factureContext = normalizeText(facture?.contexte);
    const refs = Array.isArray(facture?.reparations) ? facture.reparations : [];
    for (const ref of refs) {
      const refKind = normalizeText(ref?.kind);
      const refItemId = Number(ref?.item_id);
      const refCode = cleanNullableText(ref?.code)?.toUpperCase()
        || extractProjectCode(ref?.title);
      const isProjectRef = refKind === "project" || refKind === "projet" || (!refKind && factureContext === "projet");
      if (!isProjectRef) continue;
      if (Number.isInteger(projectId) && projectId > 0 && Number.isInteger(refItemId) && refItemId === projectId) {
        return facture;
      }
      if (projectCode && refCode === projectCode) return facture;
    }
  }
  return null;
}

function roundFactureMoney(value) {
  const n = Number(value);
  if (!Number.isFinite(n)) return 0;
  return Math.round(n * 100) / 100;
}

function isProjectLaborDescription(value) {
  const normalized = normalizeText(value)
    .replace(/[’'`]/g, "")
    .replace(/\s+/g, " ")
    .trim();
  return normalized.includes("main doeuvre") || normalized.includes("main d oeuvre");
}

function ensureProjectLaborCostLines(linesValue, billableHoursValue, laborRateValue, { includeLabor = true } = {}) {
  const lines = normalizeProjectCostLines(linesValue);
  if (!includeLabor) {
    return lines.filter((line) => String(line?.source_kind || "").trim() !== "project_labor" && !isProjectLaborDescription(line?.description));
  }
  const billableHoursRaw = Number(billableHoursValue);
  const laborRateRaw = Number(laborRateValue);
  const billableHours = Number.isFinite(billableHoursRaw) && billableHoursRaw > 0
    ? roundFactureMoney(billableHoursRaw)
    : 0;
  const laborRate = Number.isFinite(laborRateRaw) && laborRateRaw >= 0
    ? roundFactureMoney(laborRateRaw)
    : 0;

  let laborIndex = lines.findIndex((line) => String(line?.source_kind || "").trim() === "project_labor");
  if (laborIndex < 0) {
    laborIndex = lines.findIndex((line) => isProjectLaborDescription(line?.description));
  }

  if (laborIndex < 0) {
    const quantity = billableHours > 0 ? billableHours : 1;
    lines.unshift({
      description: "Main d'oeuvre",
      quantity,
      unit_price: laborRate,
      amount: roundFactureMoney(quantity * laborRate),
      currency: PROJECT_COST_DEFAULT_CURRENCY,
      exchange_rate_to_cad: 1,
      expense_date: "",
      labor_rate_key: null,
      source_kind: "project_labor",
    });
    return lines;
  }

  const laborLine = { ...lines[laborIndex] };
  const qtyRaw = Number(laborLine.quantity);
  const unitRaw = Number(laborLine.unit_price);
  const quantity = Number.isFinite(qtyRaw) && qtyRaw > 0
    ? roundFactureMoney(qtyRaw)
    : (billableHours > 0 ? billableHours : 1);
  const unitPrice = Number.isFinite(unitRaw) && unitRaw >= 0
    ? roundFactureMoney(unitRaw)
    : laborRate;
  laborLine.description = "Main d'oeuvre";
  laborLine.quantity = quantity;
  laborLine.unit_price = unitPrice;
  laborLine.amount = roundFactureMoney(quantity * unitPrice);
  laborLine.currency = PROJECT_COST_DEFAULT_CURRENCY;
  laborLine.exchange_rate_to_cad = 1;
  laborLine.expense_date = "";
  laborLine.source_kind = "project_labor";
  lines[laborIndex] = laborLine;
  if (laborIndex > 0) {
    lines.splice(laborIndex, 1);
    lines.unshift(laborLine);
  }
  return lines;
}

function getFacturePrefillRepairRow(repair) {
  if (!repair) return null;
  const podioItemId = Number(repair?.podio_item_id);
  if (!Number.isInteger(podioItemId) || podioItemId === 0) return null;
  const linkedClient = resolveClientForRepair(repair);
  const clientItemIdRaw = Number(repair?.client_item_id ?? getClientReferenceId(linkedClient));
  const clientItemId = Number.isInteger(clientItemIdRaw) && clientItemIdRaw !== 0 ? clientItemIdRaw : null;
  const manufacturer = cleanNullableText(repair?.manufacturer) || "";
  const model = cleanNullableText(repair?.model) || "";
  const device = [manufacturer, model].filter(Boolean).join(" ").trim();
  return {
    ...repair,
    podioItemId,
    repCode: formatRepairCode(repair),
    clientItemId,
    clientDisplay: linkedClient ? formatClientPrimaryName(linkedClient) : (cleanNullableText(repair?.client_title) || "Sans client"),
    deviceLabel: device || "Appareil",
  };
}

function buildFactureLinesFromRepair(repair, { decorate = false } = {}) {
  const prefillRepair = getFacturePrefillRepairRow(repair);
  const fallbackLaborRate = getConfiguredFactureLaborRate();
  const normalizedLines = ensureRepairInvoiceLines(repair?.lines, {
    durationMinutes: repair?.duration_minutes,
    partsCost: repair?.parts_cost,
    defaultLaborRate: fallbackLaborRate,
    hourlyRateSnapshot: repair?.hourly_rate_snapshot,
    fallbackLaborQuantity: 0,
    includeLabor: !isPersonalRepair(repair),
  });
  const out = normalizedLines
    .map((line) => {
      const quantity = roundFactureMoney(Number(line.quantity || 0));
      const unitPrice = getProjectCostLineCadUnitPrice(line);
      const amount = getProjectCostLineCadAmount(line);
      const laborLine = isRepairLaborLine(line);
      let description = cleanNullableText(line.description) || (laborLine ? "Main d'oeuvre" : "Pièces");
      if (decorate && prefillRepair) {
        if (laborLine) {
          description = `Main d'oeuvre (${prefillRepair.repCode})`;
        } else if (normalizeText(description) === "pieces" || normalizeText(description) === "piece") {
          description = `Pièces (${prefillRepair.deviceLabel})`;
        } else {
          description = `${description} (${prefillRepair.repCode})`;
        }
      }
      return {
        description,
        quantity,
        unit_price: unitPrice,
        amount,
        labor_rate_key: cleanNullableText(line.labor_rate_key),
        source_kind: cleanNullableText(line.source_kind) || (laborLine ? "repair_labor" : "repair_parts"),
      };
    })
    .filter((line) => cleanNullableText(line.description) || Number(line.amount) > 0);

  if (out.length) return out;
  if (!prefillRepair) return [];
  return [{
    description: `Réparation (${prefillRepair.repCode})`,
    quantity: 1,
    unit_price: 0,
    amount: 0,
    labor_rate_key: null,
    source_kind: "repair_misc",
  }];
}

function buildDefaultFacturePayloadForRepair(repair) {
  if (isPersonalRepair(repair)) {
    throw new Error("Impossible de créer la facture : cette réparation est personnelle.");
  }
  const prefillRepair = getFacturePrefillRepairRow(repair);
  if (!prefillRepair) {
    throw new Error("Impossible de créer la facture : identifiant du bon invalide.");
  }

  const linkedClient = resolveClientForRepair(repair);
  const clientItemId = Number(prefillRepair.clientItemId);
  if (!Number.isInteger(clientItemId) || clientItemId === 0) {
    throw new Error("Impossible de créer la facture : aucun client lié à ce bon.");
  }

  const selectedClientRow = state.clientsByPodioItemId[String(clientItemId)] || linkedClient || null;
  const selectedClientTitle = selectedClientRow
    ? formatClientPrimaryName(selectedClientRow)
    : (cleanNullableText(prefillRepair.clientDisplay) || cleanNullableText(repair?.client_title) || "Sans client");

  const rawAddressLines = extractAddressLinesFromRaw(selectedClientRow?.raw?.address ?? null);
  const textAddressLines = extractAddressLinesFromText(selectedClientRow?.address ?? null);
  const clientStreet = rawAddressLines.street || textAddressLines.street || cleanNullableText(selectedClientRow?.address);
  const clientCityProvince = rawAddressLines.cityProvince || textAddressLines.cityProvince || null;
  const clientPostalCode = rawAddressLines.postal || textAddressLines.postal || null;
  const clientSnapshot = selectedClientRow
    ? {
        name: formatClientPrimaryName(selectedClientRow),
        type: normalizeClientType(selectedClientRow?.type) || cleanNullableText(selectedClientRow?.type),
        company: cleanNullableText(selectedClientRow?.company),
        phone: cleanNullableText(selectedClientRow?.phone),
        email: cleanNullableText(selectedClientRow?.email),
        address_line1: cleanNullableText(clientStreet),
        city_province: cleanNullableText(clientCityProvince),
        postal_code: cleanNullableText(clientPostalCode),
      }
    : null;

  const lines = buildFactureLinesFromRepair(repair, { decorate: true });

  const subtotal = roundFactureMoney(lines.reduce((sum, line) => sum + Number(line.amount || 0), 0));
  const taxRates = getConfiguredFactureTaxRates();
  const applyTaxes = false;
  const tps = applyTaxes ? roundFactureMoney(subtotal * taxRates.tps) : 0;
  const tvq = applyTaxes ? roundFactureMoney(subtotal * taxRates.tvq) : 0;
  const total = roundFactureMoney(subtotal + tps + tvq);
  const issueDate = formatDateYmd(new Date());
  const dueDate = addDaysToYmd(issueDate, FACTURE_DUE_DAYS) || null;

  return {
    client_item_id: clientItemId,
    client_title: selectedClientTitle || null,
    client_snapshot: clientSnapshot,
    date_emission: issueDate,
    date_echeance: dueDate,
    subtotal,
    tps,
    tvq,
    total,
    status_label: "Émise",
    tax_year_label: String(new Date(`${issueDate}T00:00:00`).getFullYear()),
    reparations: [{
      kind: "repair",
      item_id: Number(prefillRepair.podioItemId),
      title: `${prefillRepair.repCode} | ${prefillRepair.deviceLabel}`,
    }],
    device_label: cleanNullableText(prefillRepair.deviceLabel) || null,
    lines,
    desc_problem: cleanNullableText(htmlToText(repair?.desc_problem)),
    desc_done: cleanNullableText(htmlToText(repair?.desc_done)),
    company_snapshot: getCompanySnapshotFromConfig(),
  };
}

function showRepairFactureCreateDialog(repair) {
  if (repairFactureCreateDialogPromise) return repairFactureCreateDialogPromise;

  const repCode = formatRepairCode(repair);
  const linkedClient = resolveClientForRepair(repair);
  const clientName = linkedClient
    ? formatClientPrimaryName(linkedClient)
    : (cleanNullableText(repair?.client_title) || "ce client");

  repairFactureCreateDialogPromise = new Promise((resolve) => {
    const previousActive = document.activeElement instanceof HTMLElement ? document.activeElement : null;
    const root = document.createElement("div");
    root.id = "repairFactureCreateConfirm";
    root.className = "repair-confirm-overlay";
    root.innerHTML = `
      <div class="repair-confirm-panel" role="dialog" aria-modal="true" aria-labelledby="repairFactureCreateTitle">
        <h4 id="repairFactureCreateTitle" class="repair-confirm-title">Créer une facture</h4>
        <p class="repair-confirm-text">
          Voulez-vous créer une facture pour ${escapeHtml(clientName)} pour la réparation ${escapeHtml(repCode)} ?
        </p>
        <div class="repair-confirm-actions">
          <button type="button" class="repair-confirm-btn repair-confirm-btn-primary" data-choice="create">Oui</button>
          <button type="button" class="repair-confirm-btn" data-choice="create_edit">Oui et éditer</button>
          <button type="button" class="repair-confirm-btn" data-choice="cancel">Annuler</button>
        </div>
      </div>
    `;

    const finalize = (choice) => {
      document.removeEventListener("keydown", onKeyDown, true);
      root.remove();
      repairFactureCreateDialogPromise = null;
      if (previousActive && document.contains(previousActive)) previousActive.focus();
      resolve(choice);
    };

    const onKeyDown = (ev) => {
      if (ev.key !== "Escape") return;
      ev.preventDefault();
      ev.stopPropagation();
      finalize("cancel");
    };

    root.addEventListener("click", (ev) => {
      const btn = ev.target?.closest?.("[data-choice]");
      if (btn) {
        finalize(btn.getAttribute("data-choice"));
        return;
      }
      if (ev.target === root) finalize("cancel");
    });

    document.addEventListener("keydown", onKeyDown, true);
    document.body.appendChild(root);
    root.querySelector("[data-choice='cancel']")?.focus();
  });

  return repairFactureCreateDialogPromise;
}

function openRepairPartsSuppliersDialog() {
  if (repairPartsSuppliersDialogOpen) return;
  repairPartsSuppliersDialogOpen = true;

  const previousActive = document.activeElement instanceof HTMLElement ? document.activeElement : null;
  const suppliers = getRepairPartsSupplierLinks();
  const cardsHtml = suppliers.map((supplier) => {
    const logoSrc = cleanNullableText(supplier.faviconUrl) || "data:image/gif;base64,R0lGODlhAQABAAAAACw=";
    const note = cleanNullableText(supplier.note);
    const noteInfo = note
      ? `
        <span class="configuration-label-with-info repair-suppliers-note-wrap">
          <span class="configuration-info-btn repair-suppliers-note-btn" aria-hidden="true">i</span>
          <span class="configuration-info-tooltip repair-suppliers-note-tooltip" role="tooltip">${escapeHtml(note)}</span>
        </span>
      `
      : "";
    const disabled = !supplier.href;
    if (disabled) {
      return `
        <button type="button" class="repair-suppliers-card is-disabled" disabled>
          <span class="repair-suppliers-card-brand">
            <img class="repair-suppliers-card-logo" src="${escapeHtml(logoSrc)}" alt="" loading="lazy" decoding="async">
            <span class="repair-suppliers-card-name-wrap">
              <span class="repair-suppliers-card-name">${escapeHtml(supplier.name)}</span>
              ${noteInfo}
            </span>
          </span>
          <span class="repair-suppliers-card-link">Lien non défini</span>
        </button>
      `;
    }
    return `
      <a class="repair-suppliers-card" href="${escapeHtml(supplier.href)}" target="_blank" rel="noopener noreferrer">
        <span class="repair-suppliers-card-brand">
          <img class="repair-suppliers-card-logo" src="${escapeHtml(logoSrc)}" alt="" loading="lazy" decoding="async">
          <span class="repair-suppliers-card-name-wrap">
            <span class="repair-suppliers-card-name">${escapeHtml(supplier.name)}</span>
            ${noteInfo}
          </span>
        </span>
        <span class="repair-suppliers-card-link">${escapeHtml(supplier.domain || supplier.href)}</span>
      </a>
    `;
  }).join("");
  const gridContent = cardsHtml || `<p class="repair-suppliers-empty">Aucun lien d'achat configuré.</p>`;

  const root = document.createElement("div");
  root.id = "repairPartsSuppliersDialog";
  root.className = "repair-confirm-overlay repair-suppliers-overlay";
  root.innerHTML = `
    <div class="repair-confirm-panel repair-suppliers-panel" role="dialog" aria-modal="true" aria-labelledby="repairSuppliersTitle">
      <button type="button" class="repair-suppliers-close" data-suppliers-close="1" aria-label="Fermer">×</button>
      <h4 id="repairSuppliersTitle" class="repair-confirm-title">Liens d'achats de pièces</h4>
      <div class="repair-suppliers-grid">${gridContent}</div>
    </div>
  `;

  const closeDialog = () => {
    document.removeEventListener("keydown", onKeyDown, true);
    root.remove();
    repairPartsSuppliersDialogOpen = false;
    if (previousActive && document.contains(previousActive)) previousActive.focus();
  };

  const onKeyDown = (ev) => {
    if (ev.key !== "Escape") return;
    ev.preventDefault();
    ev.stopPropagation();
    closeDialog();
  };

  root.addEventListener("click", (ev) => {
    const closeBtn = ev.target?.closest?.("[data-suppliers-close]");
    if (closeBtn) {
      closeDialog();
      return;
    }
    if (ev.target === root) closeDialog();
  });

  document.addEventListener("keydown", onKeyDown, true);
  document.body.appendChild(root);
  root.querySelector("[data-suppliers-close]")?.focus();
}

async function promptRepairFactureCreation(repair) {
  if (!repair) return;
  if (isPersonalRepair(repair)) {
    window.alert("Aucune facture n'est émise pour une réparation personnelle.");
    return;
  }
  const existingFacture = getFactureForRepair(repair);
  if (existingFacture) {
    const factureNumber = formatFactureCode(existingFacture) || "cette facture";
    window.alert(`Une facture existe déjà pour ce bon (${factureNumber}).`);
    return;
  }

  const choice = await showRepairFactureCreateDialog(repair);
  if (choice === "cancel") return;

  const prefillRepairItemId = Number(repair?.podio_item_id);
  if (!Number.isInteger(prefillRepairItemId) || prefillRepairItemId === 0) {
    window.alert("Impossible de créer la facture : identifiant du bon invalide.");
    return;
  }

  if (choice === "create_edit") {
    openCreateFactureModal({ prefillRepairItemId });
    return;
  }

  if (choice !== "create") return;

  try {
    const payload = buildDefaultFacturePayloadForRepair(repair);
    await createFactureRecord(payload);
    renderFacturesList();
    renderBonsBoard();
    renderAccueilCards();
  } catch (err) {
    window.alert(err?.message || "Impossible de créer la facture.");
  }
}

function showProjectFactureCreateDialog(project) {
  if (projectFactureCreateDialogPromise) return projectFactureCreateDialogPromise;
  const projectCode = (cleanNullableText(project?.numero) || "P0000").toUpperCase();
  const linkedClient = state.clientsByPodioItemId?.[String(project?.client_item_id ?? "")] || null;
  const clientName = linkedClient
    ? formatClientPrimaryName(linkedClient)
    : (cleanNullableText(project?.client_title) || "ce client");

  projectFactureCreateDialogPromise = new Promise((resolve) => {
    const previousActive = document.activeElement instanceof HTMLElement ? document.activeElement : null;
    const root = document.createElement("div");
    root.id = "projectFactureCreateConfirm";
    root.className = "repair-confirm-overlay";
    root.innerHTML = `
      <div class="repair-confirm-panel" role="dialog" aria-modal="true" aria-labelledby="projectFactureCreateTitle">
        <h4 id="projectFactureCreateTitle" class="repair-confirm-title">Émettre une facture</h4>
        <p class="repair-confirm-text">
          Voulez-vous émettre une facture pour ${escapeHtml(clientName)} pour le projet ${escapeHtml(projectCode)} ?
        </p>
        <div class="repair-confirm-actions">
          <button type="button" class="repair-confirm-btn repair-confirm-btn-primary" data-choice="create">Oui</button>
          <button type="button" class="repair-confirm-btn" data-choice="create_edit">Oui et éditer</button>
          <button type="button" class="repair-confirm-btn" data-choice="cancel">Annuler</button>
        </div>
      </div>
    `;

    const finalize = (choice) => {
      document.removeEventListener("keydown", onKeyDown, true);
      root.remove();
      projectFactureCreateDialogPromise = null;
      if (previousActive && document.contains(previousActive)) previousActive.focus();
      resolve(choice);
    };

    const onKeyDown = (ev) => {
      if (ev.key !== "Escape") return;
      ev.preventDefault();
      ev.stopPropagation();
      finalize("cancel");
    };

    root.addEventListener("click", (ev) => {
      const btn = ev.target?.closest?.("[data-choice]");
      if (btn) {
        finalize(btn.getAttribute("data-choice"));
        return;
      }
      if (ev.target === root) finalize("cancel");
    });

    document.addEventListener("keydown", onKeyDown, true);
    document.body.appendChild(root);
    root.querySelector("[data-choice='cancel']")?.focus();
  });

  return projectFactureCreateDialogPromise;
}

function buildDefaultFacturePayloadForProject(project) {
  if (isPersonalProject(project)) {
    throw new Error("Impossible de créer la facture : ce projet est personnel.");
  }
  const projectId = Number(project?.id);
  if (!Number.isInteger(projectId) || projectId <= 0) {
    throw new Error("Impossible de créer la facture : identifiant du projet invalide.");
  }
  const projectCode = (cleanNullableText(project?.numero) || `P${String(projectId).padStart(4, "0")}`).toUpperCase();
  const projectTitle = cleanNullableText(project?.title) || projectCode;
  const linkedClient = state.clientsByPodioItemId?.[String(project?.client_item_id ?? "")] || null;
  const clientItemIdRaw = Number(project?.client_item_id ?? getClientReferenceId(linkedClient));
  const clientItemId = Number.isInteger(clientItemIdRaw) && clientItemIdRaw > 0 ? clientItemIdRaw : null;
  if (!clientItemId) {
    throw new Error("Impossible de créer la facture : aucun client lié à ce projet.");
  }

  const selectedClientRow = state.clientsByPodioItemId[String(clientItemId)] || linkedClient || null;
  const selectedClientTitle = selectedClientRow
    ? formatClientPrimaryName(selectedClientRow)
    : (cleanNullableText(project?.client_title) || "Sans client");
  const rawAddressLines = extractAddressLinesFromRaw(selectedClientRow?.raw?.address ?? null);
  const textAddressLines = extractAddressLinesFromText(selectedClientRow?.address ?? null);
  const clientStreet = rawAddressLines.street || textAddressLines.street || cleanNullableText(selectedClientRow?.address);
  const clientCityProvince = rawAddressLines.cityProvince || textAddressLines.cityProvince || null;
  const clientPostalCode = rawAddressLines.postal || textAddressLines.postal || null;
  const clientSnapshot = selectedClientRow
    ? {
      name: formatClientPrimaryName(selectedClientRow),
      type: normalizeClientType(selectedClientRow?.type) || cleanNullableText(selectedClientRow?.type),
      company: cleanNullableText(selectedClientRow?.company),
      phone: cleanNullableText(selectedClientRow?.phone),
      email: cleanNullableText(selectedClientRow?.email),
      address_line1: cleanNullableText(clientStreet),
      city_province: cleanNullableText(clientCityProvince),
      postal_code: cleanNullableText(clientPostalCode),
    }
    : null;

  const sourceLines = ensureProjectLaborCostLines(
    project?.cost_lines,
    project?.billable_hours,
    getConfiguredFactureLaborRate(),
    { includeLabor: !isPersonalProject(project) },
  );
  const lines = sourceLines
    .map((line) => {
      const quantity = Math.max(0, roundFactureMoney(Number(line?.quantity ?? 0)));
      const unitPriceCad = Math.max(0, getProjectCostLineCadUnitPrice(line));
      const amount = Math.max(0, getProjectCostLineCadAmount(line));
      const description = cleanNullableText(line?.description);
      if (!description && amount <= 0) return null;
      return {
        description: description || "Ligne de projet",
        quantity: quantity > 0 ? quantity : 1,
        unit_price: unitPriceCad,
        amount: quantity > 0 ? amount : roundFactureMoney(unitPriceCad),
      };
    })
    .filter(Boolean);
  if (!lines.length) {
    lines.push({
      description: projectTitle,
      quantity: 1,
      unit_price: 0,
      amount: 0,
    });
  }

  const subtotal = roundFactureMoney(lines.reduce((sum, line) => sum + Number(line.amount || 0), 0));
  const tps = 0;
  const tvq = 0;
  const total = roundFactureMoney(subtotal + tps + tvq);
  const issueDate = formatDateYmd(new Date());
  const dueDate = addDaysToYmd(issueDate, FACTURE_DUE_DAYS) || null;

  return {
    client_item_id: clientItemId,
    client_title: selectedClientTitle || null,
    client_snapshot: clientSnapshot,
    date_emission: issueDate,
    date_echeance: dueDate,
    subtotal,
    tps,
    tvq,
    total,
    status_label: "Émise",
    tax_year_label: String(new Date(`${issueDate}T00:00:00`).getFullYear()),
    reparations: [{
      kind: "project",
      item_id: projectId,
      code: projectCode,
      title: `${projectCode} | ${projectTitle}`,
    }],
    device_label: projectTitle,
    lines,
    desc_problem: cleanNullableText(project?.description) || cleanNullableText(project?.title),
    desc_done: null,
    company_snapshot: getCompanySnapshotFromConfig(),
    contexte: "projet",
  };
}

async function promptProjectFactureCreation(project) {
  if (!project) return;
  if (isPersonalProject(project)) {
    window.alert("Aucune facture n'est émise pour un projet personnel.");
    return;
  }
  const existingFacture = getFactureForProject(project);
  if (existingFacture) {
    const factureNumber = formatFactureCode(existingFacture) || "cette facture";
    window.alert(`Une facture existe déjà pour ce projet (${factureNumber}).`);
    return;
  }

  const choice = await showProjectFactureCreateDialog(project);
  if (choice === "cancel") return;

  const projectId = Number(project?.id);
  if (!Number.isInteger(projectId) || projectId <= 0) {
    window.alert("Impossible de créer la facture : identifiant du projet invalide.");
    return;
  }

  if (choice === "create_edit") {
    const projectModal = $("projectModal");
    const isProjectModalOpen = Boolean(projectModal && !projectModal.hidden);
    const isDirty = isProjectModalOpen && typeof projectModal.__projectIsDirty === "function"
      ? Boolean(projectModal.__projectIsDirty())
      : false;
    if (isDirty) {
      const proceedChoice = await showProjectProceedToFactureDialog();
      if (proceedChoice !== "save_continue") return;
      const saveFn = typeof projectModal?.__projectSave === "function" ? projectModal.__projectSave : null;
      if (typeof saveFn !== "function") return;
      const saved = await saveFn();
      if (!saved) return;
    }
    const closed = await closeProjectModal();
    if (!closed) return;
    openCreateFactureModal({ prefillProjectId: projectId, prefillContext: "projet" });
    return;
  }
  if (choice !== "create") return;

  try {
    const payload = buildDefaultFacturePayloadForProject(project);
    await createFactureRecord(payload);
    renderFacturesList();
    renderProjectsList();
    renderAccueilCards();
  } catch (err) {
    window.alert(err?.message || "Impossible de créer la facture.");
  }
}

function formatFactureListNumber(facture) {
  const code = formatFactureCode(facture);
  if (code) return code;
  return "-";
}

function factureSortTime(facture) {
  const source = facture?.date_emission || facture?.created_at || facture?.updated_at || "";
  if (/^\d{4}-\d{2}-\d{2}/.test(source)) {
    const t = Date.parse(`${source.slice(0, 10)}T00:00:00Z`);
    return Number.isFinite(t) ? t : 0;
  }
  const t = Date.parse(source);
  return Number.isFinite(t) ? t : 0;
}

function parseDateOnly(value) {
  const raw = String(value ?? "").trim();
  if (!raw) return null;
  const datePart = /^\d{4}-\d{2}-\d{2}/.test(raw) ? raw.slice(0, 10) : raw;
  const d = new Date(`${datePart}T00:00:00`);
  if (Number.isNaN(d.getTime())) return null;
  d.setHours(0, 0, 0, 0);
  return d;
}

function resolveFactureYear(facture) {
  const candidates = [
    facture?.date_emission,
    facture?.created_at,
    facture?.updated_at,
  ];
  for (const value of candidates) {
    const raw = String(value ?? "").trim();
    if (!raw) continue;
    if (/^\d{4}-\d{2}-\d{2}/.test(raw)) {
      const y = Number(raw.slice(0, 4));
      if (Number.isInteger(y) && y > 0) return y;
      continue;
    }
    const d = parseDateTimeSafe(raw);
    if (d) return d.getFullYear();
  }
  return null;
}

function matchFactureYearFilter(facture, filterKey) {
  const key = String(filterKey ?? "all").trim() || "all";
  if (key === "all") return true;
  const year = resolveFactureYear(facture);
  const currentYear = new Date().getFullYear();
  if (key === "current") return year === currentYear;
  if (key === "previous") return Number.isInteger(year) && year < currentYear;
  if (key.startsWith("year:")) {
    const wanted = Number(key.slice(5));
    if (!Number.isInteger(wanted)) return true;
    return year === wanted;
  }
  return true;
}

function matchFactureStatusFilter(facture, filterKey) {
  const key = String(filterKey ?? "all").trim() || "all";
  if (key === "all") return true;
  const isPaid = isPaidValue(facture?.etat_label);
  if (key === "paid") return isPaid;
  if (key === "unpaid") return !isPaid;
  if (key === "overdue") return getFactureEtatDisplay(facture).className === "is-overdue";
  return true;
}

function matchFactureSearch(facture, query, repairsByItemId) {
  const q = normalizeText(query);
  if (!q) return true;

  const noFacture = formatFactureListNumber(facture);
  const client = facture?.client_title || facture?.raw_title || "";
  const dateEmission = formatDateCell(facture?.date_emission || facture?.created_at);
  const dateEcheance = formatDateCell(facture?.date_echeance);
  const total = facture?.total == null ? "" : String(facture.total);
  const repCodes = formatFactureRepCodes(facture, repairsByItemId);
  const etat = getFactureEtatDisplay(facture).label;
  const year = resolveFactureYear(facture);

  const haystack = normalizeText([
    noFacture,
    client,
    dateEmission,
    dateEcheance,
    total,
    repCodes,
    etat,
    year == null ? "" : String(year),
  ].join(" "));

  return haystack.includes(q);
}

function getFactureEtatDisplay(facture) {
  const rawLabel = facture?.etat_label || "";
  const label = String(rawLabel).trim() || "-";
  const norm = normalizeText(label);
  const isPaid = isPaidValue(label);
  const isIssued = norm.includes("emise");
  const alreadyOverdue = isOverdueLabel(label);

  if (isPaid) return { label: "Payée", className: "is-paid" };
  if (alreadyOverdue) return { label: "En retard", className: "is-overdue" };

  if (isIssued) {
    const due = parseDateOnly(facture?.date_echeance);
    if (due) {
      const today = new Date();
      today.setHours(0, 0, 0, 0);
      if (today.getTime() > due.getTime()) {
        return { label: "En retard", className: "is-overdue" };
      }
    }
  }

  return { label, className: "is-open" };
}

function extractRepCode(text) {
  const m = String(text ?? "").match(/REP\d{1,}/i);
  return m ? m[0].toUpperCase() : "";
}

function formatFactureRepCodes(facture, repairsByItemId) {
  const reparations = Array.isArray(facture?.reparations) ? facture.reparations : [];
  if (!reparations.length) return "-";

  const codes = reparations.map((rep) => {
    const refKind = normalizeText(rep?.kind);
    if (refKind === "project") {
      const code = cleanNullableText(rep?.code)?.toUpperCase() || extractProjectCode(rep?.title);
      if (code) return code;
      const itemId = rep?.item_id;
      return itemId != null ? `P#${itemId}` : "";
    }
    const itemId = rep?.item_id;
    const repair = itemId != null ? repairsByItemId.get(String(itemId)) : null;
    if (repair) return formatRepairCode(repair);
    const fromTitle = extractRepCode(rep?.title);
    if (fromTitle) return fromTitle;
    return itemId != null ? `#${itemId}` : "";
  }).filter(Boolean);

  if (!codes.length) return "-";
  return [...new Set(codes)].join(", ");
}

function setListMessage(listId, message) {
  const list = $(listId);
  if (!list) return;
  list.innerHTML = `<li class="data-empty">${escapeHtml(message)}</li>`;
}

function renderRepairsBucket(listId, metaId, repairs, metaFn, emptyText) {
  const list = $(listId);
  const meta = $(metaId);
  if (!list || !meta) return;

  meta.textContent = metaFn(repairs.length);
  if (!repairs.length) {
    setListMessage(listId, emptyText);
    return;
  }

  list.innerHTML = repairs
    .map((repair) => {
      const title = repair.title || `Réparation #${repair.id}`;
      const sub = [repair.client_title || "Sans client", repair.status_label || "", formatDateTime(repair.updated_at || repair.podio_created_at || repair.created_at)]
        .filter(Boolean)
        .join(" - ");
      return `<li class="data-item"><div class="data-item-title">${escapeHtml(title)}</div><div class="data-item-meta">${escapeHtml(sub)}</div></li>`;
    })
    .join("");
}

function parseDateTimeSafe(value) {
  const raw = String(value ?? "").trim();
  if (!raw) return null;
  const d = new Date(raw);
  if (Number.isNaN(d.getTime())) return null;
  return d;
}

function resolveRepairCreatedAt(repair) {
  const candidates = [
    repair?.podio_created_at,
    repair?.created_at,
  ];

  for (const value of candidates) {
    const date = parseDateTimeSafe(value);
    if (date) return date;
  }

  return null;
}

function resolveRepairCompletedAt(repair) {
  const parseCompletionDate = (value) => {
    const raw = String(value ?? "").trim();
    if (!raw) return null;
    if (/^\d{4}-\d{2}-\d{2}$/.test(raw)) {
      // close_date est souvent un champ "date" (sans heure): on le traite en fin de journée.
      const endOfDay = new Date(`${raw}T23:59:59.999Z`);
      return Number.isNaN(endOfDay.getTime()) ? null : endOfDay;
    }
    return parseDateTimeSafe(raw);
  };

  const candidates = [
    repair?.close_date,
    repair?.termine_le,
    repair?.terminee_le,
    repair?.completed_at,
    repair?.finished_at,
    repair?.done_at,
    repair?.closed_at,
  ];

  for (const value of candidates) {
    const date = parseCompletionDate(value);
    if (date) return date;
  }

  return null;
}

function resolveRepairClosedAtForYearStats(repair) {
  const candidates = [
    repair?.close_date,
    repair?.termine_le,
    repair?.terminee_le,
    repair?.completed_at,
    repair?.finished_at,
    repair?.done_at,
    repair?.closed_at,
  ];

  for (const value of candidates) {
    const date = parseDateTimeSafe(value);
    if (date) return date;
  }

  return null;
}

function average(numbers) {
  const values = (numbers || []).filter((n) => Number.isFinite(n));
  if (!values.length) return null;
  const total = values.reduce((sum, n) => sum + Number(n), 0);
  return total / values.length;
}

function formatDurationFromMinutes(value) {
  if (!Number.isFinite(value)) return "-";
  const rounded = Math.max(0, Math.round(Number(value)));
  const hours = Math.floor(rounded / 60);
  const minutes = rounded % 60;
  if (hours > 0 && minutes > 0) return `${hours} h ${minutes} min`;
  if (hours > 0) return `${hours} h`;
  return `${minutes} min`;
}

function formatCycleDurationFromHours(value) {
  if (!Number.isFinite(value)) return "-";
  const totalMinutes = Math.max(0, Math.round(Number(value) * 60));
  const days = Math.floor(totalMinutes / 1440);
  const hours = Math.floor((totalMinutes % 1440) / 60);
  const minutes = totalMinutes % 60;

  const parts = [];
  if (days > 0) parts.push(`${days} j`);
  if (hours > 0) parts.push(`${hours} h`);
  if (minutes > 0 && days === 0) parts.push(`${minutes} min`);
  if (!parts.length) return "0 min";
  return parts.join(" ");
}

function formatCountAndPercent(count, total) {
  const safeCount = Number(count) || 0;
  const safeTotal = Number(total) || 0;
  const percent = safeTotal > 0 ? (safeCount / safeTotal) * 100 : 0;
  return `${safeCount} (${percent.toFixed(1)}%)`;
}

function formatPercent(count, total) {
  const safeCount = Number(count) || 0;
  const safeTotal = Number(total) || 0;
  const percent = safeTotal > 0 ? (safeCount / safeTotal) * 100 : 0;
  return `${percent.toFixed(1)}%`;
}

function buildCompanyAnniversaryInfo(now = new Date()) {
  const start = new Date(2024, 2, 4);
  start.setHours(0, 0, 0, 0);

  const today = new Date(now);
  today.setHours(0, 0, 0, 0);

  const isBeforeStart = today.getTime() < start.getTime();
  const msPerDay = 24 * 60 * 60 * 1000;
  const totalDays = isBeforeStart ? 0 : Math.floor((today.getTime() - start.getTime()) / msPerDay);

  const thisYearAnniv = new Date(today.getFullYear(), 2, 4);
  thisYearAnniv.setHours(0, 0, 0, 0);

  let years = today.getFullYear() - 2024;
  if (today.getTime() < thisYearAnniv.getTime()) years -= 1;
  if (years < 0) years = 0;

  const lastAnniv = new Date(2024 + years, 2, 4);
  lastAnniv.setHours(0, 0, 0, 0);
  const daysSinceLastAnniv = isBeforeStart ? 0 : Math.floor((today.getTime() - lastAnniv.getTime()) / msPerDay);

  const isAnniversaryDay = !isBeforeStart && today.getMonth() === 2 && today.getDate() === 4;
  const nextAnnivYear = today.getTime() <= thisYearAnniv.getTime() ? today.getFullYear() : today.getFullYear() + 1;
  const nextAnniv = new Date(nextAnnivYear, 2, 4);
  nextAnniv.setHours(0, 0, 0, 0);
  const daysUntilNext = Math.max(0, Math.ceil((nextAnniv.getTime() - today.getTime()) / msPerDay));

  return {
    years,
    totalDays,
    daysSinceLastAnniv,
    daysUntilNext,
    isAnniversaryDay,
  };
}

function renderAccueilCards() {
  const annivInfo = buildCompanyAnniversaryInfo();
  const annivTitle = $("accueilAnnivTitle");
  const annivText = $("accueilAnnivText");
  const annivCounter = $("accueilAnnivCounter");
  if (annivTitle) {
    annivTitle.textContent = annivInfo.isAnniversaryDay ? "Joyeux anniversaire!" : "Fondation: 4 mars 2024";
  }
  if (annivText) {
    annivText.textContent = annivInfo.isAnniversaryDay
      ? `Rivet Electronique fete ses ${annivInfo.years} an(s) aujourd'hui.`
      : `Age de l'entreprise: ${annivInfo.years} an(s) et ${annivInfo.daysSinceLastAnniv} jour(s) depuis le dernier anniversaire.`;
  }
  if (annivCounter) {
    annivCounter.textContent = annivInfo.isAnniversaryDay
      ? `Compteur: ${annivInfo.totalDays} jour(s) depuis la fondation.`
      : `Prochain anniversaire dans ${annivInfo.daysUntilNext} jour(s).`;
  }

  const repairs = (Array.isArray(state.repairs) ? state.repairs : []).filter((repair) => !isPersonalRepair(repair));
  const factures = Array.isArray(state.factures) ? state.factures : [];

  const totalRepairs = repairs.length;
  const doneRepairs = repairs.filter((r) => {
    const key = canonicalRepairStatus(r?.status_label);
    return key === "termine" || key === "remis_client";
  });

  const cycleHours = repairs.map((repair) => {
    const createdAt = resolveRepairCreatedAt(repair);
    const completedAt = resolveRepairCompletedAt(repair);
    if (!createdAt || !completedAt) return null;
    const diffMs = completedAt.getTime() - createdAt.getTime();
    if (!Number.isFinite(diffMs) || diffMs < 0) return null;
    return diffMs / 36e5;
  }).filter((value) => Number.isFinite(value));

  const durationMinutes = repairs.map((repair) => {
    const derived = deriveRepairLegacyValuesFromLines(repair?.lines);
    const value = Number(repair?.duration_minutes ?? derived.duration_minutes);
    if (!Number.isFinite(value) || value <= 0) return null;
    return value;
  }).filter((value) => Number.isFinite(value));

  const partsCosts = repairs.map((repair) => {
    const derived = deriveRepairLegacyValuesFromLines(repair?.lines);
    const value = Number(repair?.parts_cost ?? derived.parts_cost);
    if (!Number.isFinite(value) || value < 0) return null;
    return value;
  }).filter((value) => Number.isFinite(value));

  const avgCycle = average(cycleHours);
  const avgOperation = average(durationMinutes);
  const avgParts = average(partsCosts);

  const totalRepairsEl = $("accueilTotalRepairs");
  const completedRepairsEl = $("accueilCompletedRepairs");
  const avgCycleEl = $("accueilAvgCycle");
  const avgCycleSubEl = $("accueilAvgCycleSub");
  const avgOperationEl = $("accueilAvgOperation");
  const avgOperationSubEl = $("accueilAvgOperationSub");
  const avgPartsEl = $("accueilAvgPartsCost");
  const avgPartsSubEl = $("accueilAvgPartsCostSub");

  if (totalRepairsEl) totalRepairsEl.textContent = String(totalRepairs);
  if (completedRepairsEl) completedRepairsEl.textContent = `${doneRepairs.length} terminé(es)`;
  if (avgCycleEl) avgCycleEl.textContent = formatCycleDurationFromHours(avgCycle);
  if (avgCycleSubEl) avgCycleSubEl.textContent = `${cycleHours.length} réparation(s) calculée(s)`;
  if (avgOperationEl) avgOperationEl.textContent = formatDurationFromMinutes(avgOperation);
  if (avgOperationSubEl) avgOperationSubEl.textContent = `${durationMinutes.length} durée(s) disponibles`;
  if (avgPartsEl) avgPartsEl.textContent = avgParts == null ? "-" : formatMoney(avgParts);
  if (avgPartsSubEl) avgPartsSubEl.textContent = `${partsCosts.length} coût(s) disponible(s)`;

  const brandCounts = new Map();
  for (const repair of repairs) {
    const raw = cleanNullableText(repair?.manufacturer) || "Marque inconnue";
    brandCounts.set(raw, (brandCounts.get(raw) || 0) + 1);
  }
  const allBrands = [...brandCounts.entries()]
    .map(([name, count]) => ({ name, count }))
    .sort((a, b) => b.count - a.count);

  const brandTotal = allBrands.reduce((sum, row) => sum + row.count, 0);
  const topBrandRows = allBrands.slice(0, 6);
  if (allBrands.length > 6) {
    const rest = allBrands.slice(6).reduce((sum, row) => sum + row.count, 0);
    topBrandRows.push({ name: "Autres", count: rest });
  }

  const brandColors = ["#2f9e58", "#2ca9b7", "#9e7ad6", "#d26d88", "#ad8514", "#5a84d8", "#8f9aa3"];
  const chart = $("accueilBrandChart");
  const legend = $("accueilBrandLegend");
  const brandsMeta = $("accueilBrandsMeta");

  if (brandsMeta) brandsMeta.textContent = `${brandTotal} réparation(s)`;
  if (chart) {
    if (!brandTotal) {
      chart.classList.add("is-empty");
      chart.style.background = "";
      chart.setAttribute("aria-label", "Aucune donnée de marques");
    } else {
      let cursor = 0;
      const segments = topBrandRows.map((row, idx) => {
        const pct = (row.count / brandTotal) * 100;
        const start = cursor;
        const end = cursor + pct;
        cursor = end;
        return `${brandColors[idx % brandColors.length]} ${start.toFixed(2)}% ${end.toFixed(2)}%`;
      });
      chart.classList.remove("is-empty");
      chart.style.background = `conic-gradient(${segments.join(", ")})`;
      chart.setAttribute("aria-label", "Répartition des marques");
    }
  }

  if (legend) {
    if (!topBrandRows.length) {
      legend.innerHTML = `<li class="data-empty">Aucune marque disponible.</li>`;
    } else {
      legend.innerHTML = topBrandRows.map((row, idx) => {
        const color = brandColors[idx % brandColors.length];
        const detail = formatCountAndPercent(row.count, brandTotal);
        return `
          <li class="accueil-brand-item">
            <span class="accueil-brand-dot" style="background:${escapeHtml(color)}"></span>
            <span class="accueil-brand-name">${escapeHtml(row.name)}</span>
            <span class="accueil-brand-value">${escapeHtml(detail)}</span>
          </li>
        `;
      }).join("");
    }
  }

  const yearly = new Map();
  const ensureYear = (year) => {
    if (!yearly.has(year)) {
      yearly.set(year, {
        year,
        repairs: 0,
        factures: 0,
        totalFacture: 0,
        totalPaye: 0,
      });
    }
    return yearly.get(year);
  };

  for (const repair of repairs) {
    const doneAt = resolveRepairClosedAtForYearStats(repair);
    const year = doneAt?.getFullYear?.();
    if (!Number.isInteger(year)) continue;
    ensureYear(year).repairs += 1;
  }

  for (const facture of factures) {
    const source = parseDateTimeSafe(facture?.date_emission) || parseDateTimeSafe(facture?.created_at) || parseDateTimeSafe(facture?.updated_at);
    const year = source?.getFullYear?.();
    if (!Number.isInteger(year)) continue;
    const row = ensureYear(year);
    row.factures += 1;
    const total = Number(facture?.total);
    if (Number.isFinite(total)) row.totalFacture += total;
    if (isPaidValue(facture?.etat_label) && Number.isFinite(total)) row.totalPaye += total;
  }

  const yearlyRows = [...yearly.values()].sort((a, b) => b.year - a.year);
  const yearBody = $("accueilYearTableBody");
  const yearMeta = $("accueilYearMeta");
  if (yearMeta) yearMeta.textContent = `${yearlyRows.length} année(s)`;
  if (yearBody) {
    if (!yearlyRows.length) {
      yearBody.innerHTML = `<tr><td colspan="5" class="data-empty">Aucune donnée annuelle.</td></tr>`;
    } else {
      yearBody.innerHTML = yearlyRows.map((row) => `
        <tr>
          <td>${row.year}</td>
          <td>${row.repairs}</td>
          <td>${row.factures}</td>
          <td>${escapeHtml(formatMoney(row.totalFacture))}</td>
          <td>${escapeHtml(formatMoney(row.totalPaye))}</td>
        </tr>
      `).join("");
    }
  }

  const trafficChart = $("accueilTrafficChart");
  const trafficYear = 2025;
  const monthLabels = ["Jan", "Fev", "Mar", "Avr", "Mai", "Jun", "Jul", "Aou", "Sep", "Oct", "Nov", "Dec"];
  const monthlyRepairs = Array.from({ length: 12 }, () => 0);
  for (const repair of repairs) {
    const createdAt = resolveRepairCreatedAt(repair);
    if (!createdAt) continue;
    if (createdAt.getFullYear() !== trafficYear) continue;
    const month = createdAt.getMonth();
    if (month >= 0 && month <= 11) monthlyRepairs[month] += 1;
  }
  const yearlyTrafficTotal = monthlyRepairs.reduce((sum, value) => sum + value, 0);
  if (trafficChart) {
    if (!yearlyTrafficTotal) {
      trafficChart.classList.add("is-empty");
      trafficChart.innerHTML = `<p class="data-empty">Aucune réparation créée en ${trafficYear}.</p>`;
    } else {
      trafficChart.classList.remove("is-empty");
      const vbWidth = 1060;
      const vbHeight = 212;
      const padLeft = 34;
      const padRight = 8;
      const padTop = 10;
      const padBottom = 30;
      const plotWidth = vbWidth - padLeft - padRight;
      const plotHeight = vbHeight - padTop - padBottom;
      const maxValue = Math.max(...monthlyRepairs, 1);
      const yStep = Math.max(1, Math.ceil(maxValue / 4));
      const yMax = yStep * 4;
      const xForIndex = (idx) => padLeft + ((idx / 11) * plotWidth);
      const yForValue = (value) => padTop + plotHeight - ((value / yMax) * plotHeight);
      const points = monthlyRepairs.map((value, idx) => ({
        value,
        x: xForIndex(idx),
        y: yForValue(value),
      }));
      const linePoints = points.map((pt) => `${pt.x.toFixed(2)},${pt.y.toFixed(2)}`).join(" ");
      const yTicks = [0, yStep, yStep * 2, yStep * 3, yStep * 4];
      const horizontalGrid = yTicks.map((tick) => {
        const y = yForValue(tick).toFixed(2);
        return `<line class="accueil-traffic-grid" x1="${padLeft}" y1="${y}" x2="${(vbWidth - padRight).toFixed(2)}" y2="${y}"></line>`;
      }).join("");
      const yLabels = yTicks.map((tick) => {
        const y = yForValue(tick).toFixed(2);
        return `<text class="accueil-traffic-label" x="${(padLeft - 8).toFixed(2)}" y="${y}" text-anchor="end" dominant-baseline="middle">${tick}</text>`;
      }).join("");
      const monthGrid = monthLabels.map((_, idx) => {
        const x = xForIndex(idx).toFixed(2);
        return `<line class="accueil-traffic-grid" x1="${x}" y1="${padTop}" x2="${x}" y2="${(padTop + plotHeight).toFixed(2)}"></line>`;
      }).join("");
      const xLabels = monthLabels.map((label, idx) => {
        const x = xForIndex(idx).toFixed(2);
        return `<text class="accueil-traffic-label" x="${x}" y="${(vbHeight - 12).toFixed(2)}" text-anchor="middle">${escapeHtml(label)}</text>`;
      }).join("");
      const pointsSvg = points.map((pt) => `
        <circle class="accueil-traffic-point" cx="${pt.x.toFixed(2)}" cy="${pt.y.toFixed(2)}" r="3"></circle>
      `).join("");
      trafficChart.innerHTML = `
        <svg class="accueil-traffic-svg" viewBox="0 0 ${vbWidth} ${vbHeight}" role="img" aria-label="Réparations créées par mois en ${trafficYear}">
          ${monthGrid}
          ${horizontalGrid}
          <line class="accueil-traffic-axis" x1="${padLeft}" y1="${(padTop + plotHeight).toFixed(2)}" x2="${(vbWidth - padRight).toFixed(2)}" y2="${(padTop + plotHeight).toFixed(2)}"></line>
          <line class="accueil-traffic-axis" x1="${padLeft}" y1="${padTop}" x2="${padLeft}" y2="${(padTop + plotHeight).toFixed(2)}"></line>
          ${yLabels}
          ${xLabels}
          <polyline class="accueil-traffic-line" points="${linePoints}"></polyline>
          ${pointsSvg}
        </svg>
      `;
    }
  }

  const statusBody = $("accueilStatusTableBody");
  const statusMeta = $("accueilStatusMeta");
  const statusRows = BONS_COLUMNS.map((col) => {
    const count = repairs.filter((repair) => canonicalRepairStatus(repair?.status_label) === col.key).length;
    return { label: col.label, count };
  });
  const orderedStatusRows = statusRows;
  if (statusMeta) statusMeta.textContent = `${totalRepairs} réparation(s)`;
  if (statusBody) {
    statusBody.innerHTML = orderedStatusRows.map((row) => `
      <tr>
        <td>${escapeHtml(row.label)}</td>
        <td>${row.count}</td>
        <td>${escapeHtml(formatPercent(row.count, totalRepairs))}</td>
      </tr>
    `).join("");
  }
}

function bonsOtherSortTime(repair) {
  const source = repair?.updated_at || repair?.podio_created_at || repair?.created_at || "";
  const t = Date.parse(String(source));
  return Number.isFinite(t) ? t : 0;
}

function getRepairRepSortNumber(repair) {
  const appItemId = Number(repair?.podio_app_item_id);
  if (Number.isInteger(appItemId) && appItemId > 0) return appItemId;
  const repCode = formatRepairCode(repair);
  const match = String(repCode).match(/REP(\d+)/i);
  if (match?.[1]) {
    const n = Number(match[1]);
    if (Number.isInteger(n) && n > 0) return n;
  }
  const rowId = Number(repair?.id);
  if (Number.isInteger(rowId) && rowId > 0) return rowId;
  return Number.MAX_SAFE_INTEGER;
}

function isRepairFactureOverdue(repair) {
  const facture = getFactureForRepair(repair);
  if (!facture) return false;
  return getFactureEtatDisplay(facture).className === "is-overdue";
}

function compareBonsRepairCards(a, b, columnKey) {
  if (columnKey === "termine") {
    const overdueA = isRepairFactureOverdue(a);
    const overdueB = isRepairFactureOverdue(b);
    if (overdueA !== overdueB) return overdueA ? -1 : 1;
  }

  const repA = getRepairRepSortNumber(a);
  const repB = getRepairRepSortNumber(b);
  if (repA !== repB) return repA - repB;

  const tA = bonsOtherSortTime(a);
  const tB = bonsOtherSortTime(b);
  if (tA !== tB) return tB - tA;

  const idA = Number(a?.id) || 0;
  const idB = Number(b?.id) || 0;
  return idA - idB;
}

function getRepairAgeDays(repair) {
  const createdAt = resolveRepairCreatedAt(repair);
  if (!createdAt) return null;
  const diffMs = Date.now() - createdAt.getTime();
  if (!Number.isFinite(diffMs)) return null;
  if (diffMs <= 0) return 0;
  return Math.floor(diffMs / 86400000);
}

function getRepairAgeAlertTone(repair, statusKey) {
  const key = String(statusKey || "");
  if (key === "termine" || key === "remis_client") return null;
  const ageDays = getRepairAgeDays(repair);
  if (ageDays == null) return null;
  if (ageDays >= 28) return { tone: "red", ageDays };
  if (ageDays >= 21) return { tone: "orange", ageDays };
  if (ageDays >= 14) return { tone: "yellow", ageDays };
  return null;
}

function buildBonsOtherStatusFilterOptions(selectedValue) {
  const selected = String(selectedValue ?? "other") || "other";
  const options = [
    { value: "other", label: "Anciens bons (hors tableau)" },
    { value: "all", label: "Tous les statuts" },
    { value: "default", label: "Statuts du tableau" },
    ...BONS_COLUMNS.map((col) => ({ value: col.key, label: col.label })),
  ];
  return options
    .map((opt) => {
      const isSelected = opt.value === selected ? " selected" : "";
      return `<option value="${escapeHtml(opt.value)}"${isSelected}>${escapeHtml(opt.label)}</option>`;
    })
    .join("");
}

function matchBonsOtherStatusFilter(repair, filterValue) {
  const key = canonicalRepairStatus(repair?.status_label);
  const filter = String(filterValue ?? "other") || "other";
  if (filter === "all") return true;
  if (filter === "other") return !isBonsDefaultColumn(key);
  if (filter === "default") return isBonsDefaultColumn(key);
  return key === filter;
}

function matchBonsOtherSearch(repair, query) {
  const q = normalizeText(query);
  if (!q) return true;
  const repCode = formatRepairCode(repair);
  const client = repair?.client_title || "";
  const status = repair?.status_label || "";
  const device = getRepairCardDeviceLabel(repair) || repair?.title || "";
  const problem = htmlToText(repair?.desc_problem) || "";
  const haystack = normalizeText([repCode, client, status, device, problem].join(" "));
  return haystack.includes(q);
}

function renderBonsOtherList() {
  const statusFilter = $("bonsOtherStatusFilter");
  const searchInput = $("bonsOtherSearchInput");
  const body = $("bonsOtherTableBody");
  const meta = $("bonsOtherMeta");

  if (!statusFilter || !searchInput || !body || !meta) return;

  statusFilter.innerHTML = buildBonsOtherStatusFilterOptions(state.bonsOtherStatusFilter);
  if (searchInput.value !== state.bonsOtherSearchQuery) {
    searchInput.value = state.bonsOtherSearchQuery || "";
  }

  const allRepairs = Array.isArray(state.repairs) ? state.repairs : [];
  const sorted = [...allRepairs].sort((a, b) => bonsOtherSortTime(b) - bonsOtherSortTime(a));
  const filtered = sorted.filter((repair) => {
    if (!matchBonsOtherStatusFilter(repair, state.bonsOtherStatusFilter)) return false;
    if (!matchBonsOtherSearch(repair, state.bonsOtherSearchQuery)) return false;
    return true;
  });

  meta.textContent = sorted.length ? `${filtered.length} / ${sorted.length} bon(s)` : "";

  if (!filtered.length) {
    body.innerHTML = `<tr><td colspan="5" class="data-empty">Aucun bon pour ce filtre.</td></tr>`;
    return;
  }

  body.innerHTML = filtered
    .slice(0, 250)
    .map((repair) => {
      const repairId = repair?.id != null ? String(repair.id) : "";
      if (!repairId) return "";
      const repCode = formatRepairCode(repair);
      const statusLabel = cleanNullableText(repair?.status_label) || "-";
      const statusKey = canonicalRepairStatus(repair?.status_label);
      const statusClass = isBonsDefaultColumn(statusKey) ? "is-default" : "is-extra";
      const client = cleanNullableText(repair?.client_title) || "Sans client";
      const device = getRepairCardDeviceLabel(repair) || cleanNullableText(repair?.title) || "-";
      const updatedAt = formatDateTime(repair?.updated_at || repair?.podio_created_at || repair?.created_at) || "-";
      const statusToneClass = `bons-other-status-key-${statusKey || "other"}`;

      return `
        <tr class="bons-other-row" data-repair-id="${escapeHtml(repairId)}">
          <td>
            <button type="button" class="bons-other-open-btn" data-repair-id="${escapeHtml(repairId)}">${escapeHtml(repCode)}</button>
          </td>
          <td><span class="bons-other-status ${statusClass} ${escapeHtml(statusToneClass)}">${escapeHtml(statusLabel)}</span></td>
          <td>${escapeHtml(client)}</td>
          <td>${escapeHtml(device)}</td>
          <td>${escapeHtml(updatedAt)}</td>
        </tr>
      `;
    })
    .filter(Boolean)
    .join("");
}

function renderBonsBoard() {
  const board = $("bonsBoard");
  const meta = $("bonsMeta");
  if (!board || !meta) return;
  const drag = state.bonsDrag || {};
  const isDragging = Boolean(drag.activeRepairId);
  const overColumnKey = isDragging ? drag.overColumnKey : null;
  const sourceColumnKey = isDragging ? drag.sourceColumnKey : null;
  const isSavingMove = Boolean(drag.isSaving);

  if (!state.repairs.length) {
    meta.textContent = "";
    board.innerHTML = `<div class="data-empty">Aucune réparation trouvée.</div>`;
    board.classList.remove("is-dragging", "is-updating");
    return;
  }

  const grouped = Object.fromEntries(BONS_COLUMNS.map((col) => [col.key, []]));

  for (const repair of state.repairs) {
    const key = canonicalRepairStatus(repair.status_label);
    if (grouped[key]) grouped[key].push(repair);
  }

  const columns = BONS_COLUMNS.map((col) => ({
    ...col,
    items: [...(grouped[col.key] || [])].sort((a, b) => compareBonsRepairCards(a, b, col.key)),
  }));
  const visibleColumns = columns.filter((col) => isDragging || isBonsDefaultColumn(col.key));
  const activeCount = visibleColumns.filter((col) => col.items.length > 0).length;
  meta.textContent = "";

  if (!activeCount) {
    board.innerHTML = `<div class="data-empty">Aucune réparation dans les statuts configurés.</div>`;
    board.classList.remove("is-dragging", "is-updating");
    return;
  }

  board.classList.toggle("is-dragging", isDragging);
  board.classList.toggle("is-updating", isSavingMove);

  board.innerHTML = visibleColumns.map((col) => {
    const showFactureBadges = col.key === "termine" || col.key === "remis_client";
    const isEmptyColumn = col.items.length === 0;
    const isDefaultColumn = isBonsDefaultColumn(col.key);
    const isDropTarget = Boolean(overColumnKey) && overColumnKey === col.key;
    const isDragSource = Boolean(sourceColumnKey) && sourceColumnKey === col.key;
    const activeDragId = String(drag.activeRepairId || "");
    const visibleItems = isDragging
      ? col.items.filter((repair) => String(repair.id) === activeDragId)
      : col.items;
      const itemsHtml = visibleItems.map((repair) => {
        const repCode = formatRepairCode(repair);
        const title = getRepairCardDeviceLabel(repair) || `Réparation #${repair.id}`;
        const client = repair.client_title || "Sans client";
        const problem = htmlToText(repair.desc_problem) || "Sans description";
        const ageAlert = getRepairAgeAlertTone(repair, col.key);
        const ageBadgeHtml = ageAlert
          ? `<span class="bons-repair-age-badge is-${ageAlert.tone}" title="${escapeHtml(`Bon créé il y a ${ageAlert.ageDays} jour(s)`)}">${REPAIR_AGE_ALARM_SVG}</span>`
          : "";
        const thumbnailUrl = getRepairThumbnailUrlFromCache(repair);
        const thumbnailStyle = thumbnailUrl
          ? ` style="${escapeHtml(`--bons-thumb-url:url('${String(thumbnailUrl).replace(/'/g, "%27")}')`)}"`
          : "";
        const facture = getFactureForRepair(repair);
        const isIssued = Boolean(facture) || isYesValue(repair.fac_issued_label);
        const isPaid = facture ? isPaidValue(facture.etat_label) : isPaidValue(repair.fac_paid_label);
        const factureEtatDisplay = facture ? getFactureEtatDisplay(facture) : null;
        const overdueCardClass = factureEtatDisplay?.className === "is-overdue" ? " bons-repair-btn-overdue" : "";
      const factureNumber = formatFactureNumber(facture);
      const issuedLabel = isIssued
        ? (factureNumber ? `Facture émise ${factureNumber}` : "Facture émise")
        : "Facture non émise";
      const paidLabel = isPaid ? "Payée" : "Non payée";
      const issuedBadgeHtml = (!isIssued && showFactureBadges)
        ? `<button type="button" class="bons-pill bons-pill-issued is-not-issued bons-pill-create-facture-btn" data-facture-create="1" title="Créer une facture">${escapeHtml(issuedLabel)}</button>`
        : `<span class="bons-pill bons-pill-issued ${isIssued ? "is-issued" : "is-not-issued"}">${escapeHtml(issuedLabel)}</span>`;
      const paidPillHtml = isIssued
        ? `<span class="bons-pill bons-pill-paid ${isPaid ? "is-paid" : "is-unpaid"}">${escapeHtml(paidLabel)}</span>`
        : "";

      const secondaryLine = showFactureBadges
        ? `<div class="bons-repair-meta bons-repair-meta-split">
             ${issuedBadgeHtml}
             ${paidPillHtml}
           </div>`
        : `<div class="bons-repair-meta" title="${escapeHtml(problem)}">${escapeHtml(problem)}</div>`;

      return `
        <li>
          <div class="bons-repair-btn${thumbnailUrl ? " has-thumbnail" : ""}${overdueCardClass}${ageAlert ? " has-aging-badge" : ""}${isSavingMove ? " is-disabled" : ""}" role="button" tabindex="0" data-repair-id="${escapeHtml(repair.id)}" data-repair-status-key="${escapeHtml(col.key)}" aria-disabled="${isSavingMove ? "true" : "false"}"${thumbnailStyle}>
            <div class="bons-repair-line">
              <span class="bons-repair-title">
                <span class="bons-repair-code">${escapeHtml(repCode)}</span>
                <span class="bons-repair-separator" aria-hidden="true">|</span>
                <span class="bons-repair-title-text">${escapeHtml(title)}</span>
                ${ageBadgeHtml}
              </span>
              <span class="bons-repair-client">${escapeHtml(client)}</span>
            </div>
            ${secondaryLine}
          </div>
        </li>
      `;
    }).join("");
    const listHtml = itemsHtml || `<li class="bons-drop-placeholder">Déposer ici</li>`;
    const classNames = [
      "bons-column",
      `bons-column-status-${col.key}`,
      isEmptyColumn ? "bons-column-empty" : "",
      isDefaultColumn ? "" : "bons-column-extra",
      isDropTarget ? "is-drop-target" : "",
      isDragSource ? "is-drag-source" : ""
    ].filter(Boolean).join(" ");

    return `
      <article class="${classNames}" data-column-key="${escapeHtml(col.key)}">
        <div class="bons-column-head">
          <h3 class="bons-column-title">${escapeHtml(col.label)}</h3>
          <span class="bons-column-count">${col.items.length}</span>
        </div>
        <ul class="bons-list">${listHtml}</ul>
      </article>
    `;
  }).join("");

  const visibleRepairs = visibleColumns.flatMap((col) => col.items || []);
  void hydrateRepairThumbnailUrls(visibleRepairs)
    .then((didUpdate) => {
      if (!didUpdate) return;
      if ($("bonsBoard") !== board) return;
      renderBonsBoard();
    })
    .catch(() => {
      // ignore thumbnail hydration errors
    });
}

function estimationSortTime(item) {
  const source = item?.date_emission || item?.created_at || item?.updated_at || "";
  if (/^\d{4}-\d{2}-\d{2}/.test(source)) {
    const t = Date.parse(`${source.slice(0, 10)}T00:00:00Z`);
    return Number.isFinite(t) ? t : 0;
  }
  const t = Date.parse(source);
  return Number.isFinite(t) ? t : 0;
}

function normalizeDevisStatusKey(value) {
  const text = normalizeText(value);
  if (text === "en redaction" || text === "brouillon" || text === "draft" || text === "en_redaction") return "en_redaction";
  if (text === "emise" || text === "envoyee" || text === "sent") return "emise";
  if (text === "accepte" || text === "acceptee" || text === "accepted") return "accepte";
  if (text === "refuse" || text === "refusee" || text === "rejected") return "refuse";
  if (text === "annule" || text === "annulee" || text === "cancelled") return "annule";
  if (text === "echoue" || text === "expire" || text === "expired" || text === "failed") return "echoue";
  return "";
}

function matchDevisStatusFilter(item, filterKey) {
  const key = String(filterKey ?? "all").trim() || "all";
  if (key === "all") return true;
  const rowKey = normalizeDevisStatusKey(item?.etat_key) || normalizeDevisStatusKey(item?.etat_label);
  return rowKey === key;
}

function matchDevisSearch(item, query) {
  const q = normalizeText(query);
  if (!q) return true;
  const haystack = normalizeText([
    item?.numero,
    item?.client_title,
    item?.raw_title,
    item?.description,
    item?.etat_label,
    item?.total == null ? "" : String(item.total),
    formatDateCell(item?.date_emission || item?.created_at),
    formatDateCell(item?.date_valid_until),
    item?.contexte,
  ].join(" "));
  return haystack.includes(q);
}

function renderDevisList(tabKey, body, meta) {
  const headRow = $("facturesTableHeadRow");
  if (headRow) {
    headRow.innerHTML = `
      <th>No</th>
      <th>Client</th>
      <th>Date d'emission</th>
      <th>Valide jusqu'au</th>
      <th>Montant total</th>
      <th>Contexte</th>
      <th>Etat</th>
      <th>PDF</th>
    `;
  }

  if (state.estimationsLoadError) {
    meta.textContent = "Erreur";
    body.innerHTML = `<tr><td class="data-empty" colspan="8">${escapeHtml(state.estimationsLoadError)}</td></tr>`;
    return;
  }

  const wantedType = tabKey === "soumissions" ? "soumission" : "estimation";
  const docs = (state.estimations || []).filter((row) => normalizeText(row?.type_document) === wantedType);
  const sorted = [...docs].sort((a, b) => estimationSortTime(b) - estimationSortTime(a));
  const filtered = sorted.filter((row) => {
    if (!matchDevisStatusFilter(row, state.devisStatusFilter || "all")) return false;
    if (!matchDevisSearch(row, state.devisSearchQuery || "")) return false;
    return true;
  });
  const hasActiveFilters = Boolean(normalizeText(state.devisSearchQuery || "")) || String(state.devisStatusFilter || "all") !== "all";

  const label = tabKey === "soumissions" ? "soumission" : "estimation";
  meta.textContent = hasActiveFilters
    ? `${filtered.length} / ${sorted.length} ${label}(s)`
    : `${filtered.length} ${label}(s)`;

  if (!filtered.length) {
    body.innerHTML = `<tr><td class="data-empty" colspan="8">${hasActiveFilters ? `Aucune ${label} ne correspond aux filtres.` : `Aucune ${label} trouvée.`}</td></tr>`;
    return;
  }

  body.innerHTML = filtered.map((doc) => {
    const numero = cleanNullableText(doc?.numero) || "-";
    const client = cleanNullableText(doc?.client_title) || cleanNullableText(doc?.raw_title) || "Sans client";
    const dateEmission = formatDateCell(doc?.date_emission || doc?.created_at);
    const dateValid = formatDateCell(doc?.date_valid_until);
    const total = doc?.total != null ? formatMoney(doc.total) : "-";
    const contexte = normalizeText(doc?.contexte) === "reparation" ? "Reparation" : "Projet";
    const etat = isPaidValue(doc?.etat_label) ? "Payée" : (cleanNullableText(doc?.etat_label) || "-");
    const docId = String(doc?.id ?? "");
    const hasPdfFilename = Boolean(cleanNullableText(doc?.pdf_filename));
    const previewTitle = hasPdfFilename ? "Voir le PDF" : "Générer et voir le PDF";
    const canGeneratePreview = !isDeveloperModeEnabled();
    const previewDisabled = !docId || (!hasPdfFilename && !canGeneratePreview);
    return `
      <tr data-estimation-edit-id="${escapeHtml(docId)}">
        <td>${escapeHtml(numero)}</td>
        <td>${escapeHtml(client)}</td>
        <td>${escapeHtml(dateEmission)}</td>
        <td>${escapeHtml(dateValid)}</td>
        <td>${escapeHtml(total)}</td>
        <td>${escapeHtml(contexte)}</td>
        <td><span class="facture-etat is-open">${escapeHtml(etat)}</span></td>
        <td class="facture-preview-cell">
          <button
            type="button"
            class="facture-preview-btn"
            data-estimation-preview-id="${escapeHtml(docId)}"
            title="${escapeHtml(previewTitle)}"
            aria-label="${escapeHtml(previewTitle)}"
            ${previewDisabled ? "disabled" : ""}
          >${FACTURE_PREVIEW_EYE_SVG}</button>
        </td>
      </tr>
    `;
  }).join("");
}

function renderFacturesList() {
  const body = $("facturesTableBody");
  const meta = $("facturesMeta");
  const searchInput = $("facturesSearchInput");
  const yearFilterSelect = $("facturesYearFilterSelect");
  const statusFilterSelect = $("facturesStatusFilterSelect");
  const headRow = $("facturesTableHeadRow");
  if (!body || !meta) return;

  const activeTab = FACTURATION_TABS.has(state.facturationActiveTab) ? state.facturationActiveTab : "factures";
  if (activeTab !== "factures") {
    if (searchInput && searchInput.value !== state.devisSearchQuery) {
      searchInput.value = state.devisSearchQuery || "";
    }
    renderDevisList(activeTab, body, meta);
    return;
  }

  if (headRow) {
    headRow.innerHTML = `
      <th>No de facture</th>
      <th>Client</th>
      <th>Date d'emission</th>
      <th>Date d'echeance</th>
      <th>Montant total</th>
      <th>Réf.</th>
      <th>Etat</th>
      <th>Actions</th>
    `;
  }

  const allFactures = Array.isArray(state.factures) ? state.factures : [];
  const sorted = [...allFactures].sort((a, b) => factureSortTime(b) - factureSortTime(a));
  const repairsByItemId = new Map(
    (state.repairs || [])
      .filter((r) => r?.podio_item_id != null)
      .map((r) => [String(r.podio_item_id), r])
  );

  if (searchInput && searchInput.value !== state.facturesSearchQuery) {
    searchInput.value = state.facturesSearchQuery || "";
  }

  if (yearFilterSelect) {
    const availableYears = [...new Set(
      sorted
        .map((facture) => resolveFactureYear(facture))
        .filter((year) => Number.isInteger(year))
    )].sort((a, b) => b - a);
    const currentYear = new Date().getFullYear();
    const options = [
      { value: "all", label: "Toutes les annees" },
      { value: "current", label: `${currentYear} (annee en cours)` },
      { value: "previous", label: "Annees anterieures" },
      ...availableYears.map((year) => ({ value: `year:${year}`, label: String(year) })),
    ];
    yearFilterSelect.innerHTML = options
      .map((opt) => `<option value="${escapeHtml(opt.value)}">${escapeHtml(opt.label)}</option>`)
      .join("");
    const current = String(state.facturesYearFilter ?? "all");
    const resolved = options.some((opt) => opt.value === current) ? current : "all";
    state.facturesYearFilter = resolved;
    yearFilterSelect.value = resolved;
  }

  if (statusFilterSelect) {
    const allowed = new Set(["all", "paid", "unpaid", "overdue"]);
    const current = String(state.facturesStatusFilter ?? "all");
    const resolved = allowed.has(current) ? current : "all";
    state.facturesStatusFilter = resolved;
    statusFilterSelect.value = resolved;
  }

  const searchQuery = state.facturesSearchQuery || "";
  const yearFilter = state.facturesYearFilter || "all";
  const statusFilter = state.facturesStatusFilter || "all";
  const filtered = sorted.filter((facture) => {
    if (!matchFactureYearFilter(facture, yearFilter)) return false;
    if (!matchFactureStatusFilter(facture, statusFilter)) return false;
    if (!matchFactureSearch(facture, searchQuery, repairsByItemId)) return false;
    return true;
  });
  const hasActiveFilters = Boolean(normalizeText(searchQuery))
    || yearFilter !== "all"
    || statusFilter !== "all";

  meta.textContent = hasActiveFilters
    ? `${filtered.length} / ${sorted.length} facture(s)`
    : `${filtered.length} facture(s)`;

  if (!filtered.length) {
    body.innerHTML = `<tr><td class="data-empty" colspan="8">${hasActiveFilters ? "Aucune facture ne correspond aux filtres." : "Aucune facture trouvée."}</td></tr>`;
    return;
  }

  body.innerHTML = filtered.map((facture) => {
    const noFacture = formatFactureListNumber(facture);
    const client = facture.client_title || facture.raw_title || "Sans client";
    const dateEmission = formatDateCell(facture.date_emission || facture.created_at);
    const dateEcheance = formatDateCell(facture.date_echeance);
    const total = facture.total != null ? formatMoney(facture.total) : "-";
    const repCodes = formatFactureRepCodes(facture, repairsByItemId);
    const etatDisplay = getFactureEtatDisplay(facture);
    const etat = etatDisplay.label;
    const etatClass = etatDisplay.className;
    const hasPdf = Boolean(cleanNullableText(facture.pdf_filename));
    const factureId = String(facture.id ?? facture.podio_item_id ?? "");
    const previewTitle = hasPdf ? "Voir le PDF" : "Aucun PDF associé";

    return `
      <tr>
        <td>${escapeHtml(noFacture)}</td>
        <td>${escapeHtml(client)}</td>
        <td>${escapeHtml(dateEmission)}</td>
        <td>${escapeHtml(dateEcheance)}</td>
        <td>${escapeHtml(total)}</td>
        <td>${escapeHtml(repCodes)}</td>
        <td><span class="facture-etat ${etatClass}">${escapeHtml(etat)}</span></td>
        <td class="facture-preview-cell">
          <div class="facture-table-actions">
            <button
              type="button"
              class="facture-preview-btn"
              data-facture-edit-id="${escapeHtml(factureId)}"
              title="Modifier la facture"
              aria-label="Modifier la facture"
            >${FACTURE_EDIT_PENCIL_SVG}</button>
            <button
              type="button"
              class="facture-preview-btn"
              data-facture-preview-id="${escapeHtml(factureId)}"
              title="${escapeHtml(previewTitle)}"
              aria-label="${escapeHtml(previewTitle)}"
              ${hasPdf ? "" : "disabled"}
            >${FACTURE_PREVIEW_EYE_SVG}</button>
          </div>
        </td>
      </tr>
    `;
  }).join("");
}

function getExpenseById(id) {
  return (state.expenses || []).find((expense) => String(expense?.id) === String(id)) || null;
}

async function openExpenseInvoice(expense) {
  if (!expense) throw new Error("Dépense introuvable.");
  const invoicePath = cleanNullableText(expense.invoice_path);
  if (!invoicePath && !cleanNullableText(expense.preview_url)) {
    throw new Error("Aucune facture liée à cette dépense.");
  }
  const developerMode = isDeveloperModeEnabled();
  const client = state.authClient || ensureAuthClient();
  if (!developerMode && !client) throw new Error("Connexion Supabase indisponible.");
  const url = developerMode && cleanNullableText(expense.preview_url)
    ? cleanNullableText(expense.preview_url)
    : await getStorageSignedObjectUrl(client, expense.invoice_bucket || EXPENSES_STORAGE_BUCKET, invoicePath);
  if (!url) throw new Error("Impossible d'ouvrir la facture.");
  const fileName = cleanNullableText(expense.invoice_filename) || invoicePath.split("/").pop() || "Facture";
  const contentType = String(expense.invoice_content_type || "").toLowerCase();
  const isPdf = contentType.includes("pdf") || /\.pdf$/i.test(fileName);
  const isImage = contentType.startsWith("image/") || /\.(png|jpe?g|gif|webp|bmp|svg)$/i.test(fileName);
  if (isPdf) {
    const openedInModal = await openFacturePdfModal(url, `Aperçu PDF - ${fileName}`);
    if (openedInModal) return;
    if (openPdfInNewTab(url)) return;
  }
  if (isImage) {
    openImagePreviewGallery([
      {
        title: fileName,
        getUrl: () => Promise.resolve(url),
      }
    ], 0, fileName);
    return;
  }
  window.open(url, "_blank", "noopener,noreferrer");
}

function renderExpensesList() {
  const body = $("expensesTableBody");
  const meta = $("expensesMeta");
  if (!body || !meta) return;

  if (state.expensesLoadError) {
    meta.textContent = "Erreur";
    body.innerHTML = `<tr><td colspan="5" class="data-empty">${escapeHtml(state.expensesLoadError)}</td></tr>`;
    return;
  }

  const rows = [...(state.expenses || [])]
    .map((expense) => normalizeExpenseRow(expense))
    .filter(Boolean)
    .sort((a, b) => {
      const dateA = Date.parse(`${formatDateCell(a.purchase_date)}T00:00:00`);
      const dateB = Date.parse(`${formatDateCell(b.purchase_date)}T00:00:00`);
      return (Number.isFinite(dateB) ? dateB : 0) - (Number.isFinite(dateA) ? dateA : 0);
    });

  meta.textContent = `${rows.length} dépense(s)`;

  if (!rows.length) {
    body.innerHTML = `<tr><td colspan="5" class="data-empty">Aucune dépense trouvée.</td></tr>`;
    return;
  }

  body.innerHTML = rows.map((expense) => {
    const description = cleanNullableText(expense.description) || "-";
    const hasInvoice = Boolean(cleanNullableText(expense.invoice_path) || cleanNullableText(expense.preview_url));
    return `
      <tr>
        <td>
          <button type="button" class="expenses-open-btn" data-expense-id="${escapeHtml(expense.id)}">${escapeHtml(expense.title)}</button>
        </td>
        <td class="expenses-description-cell">${escapeHtml(description)}</td>
        <td>${escapeHtml(formatDateCell(expense.purchase_date))}</td>
        <td>${escapeHtml(formatMoney(expense.amount))}</td>
        <td class="expenses-invoice-cell">
          ${hasInvoice
            ? `<button type="button" class="facture-preview-btn" data-expense-open-invoice="${escapeHtml(expense.id)}" title="Voir la facture" aria-label="Voir la facture">${FACTURE_PREVIEW_EYE_SVG}</button>`
            : `<span class="expenses-invoice-empty">Aucune</span>`}
        </td>
      </tr>
    `;
  }).join("");
}

function renderClientsList() {
  const body = $("clientsTableBody");
  const meta = $("clientsMeta");
  if (!body || !meta) return;

  if (state.clientsLoadError) {
    meta.textContent = "Erreur";
    body.innerHTML = `<tr><td class="data-empty" colspan="6">${escapeHtml(state.clientsLoadError)}</td></tr>`;
    return;
  }

  const filtered = (state.clients || []).filter((client) => matchesClientSearch(client, state.clientsSearchQuery));
  const sorted = [...filtered].sort((a, b) => {
    return normalizeText(formatClientPrimaryName(a)).localeCompare(normalizeText(formatClientPrimaryName(b)), "fr-CA");
  });

  const total = state.clients.length;
  meta.textContent = state.clientsSearchQuery
    ? `${sorted.length} / ${total} client(s)`
    : `${sorted.length} client(s)`;

  if (!sorted.length) {
    body.innerHTML = `<tr><td class="data-empty" colspan="6">Aucun client trouvé.</td></tr>`;
    return;
  }

  const query = state.clientsSearchQuery;
  body.innerHTML = sorted.map((client) => {
    const primary = formatClientPrimaryName(client);
    const company = String(client.company ?? "").trim() || "-";
    const phone = formatPhoneForDisplay(client.phone);
    const email = String(client.email ?? "").trim() || "-";
    const address = String(client.address ?? "").trim() || "-";
    const type = normalizeClientType(client.type) || (String(client.type ?? "").trim() || "-");

    return `
      <tr class="clients-row" data-client-id="${escapeHtml(client.id)}">
        <td title="${escapeHtml(primary)}">
          <div class="clients-name">${highlightSearchText(primary, query)}</div>
        </td>
        <td title="${escapeHtml(company)}"><span class="clients-cell-ellipsis">${highlightSearchText(company, query)}</span></td>
        <td title="${escapeHtml(phone)}"><span class="clients-cell-phone">${highlightSearchText(phone, query)}</span></td>
        <td title="${escapeHtml(email)}"><span class="clients-cell-ellipsis">${highlightSearchText(email, query)}</span></td>
        <td title="${escapeHtml(address)}"><span class="clients-cell-ellipsis">${highlightSearchText(address, query)}</span></td>
        <td title="${escapeHtml(type)}"><span class="clients-cell-ellipsis">${highlightSearchText(type, query)}</span></td>
      </tr>
    `;
  }).join("");
}

function renderProjectsList() {
  const currentList = $("projectsList");
  const archivedList = $("projectsArchivedList");
  const meta = $("projectsMeta");
  const activeTab = state.projectsActiveTab === "archives" ? "archives" : "board";
  const list = activeTab === "archives" ? archivedList : currentList;
  if (!list || !meta) return;

  if (state.projectsLoadError) {
    meta.textContent = "Erreur";
    list.innerHTML = `<div class="data-empty">${escapeHtml(state.projectsLoadError)}</div>`;
    return;
  }

  const query = state.projectsSearchQuery || "";
  const allProjects = Array.isArray(state.projects) ? state.projects : [];
  const filtered = allProjects.filter((project) => matchProjectSearch(project, query));
  const tabProjects = filtered.filter((project) => (
    activeTab === "archives" ? isProjectArchived(project) : !isProjectArchived(project)
  ));
  const sorted = [...tabProjects].sort((a, b) => {
    const codeDelta = parseProjectCodeSortValue(formatProjectCode(b)) - parseProjectCodeSortValue(formatProjectCode(a));
    if (codeDelta !== 0) return codeDelta;
    const tA = Date.parse(a?.updated_at || a?.created_at || "") || 0;
    const tB = Date.parse(b?.updated_at || b?.created_at || "") || 0;
    return tB - tA;
  });

  meta.textContent = query
    ? `${sorted.length} / ${filtered.length} ${activeTab === "archives" ? "ancien(s) projet(s)" : "projet(s)"}`
    : `${sorted.length} ${activeTab === "archives" ? "ancien(s) projet(s)" : "projet(s)"}`;

  if (!sorted.length) {
    list.innerHTML = `<div class="data-empty">${activeTab === "archives" ? "Aucun ancien projet." : "Aucun projet trouve."}</div>`;
    return;
  }

  list.innerHTML = sorted.map((project) => {
    const projectId = String(project?.id ?? "");
    const code = formatProjectCode(project);
    const client = cleanNullableText(project?.client_title) || "Sans client";
    const projectTitle = cleanNullableText(project?.title) || "Sans titre";
    const progress = getProjectProgressSummary(project?.stages);
    const percent = Math.max(0, Math.min(100, Number(progress.percent || 0)));
    const phaseLabel = cleanNullableText(progress.phaseLabel) || "Ouverture";
    const description = cleanNullableText(project?.description) || cleanNullableText(project?.title) || "Sans description";
    const thumbUrl = getProjectThumbnailUrlFromCache(project);
    const thumbStyle = thumbUrl
      ? ` style="${escapeHtml(`--project-thumb-url:url('${String(thumbUrl).replace(/'/g, "%27")}')`)}"`
      : "";

    return `
      <button type="button" class="project-card${thumbUrl ? " has-thumbnail" : ""}" data-project-id="${escapeHtml(projectId)}"${thumbStyle}>
        <div class="project-card-head">
          <p class="project-card-title-line">
            <span class="project-card-code">${escapeHtml(code)}</span>
            <span class="project-card-sep" aria-hidden="true">|</span>
            <span class="project-card-client">${escapeHtml(client)}</span>
            <span class="project-card-sep" aria-hidden="true">|</span>
            <span class="project-card-project-title">${escapeHtml(projectTitle)}</span>
          </p>
        </div>
        <div class="project-thermometer" role="progressbar" aria-valuemin="0" aria-valuemax="100" aria-valuenow="${Math.round(percent)}">
          <span class="project-thermometer-fill" style="width:${Math.max(0, Math.min(100, percent))}%;"></span>
          <span class="project-thermometer-label">
            <span class="project-thermometer-phase">${escapeHtml(phaseLabel)}</span>
            <span class="project-thermometer-percent">${escapeHtml(`${percent}%`)}</span>
          </span>
        </div>
        <p class="project-card-description" title="${escapeHtml(description)}">${escapeHtml(description)}</p>
      </button>
    `;
  }).join("");

  void hydrateProjectThumbnailUrls(sorted)
    .then((didUpdate) => {
      if (!didUpdate) return;
      const nextList = state.projectsActiveTab === "archives" ? $("projectsArchivedList") : $("projectsList");
      if (nextList !== list) return;
      renderProjectsList();
    })
    .catch(() => {
      // ignore thumbnail hydration errors
    });
}

function getRepairById(id) {
  return state.repairs.find((repair) => String(repair.id) === String(id)) || null;
}

function getProjectById(id) {
  return (state.projects || []).find((project) => String(project?.id) === String(id)) || null;
}

function cleanNullableText(value) {
  const text = String(value ?? "").trim();
  return text || null;
}

function resolveEdgeFunctionInvokeError(error, fallbackMessage) {
  if (!error || typeof error !== "object") return fallbackMessage;
  const fallback = cleanNullableText(fallbackMessage) || "Erreur Edge Function.";
  const direct = cleanNullableText(error.message);
  const context = error.context;
  const contextDirect = cleanNullableText(
    typeof context === "string"
      ? context
      : context?.message || context?.error || context?.detail || context?.name
  );
  const genericFetchFailure = direct === "Failed to send a request to the Edge Function";
  if (!contextDirect && !genericFetchFailure) return direct || fallback;
  try {
    const parsed = typeof context === "string" ? JSON.parse(context) : context;
    if (parsed && typeof parsed === "object") {
      const detail = cleanNullableText(parsed.detail)
        || cleanNullableText(parsed.error)
        || cleanNullableText(parsed.message)
        || cleanNullableText(parsed.name);
      return detail || direct || fallback;
    }
  } catch {
    // ignore parse error and keep best available message
  }
  if (contextDirect) return contextDirect;
  if (genericFetchFailure) {
    return `${fallback} Vérifie que la fonction est déployée et que Synapse est ouverte via http://localhost:8000.`;
  }
  return direct || fallback;
}

function normalizeClientType(value) {
  const text = cleanNullableText(value);
  if (!text) return null;
  const normalized = normalizeText(text);
  if (normalized === "personnel") return "Personnel";
  if (normalized === "commercial") return "Commercial";
  return null;
}

function parseNullableNumber(value, { integer = false, min = 0 } = {}) {
  const raw = String(value ?? "").trim();
  if (!raw) return { ok: true, value: null };
  const parsed = integer ? Number.parseInt(raw, 10) : Number.parseFloat(raw.replace(",", "."));
  if (!Number.isFinite(parsed)) return { ok: false, error: "Valeur numérique invalide." };
  if (parsed < min) return { ok: false, error: `La valeur doit être supérieure ou égale à ${min}.` };
  return { ok: true, value: parsed };
}

async function updateRepairRecord(repairId, patch) {
  const client = state.authClient;
  if (!client) throw new Error("Connexion Supabase indisponible.");
  const current = getRepairById(repairId);
  if (!current) throw new Error("Réparation introuvable.");
  const nextRepairIsPersonal = Object.prototype.hasOwnProperty.call(patch || {}, "is_personal")
    ? patch?.is_personal === true
    : current?.is_personal === true;
  const normalizedPatchLines = Object.prototype.hasOwnProperty.call(patch || {}, "lines")
    ? ensureRepairInvoiceLines(patch?.lines, {
        durationMinutes: patch?.duration_minutes,
        partsCost: patch?.parts_cost,
        defaultLaborRate: getConfiguredFactureLaborRate(),
        hourlyRateSnapshot: current?.hourly_rate_snapshot,
        fallbackLaborQuantity: 0,
        includeLabor: !nextRepairIsPersonal,
      })
    : null;
  const patchWithDerivedCosts = normalizedPatchLines
    ? {
        ...patch,
        lines: normalizedPatchLines,
        ...deriveRepairLegacyValuesFromLines(normalizedPatchLines),
      }
    : patch;

  if (isDeveloperModeEnabled()) {
    const fallback = {
      ...current,
      ...patchWithDerivedCosts,
      updated_at: new Date().toISOString()
    };
    const hasRepair = state.repairs.some((repair) => String(repair.id) === String(fallback.id));
    const nextRepairs = state.repairs.map((repair) => (String(repair.id) === String(fallback.id) ? fallback : repair));
    state.repairs = hasRepair ? nextRepairs : [fallback, ...nextRepairs];
    return fallback;
  }

  const podioItemId = Number(current.podio_item_id);
  const directDbMode = !isPodioSyncEnabled();
  if (!Number.isInteger(podioItemId) || (!directDbMode && podioItemId <= 0)) {
    throw new Error("Réparation Podio introuvable.");
  }

  const { data, error } = await client.functions.invoke("repairs-submit", {
    body: {
      podio_item_id: podioItemId,
      direct_db: directDbMode,
      ...patchWithDerivedCosts
    }
  });
  if (error) throw new Error(resolveEdgeFunctionInvokeError(error, "Erreur repairs-submit."));
  if (!data?.ok) {
    const detail = data?.detail || data?.error || (directDbMode
      ? "Impossible de sauvegarder en base."
      : "Impossible de sauvegarder dans Podio.");
    throw new Error(detail);
  }

  const matchesPatch = (row, expectedPatch) => {
    if (!row || !expectedPatch) return false;
    const keys = Object.keys(expectedPatch);
    for (const key of keys) {
      const expected = expectedPatch[key];
      const actual = row[key];

      if (key === "status_label") {
        if (normalizeText(actual) !== normalizeText(expected)) return false;
        continue;
      }
      if (key === "client_item_id") {
        const a = actual == null ? null : Number(actual);
        const b = expected == null ? null : Number(expected);
        if (a !== b) return false;
        continue;
      }
      if (key === "is_personal") {
        if (Boolean(actual) !== Boolean(expected)) return false;
        continue;
      }
      if (key === "photos") {
        if (serializeRepairPhotos(actual) !== serializeRepairPhotos(expected)) return false;
        continue;
      }
      if (key === "photo_thumbnail") {
        if (serializeRepairPhotoThumbnail(actual) !== serializeRepairPhotoThumbnail(expected)) return false;
        continue;
      }
      if (key === "files") {
        if (serializeProjectFiles(actual) !== serializeProjectFiles(expected)) return false;
        continue;
      }
      if (key === "lines") {
        if (serializeRepairInvoiceLines(actual) !== serializeRepairInvoiceLines(expected)) return false;
        continue;
      }
      if (typeof expected === "number") {
        if (Number(actual) !== Number(expected)) return false;
        continue;
      }
      if (expected == null) {
        if (actual != null && String(actual).trim() !== "") return false;
        continue;
      }
      if (String(actual ?? "").trim() !== String(expected).trim()) return false;
    }
    return true;
  };

  const upsertRepairInState = (row) => {
    const hasRepair = state.repairs.some((r) => String(r.id) === String(row.id));
    const nextRepairs = state.repairs.map((r) => (String(r.id) === String(row.id) ? row : r));
    state.repairs = hasRepair ? nextRepairs : [row, ...nextRepairs];
    return row;
  };

  const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

  for (let attempt = 0; attempt < 8; attempt += 1) {
    const { data: syncedRow, error: fetchErr } = await client
      .from("repairs")
      .select("*")
      .eq("podio_item_id", podioItemId)
      .maybeSingle();

    if (!fetchErr && syncedRow && matchesPatch(syncedRow, patch)) {
      return upsertRepairInState(syncedRow);
    }
    await sleep(500);
  }

  // Fallback UX: si le webhook/sync est plus lent, on affiche une version locale.
  const fallback = {
    ...current,
    ...patchWithDerivedCosts,
    updated_at: new Date().toISOString()
  };
  return upsertRepairInState(fallback);
}

async function createRepairRecord(payload) {
  const client = state.authClient;
  if (!client) throw new Error("Connexion Supabase indisponible.");
  const normalizedLines = Object.prototype.hasOwnProperty.call(payload || {}, "lines")
    ? ensureRepairInvoiceLines(payload?.lines, {
        durationMinutes: payload?.duration_minutes,
        partsCost: payload?.parts_cost,
        hourlyRateSnapshot: payload?.hourly_rate_snapshot,
        defaultLaborRate: getConfiguredFactureLaborRate(),
        fallbackLaborQuantity: 1,
        includeLabor: payload?.is_personal !== true,
      })
    : ensureRepairInvoiceLines(payload?.lines, {
        durationMinutes: payload?.duration_minutes,
        partsCost: payload?.parts_cost,
        hourlyRateSnapshot: payload?.hourly_rate_snapshot,
        defaultLaborRate: getConfiguredFactureLaborRate(),
        fallbackLaborQuantity: 1,
        includeLabor: payload?.is_personal !== true,
      });
  const derivedRepairCosts = deriveRepairLegacyValuesFromLines(normalizedLines);
  const payloadWithDerivedCosts = {
    ...payload,
    lines: normalizedLines,
    ...derivedRepairCosts,
  };

  if (isDeveloperModeEnabled()) {
    const now = new Date().toISOString();
    const localRow = {
      id: nextLocalEntityId(),
      podio_item_id: nextLocalPodioItemId(),
      podio_app_item_id: nextLocalAppItemId(state.repairs),
      title: "",
      status_label: cleanNullableText(payload?.status_label) || "Dossier ouvert",
      manufacturer: cleanNullableText(payload?.manufacturer),
      model: cleanNullableText(payload?.model),
      serial: cleanNullableText(payload?.serial),
      client_item_id: payload?.client_item_id == null ? null : Number(payload.client_item_id),
      client_title: cleanNullableText(payload?.client_title),
      is_personal: payload?.is_personal === true,
      desc_problem: cleanNullableText(payload?.desc_problem),
      observations: cleanNullableText(payload?.observations),
      desc_done: cleanNullableText(payload?.desc_done),
      lines: normalizedLines,
      duration_minutes: derivedRepairCosts.duration_minutes,
      parts_cost: derivedRepairCosts.parts_cost,
      photos: normalizeRepairPhotos(payload?.photos),
      photo_thumbnail: normalizeRepairPhotoThumbnail(payload?.photo_thumbnail),
      fac_issued_label: "",
      fac_paid_label: "",
      created_at: now,
      updated_at: now,
    };
    const hasRepair = state.repairs.some((repair) => String(repair.id) === String(localRow.id));
    const nextRepairs = state.repairs.map((repair) => (String(repair.id) === String(localRow.id) ? localRow : repair));
    state.repairs = hasRepair ? nextRepairs : [localRow, ...nextRepairs];
    return localRow;
  }

  const directDbMode = !isPodioSyncEnabled();
  const { data, error } = await client.functions.invoke("repairs-submit", {
    body: {
      direct_db: directDbMode,
      ...payloadWithDerivedCosts
    }
  });
  if (error) throw new Error(resolveEdgeFunctionInvokeError(error, "Erreur repairs-submit."));
  if (!data?.ok) {
    const detail = data?.detail || data?.error || (directDbMode
      ? "Impossible de créer le bon en base."
      : "Impossible de créer le bon dans Podio.");
    throw new Error(detail);
  }

  const podioItemId = Number(data?.podio_item_id);
  if (!Number.isInteger(podioItemId) || (!directDbMode && podioItemId <= 0)) {
    throw new Error("Création Podio invalide (item_id manquant).");
  }

  const upsertRepairInState = (row) => {
    const hasRepair = state.repairs.some((r) => String(r.id) === String(row.id));
    const nextRepairs = state.repairs.map((r) => (String(r.id) === String(row.id) ? row : r));
    state.repairs = hasRepair ? nextRepairs : [row, ...nextRepairs];
    return row;
  };

  const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

  for (let attempt = 0; attempt < 12; attempt += 1) {
    const { data: syncedRow, error: fetchErr } = await client
      .from("repairs")
      .select("*")
      .eq("podio_item_id", podioItemId)
      .maybeSingle();

    if (!fetchErr && syncedRow) {
      return upsertRepairInState(syncedRow);
    }
    await sleep(500);
  }

  const fallback = {
    id: `pending-${podioItemId}`,
    podio_item_id: podioItemId,
    podio_app_item_id: null,
    title: "",
    status_label: cleanNullableText(payload?.status_label) || "Dossier ouvert",
    manufacturer: cleanNullableText(payload?.manufacturer),
    model: cleanNullableText(payload?.model),
    serial: cleanNullableText(payload?.serial),
    client_item_id: payload?.client_item_id == null ? null : Number(payload.client_item_id),
    client_title: cleanNullableText(payload?.client_title),
    is_personal: payload?.is_personal === true,
    desc_problem: cleanNullableText(payload?.desc_problem),
    observations: cleanNullableText(payload?.observations),
    desc_done: cleanNullableText(payload?.desc_done),
    lines: normalizedLines,
    duration_minutes: derivedRepairCosts.duration_minutes,
    parts_cost: derivedRepairCosts.parts_cost,
    photos: normalizeRepairPhotos(payload?.photos),
    photo_thumbnail: normalizeRepairPhotoThumbnail(payload?.photo_thumbnail),
    fac_issued_label: "",
    fac_paid_label: "",
    created_at: new Date().toISOString(),
    updated_at: new Date().toISOString(),
  };

  return upsertRepairInState(fallback);
}

async function createFactureRecord(payload) {
  const client = state.authClient;
  if (!client) throw new Error("Connexion Supabase indisponible.");

  if (isDeveloperModeEnabled()) {
    const devAppItemId = nextLocalAppItemId(state.factures);
    const devNumero = buildFactureNumeroFromSeqAndDate(devAppItemId, payload?.date_emission);
    let storageMeta = {
      pdf_stored: false,
      pdf_filename: null,
    };

    const { data, error } = await client.functions.invoke("factures-submit", {
      body: {
        ...payload,
        numero: cleanNullableText(payload?.numero) || devNumero,
        podio_app_item_id: devAppItemId,
        developer_mode: true,
      }
    });
    if (error) throw new Error(error.message || "Erreur factures-submit (mode développeur).");
    if (!data?.ok) {
      const detail = data?.detail || data?.error || "Impossible de traiter la facture en mode développeur.";
      throw new Error(detail);
    }
    storageMeta = {
      pdf_stored: Boolean(data?.pdf_stored),
      pdf_filename: data?.pdf_filename || null,
    };
    const savedNumero = cleanNullableText(data?.numero)
      || cleanNullableText(payload?.numero)
      || devNumero;

    const now = new Date().toISOString();
    const localRow = {
      id: Number.isInteger(Number(payload?.id)) ? Number(payload.id) : nextLocalEntityId(),
      podio_item_id: Number.isInteger(Number(payload?.podio_item_id)) ? Number(payload.podio_item_id) : nextLocalPodioItemId(),
      podio_app_item_id: Number.isInteger(Number(payload?.podio_app_item_id)) ? Number(payload.podio_app_item_id) : devAppItemId,
      numero: savedNumero,
      contexte: cleanNullableText(payload?.contexte),
      client_item_id: payload?.client_item_id == null ? null : Number(payload.client_item_id),
      reparations: Array.isArray(payload?.reparations) ? payload.reparations : [],
      lines: Array.isArray(payload?.lines) ? payload.lines : [],
      client_snapshot: payload?.client_snapshot && typeof payload?.client_snapshot === "object"
        ? payload.client_snapshot
        : null,
      company_snapshot: payload?.company_snapshot && typeof payload.company_snapshot === "object"
        ? payload.company_snapshot
        : null,
      raw_title: savedNumero,
      client_title: cleanNullableText(payload?.client_title),
      etat_label: cleanNullableText(payload?.status_label) || "Émise",
      montant_sans_taxes: Number(payload?.subtotal ?? 0),
      tps: Number(payload?.tps ?? 0),
      tvq: Number(payload?.tvq ?? 0),
      total: Number(payload?.total ?? 0),
      date_emission: cleanNullableText(payload?.date_emission),
      date_echeance: cleanNullableText(payload?.date_echeance),
      desc_problem: cleanNullableText(payload?.desc_problem),
      desc_done: cleanNullableText(payload?.desc_done),
      ...storageMeta,
      created_at: now,
      updated_at: now,
    };
    const hasFacture = state.factures.some((facture) => String(facture.id) === String(localRow.id));
    const nextFactures = state.factures.map((facture) => (String(facture.id) === String(localRow.id) ? localRow : facture));
    state.factures = hasFacture ? nextFactures : [localRow, ...nextFactures];
    state.factureByRepairItemId = buildFactureIndex(state.factures);
    return localRow;
  }

  const directDbMode = !isPodioSyncEnabled();
  const readFunctionInvokeErrorDetail = async (err) => {
    const response = err?.context;
    if (!response) return "";
    const readable = typeof response.clone === "function" ? response.clone() : response;
    if (!readable || typeof readable.text !== "function") return "";
    const rawText = await readable.text().catch(() => "");
    if (!rawText) return "";
    try {
      const parsed = JSON.parse(rawText);
      return String(parsed?.detail || parsed?.error || parsed?.message || rawText).trim();
    } catch {
      return String(rawText).trim();
    }
  };
  let data;
  let error;
  try {
    const result = await client.functions.invoke("factures-submit", {
      body: {
        ...payload,
        direct_db: directDbMode,
      }
    });
    data = result.data;
    error = result.error;
  } catch (invokeErr) {
    const detail = await readFunctionInvokeErrorDetail(invokeErr);
    if (detail) throw new Error(detail);
    throw invokeErr instanceof Error ? invokeErr : new Error(String(invokeErr));
  }
  if (error) {
    const detail = await readFunctionInvokeErrorDetail(error);
    throw new Error(detail || error.message || "Erreur factures-submit.");
  }
  if (!data?.ok) {
    const detail = data?.detail || data?.error || (directDbMode
      ? "Impossible de créer la facture en base."
      : "Impossible de créer la facture dans Podio.");
    throw new Error(detail);
  }

  const podioItemId = Number(data?.podio_item_id);
  if (!Number.isInteger(podioItemId) || (directDbMode ? podioItemId === 0 : podioItemId <= 0)) {
    throw new Error("Création facture invalide (item_id manquant).");
  }

  const upsertFactureInState = (row) => {
    const hasFacture = state.factures.some((f) => String(f.id) === String(row.id));
    const nextFactures = state.factures.map((f) => (String(f.id) === String(row.id) ? row : f));
    state.factures = hasFacture ? nextFactures : [row, ...nextFactures];
    state.factureByRepairItemId = buildFactureIndex(state.factures);
    return row;
  };

  const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));
  const selectColsWithNumero = "id,podio_item_id,podio_app_item_id,numero,client_item_id,client_title,reparations,lines,company_snapshot,pdf_filename,raw_title,etat_label,montant_sans_taxes,tps,tvq,total,date_emission,date_echeance,desc_problem,desc_done,updated_at,created_at";
  const selectColsWithoutNumero = "id,podio_item_id,podio_app_item_id,client_item_id,client_title,reparations,lines,company_snapshot,pdf_filename,raw_title,etat_label,montant_sans_taxes,tps,tvq,total,date_emission,date_echeance,desc_problem,desc_done,updated_at,created_at";
  let selectCols = selectColsWithNumero;
  let retriedWithoutNumero = false;

  for (let attempt = 0; attempt < 12; attempt += 1) {
    const { data: syncedRow, error: fetchErr } = await client
      .from("factures")
      .select(selectCols)
      .eq("podio_item_id", podioItemId)
      .maybeSingle();

    if (!fetchErr && syncedRow) {
      const finalRow = upsertFactureInState(syncedRow);
      void refreshTaxReportsData()
        .then(() => renderConfigurationView(true))
        .catch(() => {});
      return finalRow;
    }
    if (!retriedWithoutNumero && isMissingFacturesNumeroColumnError(fetchErr)) {
      selectCols = selectColsWithoutNumero;
      retriedWithoutNumero = true;
      continue;
    }
    await sleep(500);
  }

  const fallback = {
    id: `pending-${podioItemId}`,
    podio_item_id: podioItemId,
    podio_app_item_id: Number(data?.podio_app_item_id) || null,
    numero: cleanNullableText(data?.numero) || cleanNullableText(payload?.numero),
    contexte: cleanNullableText(payload?.contexte),
    client_item_id: payload?.client_item_id == null ? null : Number(payload.client_item_id),
    reparations: Array.isArray(payload?.reparations) ? payload.reparations : [],
    lines: Array.isArray(payload?.lines) ? payload.lines : [],
    client_snapshot: payload?.client_snapshot && typeof payload?.client_snapshot === "object"
      ? payload.client_snapshot
      : null,
    company_snapshot: payload?.company_snapshot && typeof payload.company_snapshot === "object"
      ? payload.company_snapshot
      : null,
    raw_title: cleanNullableText(data?.numero),
    client_title: cleanNullableText(payload?.client_title),
    etat_label: cleanNullableText(payload?.status_label) || "Émise",
    montant_sans_taxes: Number(payload?.subtotal ?? 0),
    tps: Number(payload?.tps ?? 0),
    tvq: Number(payload?.tvq ?? 0),
    total: Number(payload?.total ?? 0),
    date_emission: cleanNullableText(payload?.date_emission),
    date_echeance: cleanNullableText(payload?.date_echeance),
    desc_problem: cleanNullableText(payload?.desc_problem),
    desc_done: cleanNullableText(payload?.desc_done),
    pdf_filename: cleanNullableText(data?.pdf_filename),
    created_at: new Date().toISOString(),
    updated_at: new Date().toISOString(),
  };

  const finalRow = upsertFactureInState(fallback);
  void refreshTaxReportsData()
    .then(() => renderConfigurationView(true))
    .catch(() => {});
  return finalRow;
}

function buildDevisNumeroPrefix(typeDocument, emissionDateYmd) {
  const letter = normalizeText(typeDocument) === "soumission" ? "S" : "E";
  const ymd = String(emissionDateYmd ?? "").trim();
  if (!/^\d{4}-\d{2}-\d{2}$/.test(ymd)) return `${letter}000000`;
  const yy = ymd.slice(2, 4);
  const mm = ymd.slice(5, 7);
  const dd = ymd.slice(8, 10);
  return `${letter}${yy}${mm}${dd}`;
}

function nextLocalDevisNumero(typeDocument, emissionDateYmd) {
  const prefix = buildDevisNumeroPrefix(typeDocument, emissionDateYmd);
  const regex = new RegExp(`^${prefix}(\\d{4})(?:\\.\\d+)?$`, "i");
  let maxSeq = 0;
  for (const row of state.estimations || []) {
    const numero = String(row?.numero ?? "").trim().toUpperCase();
    const match = numero.match(regex);
    if (!match?.[1]) continue;
    const n = Number(match[1]);
    if (Number.isInteger(n) && n > maxSeq) maxSeq = n;
  }
  return `${prefix}${String(maxSeq + 1).padStart(4, "0")}`;
}

async function createDevisRecord(typeDocument) {
  openCreateDevisModal(typeDocument);
  return null;
}

async function updateFacturePaymentStatus(facture, isPaid) {
  const client = state.authClient;
  if (!client) throw new Error("Connexion Supabase indisponible.");

  const directDbMode = !isPodioSyncEnabled();
  const podioItemId = Number(facture?.podio_item_id);
  if (!Number.isInteger(podioItemId) || (directDbMode ? podioItemId === 0 : podioItemId <= 0)) {
    throw new Error(directDbMode ? "Identifiant facture introuvable." : "Facture Podio introuvable.");
  }

  const nextLabel = isPaid ? "Payée" : "Émise";
  const upsertFactureInState = (label) => {
    const now = new Date().toISOString();
    const currentFactures = Array.isArray(state.factures) ? state.factures : [];
    const nextFactures = currentFactures.map((row) => {
      if (String(row?.podio_item_id) !== String(podioItemId)) return row;
      return {
        ...row,
        etat_label: label,
        updated_at: now,
      };
    });
    state.factures = nextFactures;
    state.factureByRepairItemId = buildFactureIndex(state.factures);
    return state.factures.find((row) => String(row?.podio_item_id) === String(podioItemId)) || null;
  };

  if (isDeveloperModeEnabled()) {
    return upsertFactureInState(nextLabel);
  }

  const { data, error } = await client.functions.invoke("factures-status-update", {
    body: {
      podio_item_id: podioItemId,
      is_paid: Boolean(isPaid),
      direct_db: directDbMode,
    }
  });

  if (error) throw new Error(resolveEdgeFunctionInvokeError(error, "Erreur factures-status-update."));
  if (!data?.ok) {
    const detail = data?.detail || data?.error || "Impossible de mettre à jour l'état de paiement.";
    throw new Error(detail);
  }

  const finalLabel = cleanNullableText(data?.etat_label) || nextLabel;
  const updatedRow = upsertFactureInState(finalLabel);
  void refreshTaxReportsData()
    .then(() => renderConfigurationView(true))
    .catch(() => {});
  return updatedRow;
}

async function refreshTaxReportsData() {
  const client = state.authClient;
  if (!client) throw new Error("Connexion Supabase indisponible.");

  state.taxReportsLoading = true;
  const { data, error } = await client
    .from("tax_reports")
    .select("*")
    .order("report_year", { ascending: false });

  state.taxReportsLoading = false;
  if (error) {
    state.taxReports = [];
    state.taxReportsLoadError = `Documents fiscaux: ${error.message || "Impossible de charger les rapports."}`;
    throw error;
  }

  const normalizedReports = Array.isArray(data)
    ? data.map((row) => normalizeTaxReportRow(row)).filter(Boolean)
    : [];
  state.taxReports = await hydrateTaxReportStorageMeta(client, normalizedReports);
  state.taxReportsLoadError = null;
  return state.taxReports;
}

async function openTaxReportPdf(reportYear) {
  const client = state.authClient || ensureAuthClient();
  if (!client) throw new Error("Connexion Supabase indisponible.");
  const year = Number(reportYear);
  const report = (state.taxReports || []).find((row) => Number(row?.report_year) === year);
  if (!report) throw new Error("Rapport fiscal introuvable.");
  const storagePath = cleanNullableText(report.storage_path);
  if (!storagePath) throw new Error("Aucun PDF disponible pour ce rapport.");

  const pdfUrl = await getStorageSignedObjectUrl(client, report.storage_bucket || FISCAL_REPORTS_BUCKET, storagePath);
  const openedInModal = await openFacturePdfModal(pdfUrl, `Rapport fiscal ${year}`);
  if (!openedInModal) {
    const opened = openPdfInNewTab(pdfUrl);
    if (!opened) throw new Error("Impossible d'ouvrir ce PDF.");
  }
}

async function downloadTaxReportPdf(reportYear) {
  const client = state.authClient || ensureAuthClient();
  if (!client) throw new Error("Connexion Supabase indisponible.");
  const year = Number(reportYear);
  const report = (state.taxReports || []).find((row) => Number(row?.report_year) === year);
  if (!report) throw new Error("Rapport fiscal introuvable.");
  const storagePath = cleanNullableText(report.storage_path);
  if (!storagePath) throw new Error("Aucun PDF disponible pour ce rapport.");
  const bucket = cleanNullableText(report.storage_bucket) || FISCAL_REPORTS_BUCKET;
  const fileName = getTaxReportFileName(report, year);

  const signed = await client.storage
    .from(bucket)
    .createSignedUrl(storagePath, 900, { download: fileName });
  const pdfUrl = !signed?.error && signed?.data?.signedUrl
    ? signed.data.signedUrl
    : cleanNullableText(client.storage.from(bucket).getPublicUrl(storagePath)?.data?.publicUrl);
  if (!pdfUrl) throw new Error(signed?.error?.message || "Impossible de télécharger le rapport fiscal.");

  const anchor = document.createElement("a");
  anchor.href = pdfUrl;
  anchor.download = fileName;
  anchor.rel = "noopener noreferrer";
  anchor.style.display = "none";
  document.body.appendChild(anchor);
  anchor.click();
  anchor.remove();
}

async function generateFiscalReport(reportYear) {
  const client = state.authClient || ensureAuthClient();
  if (!client) throw new Error("Connexion Supabase indisponible.");
  const year = Number(reportYear);
  if (!Number.isInteger(year) || year <= 0) throw new Error("Ann\u00e9e de rapport invalide.");

  state.taxReportsGeneratingYear = year;
  renderConfigurationView(true);
  try {
    const { data, error } = await client.functions.invoke("fiscal-report-generate", {
      body: { report_year: year }
    });
    if (error) throw new Error(resolveEdgeFunctionInvokeError(error, "Erreur fiscal-report-generate."));
    if (!data?.ok) {
      throw new Error(data?.detail || data?.error || "Impossible de g\u00e9n\u00e9rer le rapport fiscal.");
    }
    await refreshTaxReportsData();
    renderConfigurationView(true);
    return data?.report || null;
  } finally {
    state.taxReportsGeneratingYear = null;
    renderConfigurationView(true);
  }
}

async function refreshClientsData() {
  const client = state.authClient;
  if (!client) throw new Error("Connexion Supabase indisponible.");

  const { data, error } = await client
    .from("clients")
    .select("id,podio_item_id,podio_app_item_id,name,title,company,phone,email,address,type,updated_at,created_at")
    .order("name", { ascending: true })
    .limit(1000);

  if (error) throw error;
  state.clients = data || [];
  state.clientsByPodioItemId = buildClientIndex(state.clients);
  state.clientsLoadError = null;
}

function upsertClientInState(row) {
  const hasClient = state.clients.some((client) => String(client.id) === String(row.id));
  const nextClients = state.clients.map((client) => (String(client.id) === String(row.id) ? row : client));
  state.clients = hasClient ? nextClients : [row, ...nextClients];
  state.clientsByPodioItemId = buildClientIndex(state.clients);
  state.clientsLoadError = null;
  return row;
}

function getClientById(id) {
  return state.clients.find((client) => String(client.id) === String(id)) || null;
}

function createEmptyClientDraft() {
  return {
    name: "",
    title: "",
    company: "",
    phone: "",
    email: "",
    address: "",
    type: ""
  };
}

function createEmptyRepairDraft() {
  return {
    id: `new-${Date.now()}`,
    podio_item_id: null,
    podio_app_item_id: null,
    title: "",
    status_label: "Dossier ouvert",
    manufacturer: "",
    model: "",
    serial: "",
    client_item_id: null,
    client_title: "",
    desc_problem: "",
    observations: "",
    desc_done: "",
    lines: [],
    duration_minutes: null,
    parts_cost: null,
    fac_issued_label: "",
    fac_paid_label: "",
    photos: [],
    photo_thumbnail: null,
    created_at: new Date().toISOString(),
    updated_at: new Date().toISOString(),
  };
}

function toClientEditableDraft(client) {
  return {
    name: String(client?.name ?? "").trim(),
    title: String(client?.title ?? "").trim(),
    company: String(client?.company ?? "").trim(),
    phone: normalizePhoneForInput(String(client?.phone ?? "").trim()),
    email: String(client?.email ?? "").trim(),
    address: String(client?.address ?? "").trim(),
    type: String(client?.type ?? "").trim()
  };
}

function clientDraftEquals(a, b) {
  return a.name === b.name
    && a.title === b.title
    && a.company === b.company
    && a.phone === b.phone
    && a.email === b.email
    && a.address === b.address
    && a.type === b.type;
}

function clientDraftDisplayName(draft) {
  const name = cleanNullableText(draft?.name);
  if (name) return name;
  const title = cleanNullableText(draft?.title);
  if (title) return title;
  return "sans nom";
}

function isClientModalDirty() {
  if (!clientModalState.original || !clientModalState.draft) return false;
  return !clientDraftEquals(clientModalState.original, clientModalState.draft);
}

function readClientModalDraftFromInputs() {
  if (!clientModalState.editing) return;
  const normalizedPhone = normalizePhoneForInput($("clientPhoneInput")?.value);
  const phoneInput = $("clientPhoneInput");
  if (phoneInput && phoneInput.value !== normalizedPhone) {
    phoneInput.value = normalizedPhone;
  }
  clientModalState.draft = {
    name: String($("clientNameInput")?.value ?? "").trim(),
    title: String($("clientTitleInput")?.value ?? "").trim(),
    company: String($("clientCompanyInput")?.value ?? "").trim(),
    phone: normalizedPhone,
    email: String($("clientEmailInput")?.value ?? "").trim(),
    address: String($("clientAddressInput")?.value ?? "").trim(),
    type: String($("clientTypeInput")?.value ?? "").trim()
  };
  syncClientAddressPartsFromInput();
}

function updateClientModalControlsState() {
  const saveBtn = $("clientEditSaveBtn");
  const cancelBtn = $("clientEditCancelBtn");
  const deleteBtn = $("clientDeleteBtn");
  const editBtn = $("clientModalEditBtn");
  const msg = $("clientEditMessage");
  const dirty = isClientModalDirty();

  if (saveBtn) saveBtn.disabled = clientModalState.saving || !dirty;
  if (cancelBtn) cancelBtn.disabled = clientModalState.saving;
  if (deleteBtn) deleteBtn.disabled = clientModalState.saving;
  if (editBtn) {
    editBtn.hidden = Boolean(clientModalState.isCreating);
    editBtn.classList.toggle("is-active", clientModalState.editing);
    editBtn.setAttribute("aria-pressed", String(clientModalState.editing));
  }
  if (msg) msg.textContent = clientModalState.message || "";
}

function renderClientModal() {
  const modal = $("clientModal");
  const title = $("clientModalTitle");
  const body = $("clientModalBody");
  if (!modal || !title || !body || !clientModalState.draft) return;

  const draft = clientModalState.draft;
  const displayName = clientDraftDisplayName(draft);
  const displayType = normalizeClientType(draft.type) || cleanNullableText(draft.type) || "-";
  title.textContent = clientModalState.isCreating ? "Nouveau client" : `Client | ${displayName}`;

  if (!clientModalState.editing) {
    body.innerHTML = `
      <section class="client-view-grid">
        <div class="client-view-item"><p class="client-view-label">Nom</p><p class="client-view-value">${escapeHtml(draft.name || "-")}</p></div>
        <div class="client-view-item"><p class="client-view-label">Titre</p><p class="client-view-value">${escapeHtml(draft.title || "-")}</p></div>
        <div class="client-view-item"><p class="client-view-label">Entreprise</p><p class="client-view-value">${escapeHtml(draft.company || "-")}</p></div>
        <div class="client-view-item"><p class="client-view-label">Type</p><p class="client-view-value">${escapeHtml(displayType)}</p></div>
        <div class="client-view-item"><p class="client-view-label">Téléphone</p><p class="client-view-value">${escapeHtml(draft.phone || "-")}</p></div>
        <div class="client-view-item"><p class="client-view-label">Courriel</p><p class="client-view-value">${escapeHtml(draft.email || "-")}</p></div>
        <div class="client-view-item client-view-item-wide"><p class="client-view-label">Adresse</p><p class="client-view-value">${escapeHtml(draft.address || "-")}</p></div>
      </section>
    `;
    updateClientModalControlsState();
    return;
  }

  const selectedType = normalizeClientType(draft.type) || "";
  const typeOptions = [`<option value="">-</option>`]
    .concat(
      CLIENT_TYPE_OPTIONS.map((option) => {
        const selected = option === selectedType ? " selected" : "";
        return `<option value="${escapeHtml(option)}"${selected}>${escapeHtml(option)}</option>`;
      })
    )
    .join("");
  const showDelete = !clientModalState.isCreating && Boolean(clientModalState.clientId);
  const deleteButtonHtml = showDelete
    ? `<button id="clientDeleteBtn" class="client-delete-btn" type="button">Supprimer ce client</button>`
    : "";

  body.innerHTML = `
    <form id="clientEditForm" class="client-edit-form">
      <div class="client-edit-grid">
        <div>
          <label class="client-input-label" for="clientNameInput">Nom</label>
          <input id="clientNameInput" class="client-input" type="text" value="${escapeHtml(draft.name)}" required>
        </div>
        <div>
          <label class="client-input-label" for="clientTitleInput">Titre</label>
          <input id="clientTitleInput" class="client-input" type="text" value="${escapeHtml(draft.title)}">
        </div>
        <div>
          <label class="client-input-label" for="clientCompanyInput">Entreprise</label>
          <input id="clientCompanyInput" class="client-input" type="text" value="${escapeHtml(draft.company)}">
        </div>
        <div>
          <label class="client-input-label" for="clientTypeInput">Type</label>
          <select id="clientTypeInput" class="client-input">
            ${typeOptions}
          </select>
        </div>
        <div>
          <label class="client-input-label" for="clientPhoneInput">Téléphone</label>
          <input id="clientPhoneInput" class="client-input" type="text" maxlength="12" placeholder="123-456-7890" value="${escapeHtml(draft.phone)}">
        </div>
        <div>
          <label class="client-input-label" for="clientEmailInput">Courriel</label>
          <input id="clientEmailInput" class="client-input" type="email" value="${escapeHtml(draft.email)}">
        </div>
      </div>
      <div>
        <label class="client-input-label" for="clientAddressInput">Adresse</label>
        <input
          id="clientAddressInput"
          class="client-input"
          type="text"
          list="clientAddressSuggestions"
          autocomplete="off"
          value="${escapeHtml(draft.address)}"
        >
        <datalist id="clientAddressSuggestions"></datalist>
      </div>
      <div class="client-form-actions">
        <p id="clientEditMessage" class="client-message"></p>
        <div class="client-form-actions-right">
          ${deleteButtonHtml}
          <button id="clientEditCancelBtn" class="client-cancel-btn" type="button">Annuler</button>
          <button id="clientEditSaveBtn" class="client-save-btn" type="submit">Sauvegarder</button>
        </div>
      </div>
    </form>
  `;

  const form = $("clientEditForm");
  if (form) {
    form.addEventListener("input", (ev) => {
      readClientModalDraftFromInputs();
      if (ev?.target?.id === "clientAddressInput") {
        scheduleAddressSuggestions($("clientAddressInput")?.value);
      }
      clientModalState.message = "";
      updateClientModalControlsState();
    });
    form.addEventListener("submit", async (ev) => {
      ev.preventDefault();
      await saveClientModalChanges();
    });
  }

  const addressInput = $("clientAddressInput");
  if (addressInput) {
    addressInput.addEventListener("focus", () => {
      scheduleAddressSuggestions(addressInput.value);
    });
  }

  clearClientAddressSuggestions();

  const cancelBtn = $("clientEditCancelBtn");
  if (cancelBtn) {
    cancelBtn.addEventListener("click", async () => {
      if (clientModalState.isCreating) {
        await closeClientModal();
        return;
      }
      await cancelClientModalEdit();
    });
  }

  const deleteBtn = $("clientDeleteBtn");
  if (deleteBtn) {
    deleteBtn.addEventListener("click", async () => {
      await deleteClientModalClient();
    });
  }

  updateClientModalControlsState();
}

async function updateClientRecord(clientId, draft) {
  const client = state.authClient;
  if (!client) throw new Error("Connexion Supabase indisponible.");
  const clientRow = clientId ? getClientById(clientId) : null;
  if (clientId && !clientRow) throw new Error("Client introuvable.");

  const name = cleanNullableText(draft.name);
  if (!name) throw new Error("Le nom du client est obligatoire.");

  const rawType = cleanNullableText(draft.type);
  const normalizedType = normalizeClientType(rawType);
  if (rawType && !normalizedType) {
    throw new Error("Le type doit être Personnel ou Commercial.");
  }

  const payload = {
    name,
    title: cleanNullableText(draft.title),
    company: cleanNullableText(draft.company),
    phone: cleanNullableText(normalizePhoneForInput(draft.phone)),
    email: cleanNullableText(draft.email),
    address: cleanNullableText(draft.address),
    address_parts: clientModalState.addressParts || null,
    type: normalizedType
  };

  if (isDeveloperModeEnabled()) {
    const now = new Date().toISOString();
    return upsertClientInState({
      ...(clientRow || {}),
      id: clientRow?.id ?? nextLocalEntityId(),
      podio_item_id: clientRow?.podio_item_id ?? nextLocalPodioItemId(),
      podio_app_item_id: clientRow?.podio_app_item_id ?? nextLocalAppItemId(state.clients),
      name: payload.name,
      title: payload.title,
      company: payload.company,
      phone: payload.phone,
      email: payload.email,
      address: payload.address,
      type: payload.type,
      updated_at: now,
      created_at: clientRow?.created_at || now,
    });
  }

  const directDbMode = !isPodioSyncEnabled();
  const podioItemId = Number(clientRow?.podio_item_id);
  if (clientId && (!Number.isInteger(podioItemId) || (directDbMode ? podioItemId === 0 : podioItemId <= 0))) {
    throw new Error(directDbMode ? "Identifiant client introuvable." : "Client Podio introuvable.");
  }
  if (clientId) payload.podio_item_id = podioItemId;

  const { data, error } = await client.functions.invoke("clients-submit", {
    body: {
      ...payload,
      direct_db: directDbMode,
    }
  });
  if (error) throw new Error(error.message || "Erreur clients-submit.");
  if (!data?.ok) {
    const detail = data?.detail || data?.error || (directDbMode
      ? "Impossible de sauvegarder le client en base."
      : "Impossible de sauvegarder dans Podio.");
    throw new Error(detail);
  }

  await refreshClientsData();
  const updatedByItemId = Number.isInteger(Number(data?.podio_item_id))
    ? state.clientsByPodioItemId[String(data.podio_item_id)]
    : null;
  const updated = (clientId ? getClientById(clientId) : null)
    || (clientId ? state.clientsByPodioItemId[String(podioItemId)] : null)
    || updatedByItemId
    || null;
  if (!updated) {
    throw new Error(clientId
      ? (directDbMode
        ? "Client modifié en base, mais introuvable après rafraîchissement."
        : "Client modifié dans Podio, mais introuvable après synchronisation.")
      : (directDbMode
        ? "Client créé en base, mais introuvable après rafraîchissement."
        : "Client créé dans Podio, mais introuvable après synchronisation."));
  }
  return updated;
}

async function deleteClientRecord(clientId) {
  const client = state.authClient;
  if (!client) throw new Error("Connexion Supabase indisponible.");
  const clientRow = clientId ? getClientById(clientId) : null;
  if (!clientRow) throw new Error("Client introuvable.");

  if (isDeveloperModeEnabled()) {
    state.clients = state.clients.filter((item) => String(item.id) !== String(clientId));
    state.clientsByPodioItemId = buildClientIndex(state.clients);
    state.clientsLoadError = null;
    return;
  }

  const directDbMode = !isPodioSyncEnabled();
  const podioItemId = Number(clientRow.podio_item_id);
  if (!Number.isInteger(podioItemId) || (directDbMode ? podioItemId === 0 : podioItemId <= 0)) {
    throw new Error(directDbMode ? "Identifiant client introuvable." : "Client Podio introuvable.");
  }

  const { data, error } = await client.functions.invoke("clients-submit", {
    body: {
      podio_item_id: podioItemId,
      delete: true,
      direct_db: directDbMode,
    }
  });
  if (error) throw new Error(error.message || "Erreur clients-submit.");
  if (!data?.ok) {
    const detail = data?.detail || data?.error || "Impossible de supprimer le client.";
    throw new Error(detail);
  }

  await refreshClientsData();
}

async function deleteClientModalClient() {
  if (clientModalState.saving) return;
  if (clientModalState.isCreating) return;
  if (!clientModalState.clientId) return;

  const name = clientDraftDisplayName(clientModalState.draft);
  const ok = window.confirm(`Tu es certain de vouloir supprimer le client ${name} ? Cette action est irréversible.`);
  if (!ok) return;

  try {
    clientModalState.saving = true;
    clientModalState.message = "Suppression...";
    updateClientModalControlsState();
    await deleteClientRecord(clientModalState.clientId);
    clientModalState.saving = false;
    clientModalState.message = "";
    clientModalState.editing = false;
    renderClientsList();
    await closeClientModal();
  } catch (err) {
    clientModalState.saving = false;
    clientModalState.message = err?.message || "Impossible de supprimer le client.";
    updateClientModalControlsState();
  }
}

function openClientModal(clientRow) {
  const modal = $("clientModal");
  if (!modal || !clientRow) return;

  clientModalState.clientId = clientRow.id;
  clientModalState.original = toClientEditableDraft(clientRow);
  clientModalState.draft = toClientEditableDraft(clientRow);
  clientModalState.editing = false;
  clientModalState.saving = false;
  clientModalState.message = "";
  clientModalState.isCreating = false;
  clientModalState.onCreated = null;
  resetClientAddressSuggestState();

  renderClientModal();
  modal.hidden = false;
  document.body.style.overflow = "hidden";
}

function openCreateClientModal(options = {}) {
  const modal = $("clientModal");
  if (!modal) return;

  clientModalState.clientId = null;
  clientModalState.original = createEmptyClientDraft();
  clientModalState.draft = createEmptyClientDraft();
  clientModalState.editing = true;
  clientModalState.saving = false;
  clientModalState.message = "";
  clientModalState.isCreating = true;
  clientModalState.onCreated = typeof options?.onCreated === "function" ? options.onCreated : null;
  resetClientAddressSuggestState();

  renderClientModal();
  modal.hidden = false;
  document.body.style.overflow = "hidden";
}

async function cancelClientModalEdit() {
  if (!clientModalState.editing) return;
  readClientModalDraftFromInputs();
  const dirty = isClientModalDirty();
  if (dirty) {
    const name = clientDraftDisplayName(clientModalState.draft);
    const ok = window.confirm(`Des modifications ont été apportées au client ${name}. Êtes-vous certain de vouloir annuler ?`);
    if (!ok) return;
  }
  clientModalState.draft = { ...clientModalState.original };
  clientModalState.editing = false;
  clientModalState.message = "";
  renderClientModal();
}

async function saveClientModalChanges() {
  if (!clientModalState.editing || clientModalState.saving) return;
  readClientModalDraftFromInputs();
  if (!isClientModalDirty()) {
    updateClientModalControlsState();
    return;
  }

  const name = clientDraftDisplayName(clientModalState.draft);
  const ok = window.confirm(`Des modifications ont été apportées au client ${name}. Êtes-vous certain de vouloir sauvegarder ?`);
  if (!ok) return;

  try {
    const wasCreating = Boolean(clientModalState.isCreating);
    const onCreated = typeof clientModalState.onCreated === "function" ? clientModalState.onCreated : null;
    clientModalState.saving = true;
    clientModalState.message = "Sauvegarde...";
    updateClientModalControlsState();
    const updated = await updateClientRecord(clientModalState.clientId, clientModalState.draft);
    clientModalState.clientId = updated.id;
    clientModalState.original = toClientEditableDraft(updated);
    clientModalState.draft = toClientEditableDraft(updated);
    clientModalState.editing = false;
    clientModalState.saving = false;
    clientModalState.message = "";
    clientModalState.isCreating = false;
    renderClientsList();
    if (wasCreating && onCreated) {
      try {
        onCreated(updated);
      } catch (_e) {
        // no-op: callback UX helper
      }
      await closeClientModal();
      return;
    }
    renderClientModal();
  } catch (err) {
    clientModalState.saving = false;
    clientModalState.message = err?.message || "Impossible de sauvegarder le client.";
    updateClientModalControlsState();
  }
}

async function toggleClientModalEditMode() {
  if (clientModalState.saving) return;
  if (clientModalState.isCreating) return;
  if (!clientModalState.clientId) return;
  if (!clientModalState.editing) {
    clientModalState.editing = true;
    clientModalState.message = "";
    renderClientModal();
    return;
  }
  await cancelClientModalEdit();
}

async function closeClientModal() {
  const modal = $("clientModal");
  if (!modal || modal.hidden) return;
  if (clientModalState.saving) return;

  if (clientModalState.editing) {
    readClientModalDraftFromInputs();
    const dirty = isClientModalDirty();
    if (dirty) {
      const name = clientDraftDisplayName(clientModalState.draft);
      const ok = window.confirm(`Des modifications ont été apportées au client ${name}. Êtes-vous certain de vouloir annuler ?`);
      if (!ok) return;
    }
  }

  modal.hidden = true;
  clientModalState.clientId = null;
  clientModalState.original = null;
  clientModalState.draft = null;
  clientModalState.editing = false;
  clientModalState.saving = false;
  clientModalState.message = "";
  clientModalState.isCreating = false;
  clientModalState.onCreated = null;
  resetClientAddressSuggestState();
  if (
    (!$("repairModal") || $("repairModal").hidden)
    && (!$("factureModal") || $("factureModal").hidden)
    && (!$("projectModal") || $("projectModal").hidden)
    && (!$("expenseModal") || $("expenseModal").hidden)
  ) {
    document.body.style.overflow = "";
  }
}

function openRepairModal(repair, options = {}) {
  const modal = $("repairModal");
  const title = $("repairModalTitle");
  const body = $("repairModalBody");
  if (!modal || !title || !body || !repair) return;
  if (typeof modal.__repairCleanup === "function") {
    try {
      modal.__repairCleanup();
    } catch {
      // ignore
    }
  }
  modal.__repairCleanup = null;

  const isCreatingRepair = Boolean(options?.isCreating);
  const repCode = formatRepairCode(repair);
  const repCodeLabel = repCode || "Nouveau bon";
  const initialStatusValue = cleanNullableText(repair.status_label) || (isCreatingRepair ? "Dossier ouvert" : null);
  let draftStatusValue = initialStatusValue;
  const initialIsPersonalRepair = isPersonalRepair(repair);
  let isPersonalRepairValue = initialIsPersonalRepair;
  const statusFallbackLabel = isCreatingRepair ? "Dossier ouvert" : "Statut inconnu";
  const initialRepairLines = ensureRepairInvoiceLines(repair?.lines, {
    durationMinutes: repair?.duration_minutes,
    partsCost: repair?.parts_cost,
    hourlyRateSnapshot: repair?.hourly_rate_snapshot,
    defaultLaborRate: getConfiguredFactureLaborRate(),
    fallbackLaborQuantity: isCreatingRepair ? 1 : 0,
    includeLabor: !initialIsPersonalRepair,
  });
  const initialDerivedRepairCosts = deriveRepairLegacyValuesFromLines(initialRepairLines);
  const initialFormValues = {
    status_label: initialStatusValue,
    desc_problem: cleanNullableText(htmlToText(repair.desc_problem)),
    observations: cleanNullableText(htmlToText(repair.observations)),
    desc_done: cleanNullableText(htmlToText(repair.desc_done)),
    manufacturer: cleanNullableText(repair.manufacturer),
    model: cleanNullableText(repair.model),
    serial: cleanNullableText(repair.serial),
    lines: initialRepairLines,
    duration_minutes: initialDerivedRepairCosts.duration_minutes,
    parts_cost: initialDerivedRepairCosts.parts_cost,
    is_personal: initialIsPersonalRepair,
    photos: normalizeRepairPhotos(repair.photos),
    photo_thumbnail: normalizeRepairPhotoThumbnail(repair.photo_thumbnail),
  };
  const initialFacture = isCreatingRepair ? null : getFactureForRepair(repair);
  const initialFacturePaidValue = initialFacture ? isPaidValue(initialFacture.etat_label) : null;
  let draftFacturePaidValue = initialFacturePaidValue;
  const hasFacturePaymentDirty = () => {
    if (!initialFacture || initialFacturePaidValue == null || draftFacturePaidValue == null) return false;
    return Boolean(draftFacturePaidValue) !== Boolean(initialFacturePaidValue);
  };
  let hasUnsavedChanges = false;
  const manufacturer = cleanNullableText(repair.manufacturer);
  const model = cleanNullableText(repair.model);

  const linkedClient = resolveClientForRepair(repair);
  const clientName = linkedClient ? formatClientPrimaryName(linkedClient) : (cleanNullableText(repair.client_title) || "Sans client");
  const clientAddress = cleanNullableText(linkedClient?.address);
  const clientCompany = cleanNullableText(linkedClient?.company);
  const clientPhone = cleanNullableText(linkedClient?.phone);
  const clientEmail = cleanNullableText(linkedClient?.email);
  const clientInfoHtml = buildRepairClientInfoHtml(
    linkedClient,
    clientAddress,
    clientCompany,
    clientPhone,
    clientEmail,
  );
  let selectedClientId = String(getClientReferenceId(linkedClient) ?? repair.client_item_id ?? "").trim();
  let selectedClientItemId = selectedClientId ? Number(selectedClientId) : null;
  let selectedClientTitle = linkedClient ? formatClientPrimaryName(linkedClient) : cleanNullableText(repair.client_title);
  let selectedClientCompany = cleanNullableText(linkedClient?.company);
  if (isCreatingRepair) {
    selectedClientId = "";
    selectedClientItemId = null;
    selectedClientTitle = null;
    selectedClientCompany = null;
  }
  const initialClientState = {
    itemId: selectedClientItemId,
    title: selectedClientTitle,
    company: selectedClientCompany,
  };
  if (initialIsPersonalRepair) {
    const personalOwner = getPersonalOwnerSelection();
    selectedClientId = personalOwner.itemId != null ? String(personalOwner.itemId) : "";
    selectedClientItemId = personalOwner.itemId;
    selectedClientTitle = personalOwner.title;
    selectedClientCompany = personalOwner.company;
    initialClientState.itemId = personalOwner.itemId;
    initialClientState.title = personalOwner.title;
    initialClientState.company = personalOwner.company;
  }
  const hasCurrentClientLink = () => Boolean(
    (selectedClientId && String(selectedClientId).trim())
    || cleanNullableText(selectedClientTitle),
  );
  const serial = cleanNullableText(repair.serial) || "-";
  const marqueValue = manufacturer || "-";
  const modeleValue = model || "-";
  const serialValue = serial;
  const deviceMainLabel = [marqueValue, modeleValue].filter((v) => v && v !== "-").join(" ").trim() || "-";
  const hasSerialValue = Boolean(serialValue && serialValue !== "-");
  let draftRepairLines = normalizeRepairInvoiceLines(initialFormValues.lines);
  let quickRepairCostDraft = null;
  const createdAt = isCreatingRepair ? "-" : (formatDateTime(repair.podio_created_at || repair.created_at) || "-");
  let draftRepairPhotos = normalizeRepairPhotos(initialFormValues.photos);
  let existingRepairFiles = normalizeProjectFiles(repair.files);
  let draftPendingPhotoEntries = [];
  let draftPhotoThumbnail = normalizeRepairPhotoThumbnail(initialFormValues.photo_thumbnail);
  let draftPhotoThumbnailPendingBlob = null;
  let draftPhotoThumbnailPendingPreviewUrl = null;
  let draftPhotoThumbnailPendingSource = null;
  const repairPhotoObjectUrls = new Set();
  const repairPhotoPreviewUrlByFilename = {};
  const revokeRepairPhotoObjectUrl = (url) => {
    if (!url) return;
    if (!repairPhotoObjectUrls.has(url)) return;
    try {
      URL.revokeObjectURL(url);
    } catch {
      // ignore
    }
    repairPhotoObjectUrls.delete(url);
  };
  const cleanupRepairPhotoResources = () => {
    for (const previewUrl of Array.from(repairPhotoObjectUrls)) {
      revokeRepairPhotoObjectUrl(previewUrl);
    }
    draftPendingPhotoEntries = [];
    draftPhotoThumbnailPendingBlob = null;
    draftPhotoThumbnailPendingSource = null;
    if (draftPhotoThumbnailPendingPreviewUrl) {
      revokeRepairPhotoObjectUrl(draftPhotoThumbnailPendingPreviewUrl);
      draftPhotoThumbnailPendingPreviewUrl = null;
    }
    for (const key of Object.keys(repairPhotoPreviewUrlByFilename)) {
      delete repairPhotoPreviewUrlByFilename[key];
    }
  };

  let clientPickerRows = [];
  const rebuildRepairClientPickerRows = () => {
    clientPickerRows = [...state.clients]
      .map((client) => {
        const id = String(getClientReferenceId(client) ?? "").trim();
        if (!id) return null;
        return {
          id,
          name: formatClientPrimaryName(client),
          company: cleanNullableText(client?.company),
        };
      })
      .filter(Boolean)
      .sort((a, b) => normalizeText(a.name).localeCompare(normalizeText(b.name), "fr-CA"));
  };
  rebuildRepairClientPickerRows();
  const setSelectedRepairClientById = (clientId) => {
    const nextId = String(clientId ?? "").trim();
    if (!nextId) return false;
    const selectedClient = state.clientsByPodioItemId[nextId];
    if (!selectedClient) return false;

    const clientItemId = getClientReferenceId(selectedClient);
    selectedClientId = nextId;
    selectedClientItemId = Number.isInteger(clientItemId) ? clientItemId : null;
    selectedClientTitle = formatClientPrimaryName(selectedClient);
    selectedClientCompany = cleanNullableText(selectedClient.company);
    return true;
  };

  function sameNullableNumber(a, b) {
    if (a == null && b == null) return true;
    if (a == null || b == null) return false;
    return Number(a) === Number(b);
  }

  function readRepairFormValues() {
    const statusSelect = $("repairStatusSelect");
    if (statusSelect) {
      const selected = cleanNullableText(statusSelect.value);
      if (selected) draftStatusValue = selected;
    }
    const nextLines = normalizeRepairInvoiceLines(draftRepairLines).map((line) => syncRepairCostLineDerivedValues({ ...line }));
    const derivedCosts = deriveRepairLegacyValuesFromLines(nextLines);

    return {
      status_label: cleanNullableText(draftStatusValue),
      desc_problem: cleanNullableText($("repairProblemInput")?.value),
      observations: cleanNullableText($("repairObservationInput")?.value),
      desc_done: cleanNullableText($("repairWorkDoneInput")?.value),
      manufacturer: cleanNullableText($("repairManufacturerInput")?.value),
      model: cleanNullableText($("repairModelInput")?.value),
      serial: cleanNullableText($("repairSerialInput")?.value),
      lines: nextLines,
      duration_minutes: derivedCosts.duration_minutes,
      parts_cost: derivedCosts.parts_cost,
      is_personal: isPersonalRepairValue,
    };
  }

  function isRepairFormDirty() {
    if (quickRepairCostDraft) return true;
    const current = readRepairFormValues();
    const currentClientId = selectedClientItemId == null ? null : Number(selectedClientItemId);
    const initialClientId = initialClientState.itemId == null ? null : Number(initialClientState.itemId);
    if (currentClientId !== initialClientId) return true;
    if (cleanNullableText(selectedClientTitle) !== cleanNullableText(initialClientState.title)) return true;
    if (serializeRepairPhotos(getRepairPhotosState()) !== serializeRepairPhotos(initialFormValues.photos)) return true;
    if (serializeRepairPhotoThumbnail(draftPhotoThumbnail) !== serializeRepairPhotoThumbnail(initialFormValues.photo_thumbnail)) return true;
    if (Boolean(draftPhotoThumbnailPendingBlob)) return true;
    if (hasFacturePaymentDirty()) return true;
    return Object.keys(buildRepairPayload(current)).length > 0;
  }

  function applyRepairDirtyIndicator() {
    const marker = $("repairModalDirtyAsterisk");
    const saveBtn = $("repairSaveBtn");
    const cancelBtn = $("repairCancelChangesBtn");
    const completeBtn = $("repairCompleteBtn");
    const handoverBtn = $("repairHandoverBtn");
    if (isCreatingRepair) {
      if (marker) marker.hidden = true;
      if (saveBtn) saveBtn.hidden = false;
      if (cancelBtn) cancelBtn.hidden = false;
      if (completeBtn) completeBtn.hidden = true;
      if (handoverBtn) handoverBtn.hidden = true;
      return;
    }
    if (marker) marker.hidden = !hasUnsavedChanges;
    if (saveBtn) saveBtn.hidden = !hasUnsavedChanges;
    if (cancelBtn) cancelBtn.hidden = !hasUnsavedChanges;
    const currentStatus = cleanNullableText(draftStatusValue) || statusFallbackLabel;
    const currentStatusKey = canonicalRepairStatus(currentStatus);
    if (completeBtn) completeBtn.hidden = currentStatusKey === "termine" || currentStatusKey === "remis_client";
    if (handoverBtn) handoverBtn.hidden = currentStatusKey !== "termine";
  }

  function refreshRepairDirtyState() {
    hasUnsavedChanges = isRepairFormDirty();
    applyRepairDirtyIndicator();
  }

  const renderRepairModalTitle = (isEditing) => {
    const selectedStatus = draftStatusValue || statusFallbackLabel;
    const tone = repairStatusToneKey(selectedStatus);
    const selectedStatusKey = canonicalRepairStatus(selectedStatus);
    const ageAlert = isCreatingRepair ? null : getRepairAgeAlertTone(repair, selectedStatusKey);
    const showFactureBadges = !isPersonalRepairValue && (selectedStatusKey === "termine" || selectedStatusKey === "remis_client");
    const facture = isCreatingRepair ? null : (getFactureForRepair(repair) || initialFacture);
    const isIssued = isYesValue(repair.fac_issued_label) || Boolean(facture);
    const resolvedPaidValue = facture ? isPaidValue(facture.etat_label) : isPaidValue(repair.fac_paid_label);
    const isPaid = draftFacturePaidValue == null ? resolvedPaidValue : Boolean(draftFacturePaidValue);
    const factureEtatDisplay = facture ? getFactureEtatDisplay(facture) : null;
    const isOverdue = factureEtatDisplay?.className === "is-overdue";
    const showOverdueInline = showFactureBadges && isIssued && !isPaid && isOverdue;
    const factureNumber = formatFactureNumber(facture);
    const issuedLabel = isIssued
      ? (factureNumber ? `Facture émise ${factureNumber}` : "Facture émise")
      : "Facture non émise";
    const canOpenIssuedFacturePdf = Boolean(facture && isIssued);
    const canCreateFacture = !isIssued && showFactureBadges && !isCreatingRepair;
    const paidLabel = isPaid ? "Payée" : "Non payée";
    const canEditFacturePayment = Boolean(facture) && isIssued && showFactureBadges && isEditing && !isCreatingRepair;
    const paidBadgeHtml = canEditFacturePayment
      ? `
          <select id="repairFacturePaidSelect" class="repair-facture-paid-select ${isPaid ? "is-paid" : "is-unpaid"}" aria-label="État de paiement de la facture">
            <option value="unpaid"${isPaid ? "" : " selected"}>Non payée</option>
            <option value="paid"${isPaid ? " selected" : ""}>Payée</option>
          </select>
        `
      : `<span class="bons-pill bons-pill-paid ${isPaid ? "is-paid" : "is-unpaid"}">${escapeHtml(paidLabel)}</span>`;
    const issuedBadgeHtml = canOpenIssuedFacturePdf
      ? `<button id="repairFactureIssuedPreviewBtn" type="button" class="bons-pill bons-pill-issued ${isIssued ? "is-issued" : "is-not-issued"} repair-facture-preview-btn" title="Voir la facture">${escapeHtml(issuedLabel)}</button>`
      : canCreateFacture
        ? `<button id="repairFactureCreateBtn" type="button" class="bons-pill bons-pill-issued is-not-issued bons-pill-create-facture-btn" title="Créer une facture">${escapeHtml(issuedLabel)}</button>`
      : `<span class="bons-pill bons-pill-issued ${isIssued ? "is-issued" : "is-not-issued"}">${escapeHtml(issuedLabel)}</span>`;
    const factureBadgesHtml = showFactureBadges
      ? `
        <span class="repair-modal-title-badges">
          ${issuedBadgeHtml}
          ${paidBadgeHtml}${showOverdueInline ? '<span class="repair-modal-overdue-inline">En retard !</span>' : ""}
        </span>
      `
      : "";

    const statusHtml = isEditing
      ? `<select id="repairStatusSelect" class="repair-status-select repair-status-select-inline repair-modal-status repair-modal-status-${escapeHtml(tone)}">${buildRepairStatusOptions(selectedStatus)}</select>`
      : `<span class="repair-modal-status repair-modal-status-${escapeHtml(tone)}">${escapeHtml(selectedStatus)}</span>`;
    const statusAgeBadgeHtml = ageAlert
      ? `<span class="repair-modal-age-badge is-${ageAlert.tone}" title="${escapeHtml(`Bon créé il y a ${ageAlert.ageDays} jour(s)`)}">${REPAIR_AGE_ALARM_SVG}</span>`
      : "";
    const dirtyHtml = isCreatingRepair
      ? ""
      : `<span id="repairModalDirtyAsterisk" class="repair-modal-dirty" title="Modifications non sauvegardées" aria-label="Modifications non sauvegardées"${hasUnsavedChanges ? "" : " hidden"}>*</span>`;
    const personalLabelHtml = isPersonalRepairValue
      ? `<span class="modal-personal-kicker">Réparation personnelle</span>`
      : "";

    title.innerHTML = isCreatingRepair
      ? `
        <span class="repair-modal-create-kicker">Création d'un nouveau bon de réparation</span>
        ${personalLabelHtml}
        <span class="repair-modal-create-main">${statusHtml}${factureBadgesHtml}</span>
      `
      : `
        ${personalLabelHtml}
        <span class="repair-modal-title-mainline">
          <span class="repair-modal-title-code">${escapeHtml(repCodeLabel)}</span>
          <span class="repair-modal-title-sep">|</span>
          ${statusHtml}${statusAgeBadgeHtml}${dirtyHtml}
        </span>
        ${factureBadgesHtml}
      `;

    const statusSelect = $("repairStatusSelect");
    if (statusSelect) {
      statusSelect.addEventListener("change", () => {
        const selected = cleanNullableText(statusSelect.value);
        if (selected) draftStatusValue = selected;
        const nextTone = repairStatusToneKey(statusSelect.value);
        statusSelect.className = `repair-status-select repair-status-select-inline repair-modal-status repair-modal-status-${nextTone}`;
        refreshRepairDirtyState();
        renderRepairModalTitle(isEditing);
      });
    }

    const facturePaidSelect = $("repairFacturePaidSelect");
    if (facturePaidSelect) {
      facturePaidSelect.addEventListener("change", () => {
        draftFacturePaidValue = facturePaidSelect.value === "paid";
        refreshRepairDirtyState();
        renderRepairModalTitle(isEditing);
      });
    }

    const factureIssuedPreviewBtn = $("repairFactureIssuedPreviewBtn");
    if (factureIssuedPreviewBtn && facture) {
      factureIssuedPreviewBtn.addEventListener("click", async (ev) => {
        ev.preventDefault();
        ev.stopPropagation();
        try {
          await openFacturePdfPreview(facture);
        } catch (err) {
          window.alert(err?.message || "Impossible d'ouvrir le PDF.");
        }
      });
    }

    const factureCreateBtn = $("repairFactureCreateBtn");
    if (factureCreateBtn) {
      factureCreateBtn.addEventListener("click", async (ev) => {
        ev.preventDefault();
        ev.stopPropagation();
        await promptRepairFactureCreation(repair);
        renderRepairModalTitle(isEditing);
      });
    }
  };

  const createClientSectionHtml = `
      <section class="repair-plain-section repair-plain-client">
        <h4 class="repair-plain-title">Client</h4>
        <button id="repairClientAddBtn" type="button" class="repair-client-add-btn">Ajouter un client</button>
        <div id="repairCreateClientBadge" class="repair-create-client-badge" hidden>
          <span id="repairCreateClientBadgeName" class="repair-create-client-badge-name"></span>
          <span id="repairCreateClientBadgeCompany" class="repair-create-client-badge-company" hidden></span>
          <button id="repairCreateClientRemoveBtn" type="button" class="repair-create-client-badge-remove" aria-label="Retirer le client" title="Retirer le client">×</button>
        </div>
        <div id="repairClientPicker" class="repair-client-picker" hidden>
          <input id="repairClientSearchInput" class="repair-input" type="search" placeholder="Rechercher un client">
          <div id="repairClientList" class="repair-client-list"></div>
        </div>
        <p id="repairClientMessage" class="repair-message"></p>
      </section>
  `;
  const existingClientSectionHtml = `
      <section class="repair-plain-section repair-plain-client">
        <div class="repair-client-view">
          <h4 class="repair-client-name">${escapeHtml(clientName)}</h4>
          ${clientInfoHtml}
        </div>
        <div class="repair-client-edit">
          <h4 class="repair-plain-title">Client</h4>
          <button id="repairClientAddBtn" type="button" class="repair-client-add-btn">Ajouter un client</button>
          <div id="repairCreateClientBadge" class="repair-create-client-badge" hidden>
            <span id="repairCreateClientBadgeName" class="repair-create-client-badge-name"></span>
            <span id="repairCreateClientBadgeCompany" class="repair-create-client-badge-company" hidden></span>
            <button id="repairCreateClientRemoveBtn" type="button" class="repair-create-client-badge-remove" aria-label="Retirer le client" title="Retirer le client">×</button>
          </div>
          <div id="repairClientPicker" class="repair-client-picker" hidden>
            <input id="repairClientSearchInput" class="repair-input" type="search" placeholder="Rechercher un client">
            <div id="repairClientList" class="repair-client-list"></div>
          </div>
          <p id="repairClientMessage" class="repair-message"></p>
        </div>
      </section>
  `;

  body.innerHTML = `
    <form id="repairEditForm" class="repair-clean-form">
      ${isCreatingRepair ? createClientSectionHtml : existingClientSectionHtml}

      <section class="repair-plain-section">
        <label class="personal-work-toggle">
          <input id="repairPersonalInput" type="checkbox" ${initialIsPersonalRepair ? "checked" : ""}>
          <span>Réparation personnelle</span>
        </label>
      </section>

      <section class="repair-plain-section">
        <h4 class="repair-plain-title repair-device-title">Appareil</h4>
        <p class="repair-device-view">
          <span class="repair-device-main">${escapeHtml(deviceMainLabel)}</span>${hasSerialValue ? ` <span class="repair-device-sn">(SN: ${escapeHtml(serialValue)})</span>` : ""}
        </p>
        <div class="repair-device-edit">
          <div class="repair-device-grid repair-device-edit-grid">
            <div>
              <label class="repair-plain-title" for="repairManufacturerInput">Marque</label>
              <input id="repairManufacturerInput" class="repair-input" type="text" value="${escapeHtml(marqueValue === "-" ? "" : marqueValue)}">
            </div>
            <div>
              <label class="repair-plain-title" for="repairModelInput">Modèle</label>
              <input id="repairModelInput" class="repair-input" type="text" value="${escapeHtml(modeleValue === "-" ? "" : modeleValue)}">
            </div>
            <div>
              <label class="repair-plain-title" for="repairSerialInput">Numéro de série</label>
              <input id="repairSerialInput" class="repair-input" type="text" value="${escapeHtml(serialValue === "-" ? "" : serialValue)}">
            </div>
          </div>
        </div>
      </section>

      <section class="repair-plain-section">
        <input id="repairPhotosInput" type="file" accept="image/*" multiple hidden>
        <div class="repair-photos-columns">
          <div class="repair-photos-column">
            <div class="repair-photos-head">
              <p class="repair-plain-title">Photos avant</p>
              <button id="repairPhotosAddBeforeBtn" type="button" class="repair-client-add-btn repair-photos-add-btn">Ajouter</button>
            </div>
            <div id="repairPhotosBeforeGrid" class="repair-photos-grid"></div>
          </div>
          <div class="repair-photos-column">
            <div class="repair-photos-head">
              <p class="repair-plain-title">Photos après</p>
              <button id="repairPhotosAddAfterBtn" type="button" class="repair-client-add-btn repair-photos-add-btn">Ajouter</button>
            </div>
            <div id="repairPhotosAfterGrid" class="repair-photos-grid"></div>
          </div>
        </div>
        <p id="repairPhotosMessage" class="repair-message"></p>
      </section>

      <section class="repair-plain-section">
        <label class="repair-plain-title" for="repairProblemInput">Description du problème</label>
        <p class="repair-problem-view">${escapeHtml(htmlToText(repair.desc_problem) || "-")}</p>
        <textarea id="repairProblemInput" class="repair-textarea repair-textarea-auto repair-problem-edit" rows="1">${escapeHtml(htmlToText(repair.desc_problem) || "")}</textarea>
      </section>

      <section class="repair-plain-section">
        <label class="repair-plain-title" for="repairObservationInput">Observation après analyse</label>
        <textarea id="repairObservationInput" class="repair-textarea repair-textarea-auto" rows="1">${escapeHtml(htmlToText(repair.observations) || "")}</textarea>
      </section>

      <section class="repair-plain-section">
        <label class="repair-plain-title" for="repairWorkDoneInput">Travail effectué</label>
        <textarea id="repairWorkDoneInput" class="repair-textarea repair-textarea-auto" rows="1">${escapeHtml(htmlToText(repair.desc_done) || "")}</textarea>
      </section>

      <section class="repair-plain-section">
        <div id="repairCostLinesSection" class="repair-cost-lines-block">
          <div class="repair-parts-cost-head">
            <p class="repair-plain-title">Lignes de facture</p>
            <button id="repairPartsShopBtn" type="button" class="repair-parts-shop-btn" title="Liens d'achats de pièces" aria-label="Ouvrir les liens d'achats de pièces">${REPAIR_PARTS_CART_SVG}</button>
          </div>
          <div id="repairCostLinesView" class="repair-cost-lines-view"></div>
          <div id="repairCostLinesEdit" class="repair-cost-lines-edit">
            <div class="facture-lines-table-wrap repair-cost-lines-table-wrap">
              <table class="facture-lines-table repair-cost-lines-table project-cost-table">
                <thead>
                  <tr>
                    <th class="project-cost-header-desc">Description</th>
                    <th class="project-cost-header-unit">Prix unitaire</th>
                    <th class="project-cost-header-currency">Devise</th>
                    <th class="project-cost-header-qty">Qté</th>
                    <th class="project-cost-header-fx">Taux</th>
                    <th class="project-cost-header-total">Total (CAD)</th>
                  </tr>
                </thead>
                <tbody id="repairLinesTbody"></tbody>
              </table>
            </div>
          </div>
          <div class="facture-lines-actions repair-cost-lines-actions">
            <button id="repairAddLineBtn" type="button" class="facture-add-line-btn">Ajouter</button>
            <p id="repairCostLinesSubtotal" class="repair-cost-lines-subtotal"></p>
          </div>
        </div>
      </section>

      <section class="repair-plain-section project-files-section">
        <input id="repairFilesInput" type="file" multiple hidden>
        <div class="project-files-head">
          <h4 class="repair-plain-title">Fichiers</h4>
        </div>
        <div id="repairFilesList" class="project-files-list"></div>
        <div class="repair-files-actions">
          <button id="repairFilesAddBtn" type="button" class="repair-client-add-btn repair-photos-add-btn">Ajouter</button>
        </div>
        <p id="repairFilesMessage" class="repair-message"></p>
      </section>

      <section class="repair-plain-section">
        <p class="repair-created-at">Créé le ${escapeHtml(createdAt)}</p>
      </section>

      <div class="repair-form-actions">
        <p id="repairSaveMessage" class="repair-message"></p>
        <div class="repair-form-actions-buttons">
          <button id="repairCancelChangesBtn" type="button" class="repair-btn" hidden>Annuler</button>
          <button id="repairSaveBtn" type="submit" class="repair-btn repair-btn-primary" hidden>Enregistrer</button>
          <button id="repairCompleteBtn" type="button" class="repair-btn repair-btn-primary">Terminé</button>
          <button id="repairHandoverBtn" type="button" class="repair-btn repair-btn-primary" hidden>Appareil remis</button>
        </div>
      </div>
    </form>
  `;

  const clientEditBtn = $("repairModalClientEditBtn");
  const clientAddBtn = $("repairClientAddBtn");
  const createClientBadge = $("repairCreateClientBadge");
  const createClientBadgeName = $("repairCreateClientBadgeName");
  const createClientBadgeCompany = $("repairCreateClientBadgeCompany");
  const createClientBadgeRemoveBtn = $("repairCreateClientRemoveBtn");
  const clientPicker = $("repairClientPicker");
  const clientSearchInput = $("repairClientSearchInput");
  const clientList = $("repairClientList");
  const clientMessage = $("repairClientMessage");
  const personalInput = $("repairPersonalInput");
  const personalSection = personalInput?.closest?.(".repair-plain-section") || null;
  const photosBeforeGrid = $("repairPhotosBeforeGrid");
  const photosAfterGrid = $("repairPhotosAfterGrid");
  const photosInput = $("repairPhotosInput");
  const photosAddBeforeBtn = $("repairPhotosAddBeforeBtn");
  const photosAddAfterBtn = $("repairPhotosAddAfterBtn");
  const photosMessage = $("repairPhotosMessage");
  const repairFilesInput = $("repairFilesInput");
  const repairFilesAddBtn = $("repairFilesAddBtn");
  const repairFilesList = $("repairFilesList");
  const repairFilesMessage = $("repairFilesMessage");
  const repairCostLinesSection = $("repairCostLinesSection");
  const repairCostLinesView = $("repairCostLinesView");
  const repairLinesTbody = $("repairLinesTbody");
  const repairAddLineBtn = $("repairAddLineBtn");
  const repairCostLinesSubtotal = $("repairCostLinesSubtotal");
  const photoThumbInfo = $("repairPhotoThumbInfo");
  const photoThumbClearBtn = $("repairPhotoThumbClearBtn");
  const editForm = $("repairEditForm");
  const saveBtn = $("repairSaveBtn");
  const cancelChangesBtn = $("repairCancelChangesBtn");
  const completeBtn = $("repairCompleteBtn");
  const handoverBtn = $("repairHandoverBtn");
  const saveMessage = $("repairSaveMessage");
  const autoGrowTextarea = (textarea) => {
    if (!textarea) return;
    textarea.style.height = "auto";
    textarea.style.height = `${textarea.scrollHeight}px`;
  };
  const problemInput = $("repairProblemInput");
  const observationInput = $("repairObservationInput");
  const workDoneInput = $("repairWorkDoneInput");
  const refreshAutoTextareasHeight = () => {
    [problemInput, observationInput, workDoneInput].forEach((field) => autoGrowTextarea(field));
  };
  [problemInput, observationInput, workDoneInput].forEach((field) => {
    if (!field) return;
    autoGrowTextarea(field);
    field.addEventListener("input", () => autoGrowTextarea(field));
  });

  const getRepairLineAmount = (line) => roundFactureMoney(
    getProjectCostLineCadAmount(line),
  );
  const parseRepairLineNumber = (value) => {
    const raw = String(value ?? "").trim().replace(",", ".");
    const parsed = Number(raw);
    return Number.isFinite(parsed) ? parsed : 0;
  };
  const repairLaborRateOptions = (() => {
    const cfg = state.appConfig || DEFAULT_APP_CONFIG;
    return [
      {
        key: "labor_rate_diagnostic",
        label: "Taux réduit",
        rate: normalizeNonNegativeNumber(cfg.labor_rate_diagnostic, DEFAULT_APP_CONFIG.labor_rate_diagnostic),
      },
      {
        key: "labor_rate_standard",
        label: "Taux standard",
        rate: normalizeNonNegativeNumber(cfg.labor_rate_standard, DEFAULT_APP_CONFIG.labor_rate_standard),
      },
      {
        key: "labor_rate_secondary",
        label: "Taux de service accéléré",
        rate: normalizeNonNegativeNumber(cfg.labor_rate_secondary, DEFAULT_APP_CONFIG.labor_rate_secondary),
      },
      {
        key: "labor_rate_estimation",
        label: "Taux d'estimation",
        rate: normalizeNonNegativeNumber(cfg.labor_rate_estimation, DEFAULT_APP_CONFIG.labor_rate_estimation),
      },
    ];
  })();
  const repairLaborRateOptionByKey = Object.fromEntries(repairLaborRateOptions.map((opt) => [opt.key, opt]));
  const customRepairLaborRateKey = "custom";
  const getRepairLaborRateValueByKey = (key) => {
    const option = repairLaborRateOptionByKey[String(key ?? "").trim()];
    if (!option) return roundFactureMoney(getConfiguredFactureLaborRate());
    return roundFactureMoney(option.rate);
  };
  const findRepairLaborRateKeyByValue = (value) => {
    const target = roundFactureMoney(parseRepairLineNumber(value));
    for (const option of repairLaborRateOptions) {
      if (Math.abs(target - roundFactureMoney(option.rate)) < 0.005) return option.key;
    }
    return customRepairLaborRateKey;
  };
  const buildRepairLaborRateOptionsHtml = (selectedKey) => {
    const currentKey = repairLaborRateOptionByKey[selectedKey] || selectedKey === customRepairLaborRateKey
      ? selectedKey
      : DEFAULT_REPAIR_LABOR_RATE_KEY;
    const options = repairLaborRateOptions.map((option) => {
      const selected = option.key === currentKey ? "selected" : "";
      const label = `${option.label} (${formatMoneyCompact(option.rate)}/h)`;
      return `<option value="${escapeHtml(option.key)}" ${selected}>${escapeHtml(label)}</option>`;
    });
    const customSelected = currentKey === customRepairLaborRateKey ? "selected" : "";
    options.push(`<option value="${escapeHtml(customRepairLaborRateKey)}" ${customSelected}>${escapeHtml("Autre tarif")}</option>`);
    return options.join("");
  };
  const isRepairSaving = () => Boolean($("repairSaveBtn")?.disabled);
  const getRepairLinesSubtotal = () => roundFactureMoney(
    normalizeRepairInvoiceLines(draftRepairLines).reduce((sum, line) => sum + getRepairLineAmount(line), 0),
  );
  const buildRepairCurrencyOptionsHtml = (selectedCurrency) => {
    const current = normalizeProjectCostCurrency(selectedCurrency);
    return PROJECT_COST_CURRENCIES.map((currencyCode) => {
      const selected = currencyCode === current ? "selected" : "";
      return `<option value="${escapeHtml(currencyCode)}" ${selected}>${escapeHtml(currencyCode)}</option>`;
    }).join("");
  };
  const syncRepairCostLineDerivedValues = (line) => {
    if (!line || typeof line !== "object") return line;
    line.currency = normalizeProjectCostCurrency(line.currency);
    if (isRepairLaborLine(line)) {
      line.description = "Main d'oeuvre";
      line.currency = PROJECT_COST_DEFAULT_CURRENCY;
      line.exchange_rate_to_cad = 1;
      line.expense_date = null;
      line.paypal_purchase = false;
    } else if (line.currency === PROJECT_COST_DEFAULT_CURRENCY) {
      line.exchange_rate_to_cad = 1;
      line.paypal_purchase = false;
    } else {
      line.exchange_rate_to_cad = normalizeProjectCostExchangeRate(line.exchange_rate_to_cad, line.currency);
      line.paypal_purchase = normalizeProjectCostPaypalPurchase(line.paypal_purchase);
    }
    line.expense_date = normalizeProjectCostExpenseDate(line.expense_date);
    line.amount = getProjectCostLineCadAmount(line);
    return line;
  };
  const fetchAndApplyRepairCostExchangeRate = async (line, { silent = false } = {}) => {
    if (!line || typeof line !== "object") return false;
    const currency = normalizeProjectCostCurrency(line.currency);
    if (currency === PROJECT_COST_DEFAULT_CURRENCY) {
      line.exchange_rate_to_cad = 1;
      line.amount = getProjectCostLineCadAmount(line);
      return true;
    }
    clearProjectCostLineExchangeRate(line);
    const expenseDate = normalizeProjectCostExpenseDate(line.expense_date) || formatDateYmd(new Date());
    line.expense_date = expenseDate;
    const requestToken = `repair-fx-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`;
    line.__fx_request_token = requestToken;
    try {
      const nextRate = await fetchHistoricalCadExchangeRate(currency, expenseDate);
      if (line.__fx_request_token !== requestToken) return false;
      line.exchange_rate_to_cad = nextRate;
      line.amount = getProjectCostLineCadAmount(line);
      return true;
    } catch (err) {
      if (!silent && saveMessage) {
        saveMessage.textContent = err?.message || "Impossible de recuperer le taux de change.";
      }
      return false;
    } finally {
      if (line.__fx_request_token === requestToken) delete line.__fx_request_token;
    }
  };
  const buildEmptyRepairQuickCostDraft = () => ({
    description: "",
    quantity: 1,
    unit_price: 0,
    currency: PROJECT_COST_DEFAULT_CURRENCY,
    exchange_rate_to_cad: 1,
    expense_date: formatDateYmd(new Date()),
    labor_rate_key: null,
    paypal_purchase: false,
    source_kind: "repair_parts",
  });
  const persistInlineRepairCostDraft = async () => {
    if (!quickRepairCostDraft) return null;
    const draftLine = {
      description: cleanNullableText(quickRepairCostDraft.description) || "",
      quantity: roundFactureMoney(Math.max(0, parseRepairLineNumber(quickRepairCostDraft.quantity))),
      unit_price: roundFactureMoney(Math.max(0, parseRepairLineNumber(quickRepairCostDraft.unit_price))),
      currency: normalizeProjectCostCurrency(quickRepairCostDraft.currency),
      exchange_rate_to_cad: normalizeProjectCostExchangeRate(quickRepairCostDraft.exchange_rate_to_cad, quickRepairCostDraft.currency),
      expense_date: normalizeProjectCostExpenseDate(quickRepairCostDraft.expense_date) || null,
      labor_rate_key: null,
      paypal_purchase: normalizeProjectCostPaypalPurchase(quickRepairCostDraft.paypal_purchase),
      source_kind: "repair_parts",
    };
    if (!draftLine.description && draftLine.quantity <= 0 && draftLine.unit_price <= 0) {
      if (saveMessage) saveMessage.textContent = "Ajoutez une description ou un montant avant d'enregistrer la ligne.";
      renderRepairLines();
      return null;
    }
    if (draftLine.currency !== PROJECT_COST_DEFAULT_CURRENCY && !(draftLine.exchange_rate_to_cad > 0)) {
      if (saveMessage) saveMessage.textContent = "Definissez un taux CAD valide pour cette depense.";
      renderRepairLines();
      return null;
    }
    draftLine.amount = getProjectCostLineCadAmount(draftLine);
    const normalizedLines = ensureRepairInvoiceLines([...draftRepairLines, draftLine], {
      defaultLaborRate: getConfiguredFactureLaborRate(),
      hourlyRateSnapshot: repair?.hourly_rate_snapshot,
      fallbackLaborQuantity: isCreatingRepair ? 1 : 0,
      includeLabor: !isPersonalRepairValue,
    });
    const repairId = Number(repair?.id);
    if (!Number.isInteger(repairId) || repairId <= 0) {
      draftRepairLines = normalizedLines;
      initialFormValues.lines = normalizeRepairInvoiceLines(normalizedLines);
      const derived = deriveRepairLegacyValuesFromLines(normalizedLines);
      initialFormValues.duration_minutes = derived.duration_minutes;
      initialFormValues.parts_cost = derived.parts_cost;
      quickRepairCostDraft = null;
      if (saveMessage) saveMessage.textContent = "";
      renderRepairLines();
      refreshRepairDirtyState();
      return repair || true;
    }
    try {
      if (saveMessage) saveMessage.textContent = "Enregistrement de la ligne...";
      const updatedRepair = await updateRepairRecord(repairId, { lines: normalizedLines });
      repair = updatedRepair;
      draftRepairLines = ensureRepairInvoiceLines(updatedRepair?.lines, {
        durationMinutes: updatedRepair?.duration_minutes,
        partsCost: updatedRepair?.parts_cost,
        hourlyRateSnapshot: updatedRepair?.hourly_rate_snapshot,
        defaultLaborRate: getConfiguredFactureLaborRate(),
        fallbackLaborQuantity: 0,
        includeLabor: !isPersonalRepairValue,
      });
      initialFormValues.lines = normalizeRepairInvoiceLines(draftRepairLines);
      const derived = deriveRepairLegacyValuesFromLines(draftRepairLines);
      initialFormValues.duration_minutes = derived.duration_minutes;
      initialFormValues.parts_cost = derived.parts_cost;
      quickRepairCostDraft = null;
      if (saveMessage) saveMessage.textContent = "";
      renderRepairLines();
      refreshRepairDirtyState();
      renderAccueilCards();
      renderFacturesList();
      renderProjectsList();
      return updatedRepair;
    } catch (err) {
      if (saveMessage) saveMessage.textContent = err?.message || "Impossible d'enregistrer la ligne.";
      renderRepairLines();
      return null;
    }
  };
  const renderRepairLines = () => {
    const normalizedLines = ensureRepairInvoiceLines(draftRepairLines, {
      defaultLaborRate: getConfiguredFactureLaborRate(),
      hourlyRateSnapshot: repair?.hourly_rate_snapshot,
      fallbackLaborQuantity: isCreatingRepair ? 1 : 0,
      includeLabor: !isPersonalRepairValue,
    });
    draftRepairLines = normalizedLines.map((line) => syncRepairCostLineDerivedValues({ ...line }));

    if (repairCostLinesView) {
      repairCostLinesView.innerHTML = `
        <div class="repair-cost-lines-table-wrap">
          <table class="facture-lines-table repair-cost-lines-table repair-cost-lines-table-static project-cost-table">
            <thead>
              <tr>
                <th class="project-cost-header-desc">Description</th>
                <th class="project-cost-header-unit">Prix unitaire</th>
                <th class="project-cost-header-currency">Devise</th>
                <th class="project-cost-header-qty">Qté</th>
                <th class="project-cost-header-fx">Taux</th>
                <th class="project-cost-header-total">Total (CAD)</th>
              </tr>
            </thead>
            <tbody>
              ${draftRepairLines.map((line) => {
                const laborLine = isRepairLaborLine(line);
                if (laborLine) {
                  line.description = "Main d'oeuvre";
                  line.source_kind = "repair_labor";
                  const currentRateKey = String(line.labor_rate_key || "").trim();
                  const selectedRateKey = (Boolean(repairLaborRateOptionByKey[currentRateKey]) || currentRateKey === customRepairLaborRateKey)
                    ? currentRateKey
                    : findRepairLaborRateKeyByValue(line.unit_price);
                  line.labor_rate_key = selectedRateKey;
                  if (selectedRateKey !== customRepairLaborRateKey) {
                    line.unit_price = getRepairLaborRateValueByKey(selectedRateKey);
                  }
                  line.currency = PROJECT_COST_DEFAULT_CURRENCY;
                  line.exchange_rate_to_cad = 1;
                  line.expense_date = null;
                }
                const currency = normalizeProjectCostCurrency(line.currency);
                const fxStamp = formatProjectCostFxStamp(
                  currency,
                  getProjectCostLineExchangeRate(line),
                  line.expense_date,
                  { paypalPurchase: normalizeProjectCostPaypalPurchase(line.paypal_purchase) },
                );
                return `
                <tr>
                  <td class="project-cost-cell-desc"><span class="project-cost-cell-value project-cost-cell-value-left">${escapeHtml(cleanNullableText(line.description) || "Ligne")}</span></td>
                  <td class="project-cost-cell-unit"><span class="project-cost-cell-value project-cost-cell-value-right">${escapeHtml(laborLine ? `${formatMoneyCompact(line.unit_price)}/h` : roundFactureMoney(line.unit_price))}</span></td>
                  <td class="project-cost-cell-currency"><span class="project-cost-currency-static">${escapeHtml(currency)}</span></td>
                  <td class="project-cost-cell-qty"><span class="project-cost-cell-value project-cost-cell-value-right">${escapeHtml(roundFactureMoney(line.quantity))}</span></td>
                  <td class="project-cost-cell-fx"><div class="project-cost-fx-stamp">${escapeHtml(formatProjectCostFxStamp(currency, getProjectCostLineExchangeRate(line), line.expense_date, { paypalPurchase: normalizeProjectCostPaypalPurchase(line.paypal_purchase) }))}</div></td>
                  <td class="project-cost-cell-total"><span class="facture-line-total">${escapeHtml(formatMoney(getRepairLineAmount(line)))}</span></td>
                </tr>
              `;
              }).join("")}
              ${quickRepairCostDraft ? (() => {
                syncRepairCostLineDerivedValues(quickRepairCostDraft);
                const currency = normalizeProjectCostCurrency(quickRepairCostDraft.currency);
                const expenseDate = normalizeProjectCostExpenseDate(quickRepairCostDraft.expense_date);
                const exchangeRate = getProjectCostLineExchangeRate(quickRepairCostDraft);
                const paypalPurchase = normalizeProjectCostPaypalPurchase(quickRepairCostDraft.paypal_purchase);
                const fxStamp = formatProjectCostFxStamp(currency, exchangeRate, expenseDate, { paypalPurchase });
                return `
                  <tr class="project-cost-draft-row" data-repair-cost-draft-row="1">
                    <td class="project-cost-cell-desc">
                      <div class="project-cost-desc-wrap repair-cost-draft-desc-wrap">
                        <div class="project-cost-draft-actions repair-cost-draft-actions-left">
                          <button type="button" class="project-cost-draft-save-btn" data-repair-cost-draft-action="save" ${isRepairSaving() ? "disabled" : ""} aria-label="Enregistrer la ligne" title="Enregistrer">✓</button>
                          <button type="button" class="facture-line-remove" data-repair-cost-draft-action="cancel" ${isRepairSaving() ? "disabled" : ""}>×</button>
                        </div>
                        <input type="text" class="facture-input repair-cost-draft-desc-input" value="${escapeHtml(quickRepairCostDraft.description || "")}" placeholder="Description" ${isRepairSaving() ? "disabled" : ""}>
                      </div>
                    </td>
                    <td class="project-cost-cell-unit"><input type="number" min="0" step="0.01" class="facture-input repair-cost-draft-unit-input project-cost-number-input" value="${escapeHtml(roundFactureMoney(Math.max(0, parseRepairLineNumber(quickRepairCostDraft.unit_price))))}" ${isRepairSaving() ? "disabled" : ""}></td>
                    <td class="project-cost-cell-currency">
                      <div class="project-cost-currency-wrap">
                        <div class="project-cost-currency-main">
                          <select class="facture-input repair-cost-draft-currency-select" ${isRepairSaving() ? "disabled" : ""}>${buildRepairCurrencyOptionsHtml(currency)}</select>
                          <label class="project-cost-date-picker${expenseDate ? " is-set" : ""}" title="${escapeHtml(expenseDate ? `Date de depense: ${expenseDate}` : "Choisir la date de depense")}">
                            <span class="project-cost-date-btn" aria-hidden="true">${PROJECT_COST_DATE_SVG}</span>
                            <input type="date" class="project-cost-date-input repair-cost-draft-date-input" value="${escapeHtml(expenseDate)}" ${isRepairSaving() ? "disabled" : ""}>
                          </label>
                        </div>
                      </div>
                    </td>
                    <td class="project-cost-cell-qty"><input type="number" min="0" step="0.01" class="facture-input repair-cost-draft-qty-input project-cost-number-input" value="${escapeHtml(roundFactureMoney(Math.max(0, parseRepairLineNumber(quickRepairCostDraft.quantity))))}" ${isRepairSaving() ? "disabled" : ""}></td>
                    <td class="project-cost-cell-fx">
                      <div class="project-cost-fx-wrap">
                        <div class="project-cost-fx-stamp">${escapeHtml(fxStamp)}</div>
                        ${currency !== PROJECT_COST_DEFAULT_CURRENCY
                          ? buildProjectCostPaypalToggleHtml("repair-cost-draft-paypal-checkbox", paypalPurchase, isRepairSaving())
                          : ""}
                      </div>
                    </td>
                    <td class="project-cost-cell-total"><span class="facture-line-total">${escapeHtml(formatMoney(getRepairLineAmount(quickRepairCostDraft)))}</span></td>
                  </tr>
                `;
              })() : ""}
            </tbody>
          </table>
        </div>
      `;
    }

    if (repairCostLinesSubtotal) {
      repairCostLinesSubtotal.textContent = `Sous-total: ${formatMoney(getRepairLinesSubtotal())}`;
    }

    if (!repairLinesTbody) return;
    const rowsHtml = draftRepairLines.map((line, index) => {
      const laborLine = isRepairLaborLine(line);
      if (laborLine) {
        line.description = "Main d'oeuvre";
        line.source_kind = "repair_labor";
        const currentRateKey = String(line.labor_rate_key || "").trim();
        const selectedRateKey = (Boolean(repairLaborRateOptionByKey[currentRateKey]) || currentRateKey === customRepairLaborRateKey)
          ? currentRateKey
          : findRepairLaborRateKeyByValue(line.unit_price);
        line.labor_rate_key = selectedRateKey;
        if (selectedRateKey !== customRepairLaborRateKey) {
          line.unit_price = getRepairLaborRateValueByKey(selectedRateKey);
        }
        line.currency = PROJECT_COST_DEFAULT_CURRENCY;
        line.exchange_rate_to_cad = 1;
        line.expense_date = null;
      }
      const currency = normalizeProjectCostCurrency(line.currency);
      const expenseDate = normalizeProjectCostExpenseDate(line.expense_date);
      const exchangeRate = getProjectCostLineExchangeRate(line);
      const paypalPurchase = normalizeProjectCostPaypalPurchase(line.paypal_purchase);
      const fxStamp = formatProjectCostFxStamp(currency, exchangeRate, expenseDate, { paypalPurchase });
      const rateControl = laborLine
        ? (isClientEditMode
          ? `
            <div class="repair-cost-rate-wrap">
              <div class="repair-cost-rate-select-wrap">
                <span class="repair-cost-rate-collapsed">${escapeHtml(`${formatMoneyCompact(line.unit_price)}/h`)}</span>
                <select class="facture-input repair-line-rate-select" ${isRepairSaving() ? "disabled" : ""}>
                  ${buildRepairLaborRateOptionsHtml(String(line.labor_rate_key || DEFAULT_REPAIR_LABOR_RATE_KEY))}
                </select>
              </div>
              ${String(line.labor_rate_key || "") === customRepairLaborRateKey
                ? `<input type="number" min="0" step="0.01" class="facture-input repair-line-unit-input project-cost-number-input" value="${escapeHtml(roundFactureMoney(line.unit_price))}" ${isRepairSaving() ? "disabled" : ""}>`
                : ""}
            </div>
          `
          : `<span class="project-cost-cell-value project-cost-cell-value-right">${escapeHtml(`${formatMoneyCompact(line.unit_price)}/h`)}</span>`)
        : (isClientEditMode
          ? `<input type="number" min="0" step="0.01" class="facture-input repair-line-unit-input project-cost-number-input" value="${escapeHtml(roundFactureMoney(line.unit_price))}" ${isRepairSaving() ? "disabled" : ""}>`
          : `<span class="project-cost-cell-value project-cost-cell-value-right">${escapeHtml(roundFactureMoney(line.unit_price))}</span>`);
      const descriptionControl = isClientEditMode
        ? `
          <div class="project-cost-desc-wrap repair-line-desc-wrap">
            ${laborLine
              ? ""
              : `<button type="button" class="facture-line-remove repair-line-desc-remove" data-repair-line-remove="${index}" ${isRepairSaving() ? "disabled" : ""}>×</button>`}
            <input type="text" class="facture-input repair-line-desc-input" value="${escapeHtml(cleanNullableText(line.description) || (laborLine ? "Main d'oeuvre" : ""))}" ${isRepairSaving() || laborLine ? "disabled" : ""}>
          </div>
        `
        : `<span class="project-cost-cell-value project-cost-cell-value-left">${escapeHtml(cleanNullableText(line.description) || "Ligne")}</span>`;
      return `
        <tr data-repair-line-index="${index}">
          <td class="project-cost-cell-desc">${descriptionControl}</td>
          <td class="project-cost-cell-unit">${rateControl}</td>
          <td class="project-cost-cell-currency">${laborLine
            ? `<span class="project-cost-currency-static">CAD</span>`
            : (isClientEditMode
              ? `
                <div class="project-cost-currency-wrap">
                  <div class="project-cost-currency-main">
                    <select class="facture-input repair-line-currency-select" ${isRepairSaving() ? "disabled" : ""}>${buildRepairCurrencyOptionsHtml(currency)}</select>
                    <label class="project-cost-date-picker${expenseDate ? " is-set" : ""}" title="${escapeHtml(expenseDate ? `Date de depense: ${expenseDate}` : "Choisir la date de depense")}">
                      <span class="project-cost-date-btn" aria-hidden="true">${PROJECT_COST_DATE_SVG}</span>
                      <input type="date" class="project-cost-date-input repair-line-date-input" value="${escapeHtml(expenseDate)}" ${isRepairSaving() ? "disabled" : ""}>
                    </label>
                  </div>
                </div>
              `
              : `<span class="project-cost-currency-static">${escapeHtml(currency)}</span>`)}</td>
          <td class="project-cost-cell-qty">${isClientEditMode
            ? `<input type="number" min="0" step="0.01" class="facture-input repair-line-qty-input project-cost-number-input" value="${escapeHtml(roundFactureMoney(line.quantity))}" ${isRepairSaving() ? "disabled" : ""}>`
            : `<span class="project-cost-cell-value project-cost-cell-value-right">${escapeHtml(roundFactureMoney(line.quantity))}</span>`}</td>
          <td class="project-cost-cell-fx">${laborLine
            ? `<span class="project-cost-fx-empty"></span>`
            : (isClientEditMode
              ? `
                <div class="project-cost-fx-wrap">
                  <div class="project-cost-fx-stamp">${escapeHtml(fxStamp)}</div>
                  ${currency !== PROJECT_COST_DEFAULT_CURRENCY
                    ? buildProjectCostPaypalToggleHtml("repair-line-paypal-checkbox", paypalPurchase, isRepairSaving())
                    : ""}
                </div>
              `
              : `<div class="project-cost-fx-stamp">${escapeHtml(fxStamp)}</div>`)}</td>
          <td class="project-cost-cell-total"><span class="facture-line-total">${escapeHtml(formatMoney(getRepairLineAmount(line)))}</span></td>
        </tr>
      `;
    });
    if (isClientEditMode && quickRepairCostDraft) {
      syncRepairCostLineDerivedValues(quickRepairCostDraft);
      const currency = normalizeProjectCostCurrency(quickRepairCostDraft.currency);
      const expenseDate = normalizeProjectCostExpenseDate(quickRepairCostDraft.expense_date);
      const exchangeRate = getProjectCostLineExchangeRate(quickRepairCostDraft);
      const paypalPurchase = normalizeProjectCostPaypalPurchase(quickRepairCostDraft.paypal_purchase);
      const fxStamp = formatProjectCostFxStamp(currency, exchangeRate, expenseDate, { paypalPurchase });
      rowsHtml.push(`
        <tr class="project-cost-draft-row" data-repair-cost-draft-row="1">
          <td class="project-cost-cell-desc">
            <div class="project-cost-desc-wrap repair-cost-draft-desc-wrap">
              <div class="project-cost-draft-actions repair-cost-draft-actions-left">
                <button type="button" class="project-cost-draft-save-btn" data-repair-cost-draft-action="save" ${isRepairSaving() ? "disabled" : ""} aria-label="Enregistrer la ligne" title="Enregistrer">✓</button>
                <button type="button" class="facture-line-remove" data-repair-cost-draft-action="cancel" ${isRepairSaving() ? "disabled" : ""}>×</button>
              </div>
              <input type="text" class="facture-input repair-cost-draft-desc-input" value="${escapeHtml(quickRepairCostDraft.description || "")}" placeholder="Description" ${isRepairSaving() ? "disabled" : ""}>
            </div>
          </td>
          <td class="project-cost-cell-unit"><input type="number" min="0" step="0.01" class="facture-input repair-cost-draft-unit-input project-cost-number-input" value="${escapeHtml(roundFactureMoney(Math.max(0, parseRepairLineNumber(quickRepairCostDraft.unit_price))))}" ${isRepairSaving() ? "disabled" : ""}></td>
          <td class="project-cost-cell-currency">
            <div class="project-cost-currency-wrap">
              <div class="project-cost-currency-main">
                <select class="facture-input repair-cost-draft-currency-select" ${isRepairSaving() ? "disabled" : ""}>${buildRepairCurrencyOptionsHtml(currency)}</select>
                <label class="project-cost-date-picker${expenseDate ? " is-set" : ""}" title="${escapeHtml(expenseDate ? `Date de depense: ${expenseDate}` : "Choisir la date de depense")}">
                  <span class="project-cost-date-btn" aria-hidden="true">${PROJECT_COST_DATE_SVG}</span>
                  <input type="date" class="project-cost-date-input repair-cost-draft-date-input" value="${escapeHtml(expenseDate)}" ${isRepairSaving() ? "disabled" : ""}>
                </label>
              </div>
            </div>
          </td>
          <td class="project-cost-cell-qty"><input type="number" min="0" step="0.01" class="facture-input repair-cost-draft-qty-input project-cost-number-input" value="${escapeHtml(roundFactureMoney(Math.max(0, parseRepairLineNumber(quickRepairCostDraft.quantity))))}" ${isRepairSaving() ? "disabled" : ""}></td>
          <td class="project-cost-cell-fx">
            <div class="project-cost-fx-wrap">
              <div class="project-cost-fx-stamp">${escapeHtml(fxStamp)}</div>
              ${currency !== PROJECT_COST_DEFAULT_CURRENCY
                ? buildProjectCostPaypalToggleHtml("repair-cost-draft-paypal-checkbox", paypalPurchase, isRepairSaving())
                : ""}
            </div>
          </td>
          <td class="project-cost-cell-total"><span class="facture-line-total">${escapeHtml(formatMoney(getRepairLineAmount(quickRepairCostDraft)))}</span></td>
        </tr>
      `);
    }
    repairLinesTbody.innerHTML = rowsHtml.join("");
    if (repairAddLineBtn) {
      repairAddLineBtn.disabled = isRepairSaving();
      repairAddLineBtn.textContent = isClientEditMode ? "Ajouter une ligne" : "Ajouter";
    }
  };

  const getRepairPhotosState = () => {
    const stored = normalizeRepairPhotos(draftRepairPhotos);
    const pending = draftPendingPhotoEntries.map((row) => ({
      filename: row.filename,
      bucket: row.bucket || REPAIRS_PHOTOS_BUCKET,
      original_name: row.original_name || null,
      content_type: row.content_type || null,
      size: row.size == null ? null : Number(row.size),
      created_at: row.created_at || null,
      phase: normalizeRepairPhotoPhase(row.phase, "before"),
    }));
    return normalizeRepairPhotos([...stored, ...pending]);
  };

  const isPendingThumbnailSource = (row) => {
    if (!draftPhotoThumbnailPendingSource || !row) return false;
    if (draftPhotoThumbnailPendingSource.kind !== "pending") return false;
    return Number(draftPhotoThumbnailPendingSource.index) === Number(row.index);
  };

  const isStoredThumbnailSource = (row) => {
    if (!row) return false;
    if (
      draftPhotoThumbnailPendingSource
      && draftPhotoThumbnailPendingSource.kind === "stored"
      && Number(draftPhotoThumbnailPendingSource.index) === Number(row.index)
    ) {
      return true;
    }
    const currentThumb = normalizeRepairPhotoThumbnail(draftPhotoThumbnail);
    if (!currentThumb) return false;
    return (
      String(currentThumb.filename || "") === String(row.filename || "")
      && String(currentThumb.bucket || REPAIRS_PHOTOS_BUCKET) === String(row.bucket || REPAIRS_PHOTOS_BUCKET)
    );
  };

  const renderPhotoThumbnailInfo = () => {
    if (!photoThumbInfo || !photoThumbClearBtn) return;
    const currentThumb = normalizeRepairPhotoThumbnail(draftPhotoThumbnail);
    const pendingThumb = Boolean(draftPhotoThumbnailPendingBlob && draftPhotoThumbnailPendingSource);
    const canEditPhotos = Boolean(isClientEditMode);
    if (pendingThumb && draftPhotoThumbnailPendingSource) {
      const sourceLabel = draftPhotoThumbnailPendingSource.phase === "after" ? "après" : "avant";
      photoThumbInfo.textContent = `Miniature sélectionnée (photo ${sourceLabel})`;
      photoThumbClearBtn.hidden = !canEditPhotos;
      photoThumbClearBtn.disabled = !canEditPhotos;
      return;
    }
    if (currentThumb) {
      const sourceLabel = normalizeRepairPhotoPhase(currentThumb.phase, "before") === "after" ? "après" : "avant";
      photoThumbInfo.textContent = `Miniature active (photo ${sourceLabel})`;
      photoThumbClearBtn.hidden = !canEditPhotos;
      photoThumbClearBtn.disabled = !canEditPhotos;
      return;
    }
    photoThumbInfo.textContent = "Aucune miniature cartouche.";
    photoThumbClearBtn.hidden = true;
  };

  const renderRepairPhotoGroup = (phase, targetGrid) => {
    if (!targetGrid) return;
    const canEditPhotos = Boolean(isClientEditMode);
    const phaseKey = normalizeRepairPhotoPhase(phase, "before");
    const storedPhotos = normalizeRepairPhotos(draftRepairPhotos)
      .map((photo, index) => ({ ...photo, index }))
      .filter((photo) => normalizeRepairPhotoPhase(photo.phase, "before") === phaseKey);
    const pendingPhotos = draftPendingPhotoEntries
      .map((photo, index) => ({ ...photo, index }))
      .filter((photo) => normalizeRepairPhotoPhase(photo.phase, "before") === phaseKey);

    if (!storedPhotos.length && !pendingPhotos.length) {
      targetGrid.innerHTML = `<p class="repair-client-picker-empty">Aucune photo.</p>`;
      return;
    }

    const storedHtml = storedPhotos.map((photo) => {
      const key = `${photo.bucket || REPAIRS_PHOTOS_BUCKET}/${photo.filename}`;
      const previewUrl = repairPhotoPreviewUrlByFilename[key] || "";
      const name = photo.original_name || photo.filename.split("/").pop() || photo.filename;
      const isThumb = isStoredThumbnailSource(photo);
      return `
        <div class="repair-photo-item${isThumb ? " is-selected-thumb" : ""}" data-photo-kind="stored" data-photo-index="${photo.index}" data-photo-phase="${phaseKey}">
          <button type="button" class="repair-photo-thumb-btn" data-photo-open="stored" data-photo-index="${photo.index}" title="Voir la photo">
            ${previewUrl
              ? `<img class="repair-photo-thumb" src="${escapeHtml(previewUrl)}" alt="${escapeHtml(name)}">`
              : `<span class="repair-photo-thumb-placeholder" aria-hidden="true">${REPAIR_THUMB_ICON_SVG}</span>`}
          </button>
          ${canEditPhotos
            ? `
              <div class="repair-photo-actions">
                <button type="button" class="repair-photo-thumb-select-btn${isThumb ? " is-active" : ""}" data-photo-thumb="stored" data-photo-index="${photo.index}" title="Choisir comme miniature">${REPAIR_THUMB_ICON_SVG}</button>
              </div>
              <button type="button" class="repair-photo-remove-btn" data-photo-remove="stored" data-photo-index="${photo.index}">×</button>
            `
            : ""}
        </div>
      `;
    }).join("");

    const pendingHtml = pendingPhotos.map((photo) => {
      const name = photo.original_name || photo.filename;
      const isThumb = isPendingThumbnailSource(photo);
      return `
        <div class="repair-photo-item repair-photo-item-pending${isThumb ? " is-selected-thumb" : ""}" data-photo-kind="pending" data-photo-index="${photo.index}" data-photo-phase="${phaseKey}">
          <button type="button" class="repair-photo-thumb-btn" data-photo-open="pending" data-photo-index="${photo.index}" title="Voir la photo">
            ${photo.preview_url
              ? `<img class="repair-photo-thumb" src="${escapeHtml(photo.preview_url)}" alt="${escapeHtml(name)}">`
              : `<span class="repair-photo-thumb-placeholder" aria-hidden="true">${REPAIR_THUMB_ICON_SVG}</span>`}
          </button>
          ${canEditPhotos
            ? `
              <div class="repair-photo-actions">
                <button type="button" class="repair-photo-thumb-select-btn${isThumb ? " is-active" : ""}" data-photo-thumb="pending" data-photo-index="${photo.index}" title="Choisir comme miniature">${REPAIR_THUMB_ICON_SVG}</button>
              </div>
              <button type="button" class="repair-photo-remove-btn" data-photo-remove="pending" data-photo-index="${photo.index}">×</button>
            `
            : ""}
        </div>
      `;
    }).join("");

    targetGrid.innerHTML = `${storedHtml}${pendingHtml}`;
  };

  const renderRepairPhotos = () => {
    const canEditPhotos = Boolean(isClientEditMode);
    if (photosAddBeforeBtn) {
      photosAddBeforeBtn.hidden = !canEditPhotos;
      photosAddBeforeBtn.disabled = !canEditPhotos;
    }
    if (photosAddAfterBtn) {
      photosAddAfterBtn.hidden = !canEditPhotos;
      photosAddAfterBtn.disabled = !canEditPhotos;
    }
    renderPhotoThumbnailInfo();
    renderRepairPhotoGroup("before", photosBeforeGrid);
    renderRepairPhotoGroup("after", photosAfterGrid);
  };

  const renderRepairFilesList = () => {
    if (repairFilesAddBtn) {
      repairFilesAddBtn.hidden = isCreatingRepair;
      repairFilesAddBtn.disabled = isRepairSaving() || isCreatingRepair || isClientEditMode;
    }
    if (repairFilesInput) {
      repairFilesInput.disabled = isRepairSaving() || isCreatingRepair || isClientEditMode;
    }
    if (!repairFilesList) return;
    const rows = normalizeProjectFiles(existingRepairFiles);
    if (!rows.length) {
      repairFilesList.innerHTML = `<p class="repair-client-picker-empty">${isCreatingRepair ? "Enregistrez le bon avant d'ajouter des fichiers." : "Aucun fichier."}</p>`;
      return;
    }
    repairFilesList.innerHTML = rows.map((file, index) => {
      const label = cleanNullableText(file.original_name) || file.filename.split("/").pop() || "Fichier";
      const sizeLabel = Number.isFinite(Number(file.size)) && Number(file.size) > 0 ? formatBytes(Number(file.size)) : "";
      return `
        <div class="project-file-item">
          <span class="project-file-item-main">
            <span class="project-file-item-head">
              <span class="project-file-item-name">${escapeHtml(label)}</span>
              ${sizeLabel ? `<span class="project-file-item-meta">${escapeHtml(sizeLabel)}</span>` : ""}
            </span>
          </span>
          <span class="project-file-item-guide" aria-hidden="true"></span>
          <span class="project-file-item-actions">
            <button type="button" class="project-file-action-btn is-preview" data-repair-file-open="${escapeHtml(String(index))}" title="Visualiser le fichier" aria-label="Visualiser le fichier">${FACTURE_PREVIEW_EYE_SVG}</button>
            <button type="button" class="project-file-action-btn" data-repair-file-download="${escapeHtml(String(index))}" title="Télécharger le fichier" aria-label="Télécharger le fichier">${PROJECT_FILE_DOWNLOAD_SVG}</button>
          </span>
        </div>
      `;
    }).join("");
  };

  const openRepairFile = async (index) => {
    const fileRow = normalizeProjectFiles(existingRepairFiles)[index];
    if (!fileRow) return;
    const developerMode = isDeveloperModeEnabled();
    if (!developerMode && !state.authClient) throw new Error("Connexion Supabase indisponible.");
    const url = developerMode && fileRow.preview_url
      ? fileRow.preview_url
      : await getProjectFileUrl(state.authClient, fileRow);
    if (!url) throw new Error("Impossible d'ouvrir le fichier.");
    const name = cleanNullableText(fileRow.original_name) || fileRow.filename.split("/").pop() || "Fichier";
    const contentType = String(fileRow.content_type || "").toLowerCase();
    const isPdf = contentType.includes("pdf") || /\.pdf$/i.test(String(fileRow.filename || ""));
    const isImage = contentType.startsWith("image/") || /\.(png|jpe?g|gif|webp|bmp|svg)$/i.test(String(fileRow.filename || ""));
    if (isPdf) {
      const openedInModal = await openFacturePdfModal(url, `Apercu PDF - ${name}`);
      if (openedInModal) return;
      const opened = openPdfInNewTab(url);
      if (opened) return;
      throw new Error("Impossible d'ouvrir ce PDF.");
    }
    if (isImage) {
      openProjectPhotoPreviewByUrl(url, name);
      return;
    }
    window.open(url, "_blank", "noopener,noreferrer");
  };

  const downloadRepairFile = async (index) => {
    const fileRow = normalizeProjectFiles(existingRepairFiles)[index];
    if (!fileRow) return;
    const developerMode = isDeveloperModeEnabled();
    if (!developerMode && !state.authClient) throw new Error("Connexion Supabase indisponible.");
    const fileName = cleanNullableText(fileRow.original_name) || fileRow.filename.split("/").pop() || "Fichier";
    const url = developerMode && fileRow.preview_url
      ? fileRow.preview_url
      : await getProjectFileDownloadUrl(state.authClient, fileRow);
    if (!url) throw new Error("Impossible de télécharger le fichier.");
    const anchor = document.createElement("a");
    anchor.href = url;
    anchor.download = fileName;
    anchor.rel = "noopener noreferrer";
    anchor.style.display = "none";
    document.body.appendChild(anchor);
    anchor.click();
    anchor.remove();
  };

  const saveRepairFiles = async (nextFiles, successMessage = "Fichier ajouté.") => {
    if (isRepairSaving() || isCreatingRepair || !repair?.id) return false;
    const normalizedFiles = normalizeProjectFiles(nextFiles);
    try {
      if (repairFilesMessage) repairFilesMessage.textContent = "Enregistrement des fichiers...";
      const updatedRepair = await updateRepairRecord(repair.id, { files: normalizedFiles });
      repair = updatedRepair;
      existingRepairFiles = normalizeProjectFiles(updatedRepair?.files);
      if (repairFilesMessage) repairFilesMessage.textContent = "";
      renderRepairFilesList();
      return true;
    } catch (err) {
      if (repairFilesMessage) repairFilesMessage.textContent = err?.message || "Impossible d'enregistrer les fichiers.";
      renderRepairFilesList();
      return false;
    }
  };

  const hydrateStoredRepairPhotoPreviews = async () => {
    if ((!photosBeforeGrid && !photosAfterGrid) || !state.authClient) return;
    const storedPhotos = normalizeRepairPhotos(draftRepairPhotos);
    if (!storedPhotos.length) return;
    const client = state.authClient;

    for (const photo of storedPhotos) {
      const key = `${photo.bucket || REPAIRS_PHOTOS_BUCKET}/${photo.filename}`;
      if (repairPhotoPreviewUrlByFilename[key]) continue;
      try {
        const previewUrl = await getRepairPhotoUrl(client, photo);
        repairPhotoPreviewUrlByFilename[key] = previewUrl;
      } catch {
        // ignore preview errors (photo stays clickable by name)
      }
    }
    renderRepairPhotos();
  };

  const openRepairPhotoPreviewByUrl = (url, titleText = "Photo de réparation") => {
    const safeUrl = cleanNullableText(url);
    if (!safeUrl) return;
    openImagePreviewGallery([{ title: titleText, url: safeUrl }], 0, titleText);
  };

  const openRepairThumbnailCropper = (imageUrl, titleText = "Miniature") => new Promise((resolve) => {
    const ratio = getBonsCardAspectRatio();
    const viewportWidth = Math.min(window.innerWidth - 120, 780);
    const viewportHeight = Math.max(160, Math.round(viewportWidth / ratio));
    const root = document.createElement("div");
    root.className = "repair-thumb-crop-overlay";
    root.innerHTML = `
      <div class="repair-thumb-crop-panel" role="dialog" aria-modal="true" aria-label="Choisir la miniature">
        <div class="repair-thumb-crop-head">
          <p class="repair-thumb-crop-title">${escapeHtml(titleText)}</p>
          <button type="button" class="repair-thumb-crop-close" aria-label="Fermer">×</button>
        </div>
        <p class="repair-thumb-crop-help">Déplace l'image pour cadrer la miniature.</p>
        <div class="repair-thumb-crop-stage-wrap">
          <div class="repair-thumb-crop-stage" style="--crop-stage-w:${Math.round(viewportWidth)}px;--crop-stage-h:${Math.round(viewportHeight)}px;">
            <img class="repair-thumb-crop-image" crossorigin="anonymous" src="${escapeHtml(imageUrl)}" alt="${escapeHtml(titleText)}">
            <span class="repair-thumb-crop-frame" aria-hidden="true"></span>
          </div>
        </div>
        <div class="repair-thumb-crop-zoom-row">
          <label for="repairThumbCropZoomInput">Zoom</label>
          <input id="repairThumbCropZoomInput" class="repair-thumb-crop-zoom" type="range" min="1" max="3" step="0.01" value="1">
        </div>
        <p class="repair-thumb-crop-error" hidden></p>
        <div class="repair-thumb-crop-actions">
          <button type="button" class="repair-btn" data-crop-cancel>Annuler</button>
          <button type="button" class="repair-btn repair-btn-primary" data-crop-apply>Appliquer</button>
        </div>
      </div>
    `;
    const img = root.querySelector(".repair-thumb-crop-image");
    const stage = root.querySelector(".repair-thumb-crop-stage");
    const zoomInput = root.querySelector("#repairThumbCropZoomInput");
    const cropError = root.querySelector(".repair-thumb-crop-error");
    const applyBtn = root.querySelector("[data-crop-apply]");
    if (!(img instanceof HTMLImageElement) || !(stage instanceof HTMLElement) || !(zoomInput instanceof HTMLInputElement)) {
      resolve(null);
      return;
    }
    const setCropError = (message = "") => {
      if (!(cropError instanceof HTMLElement)) return;
      const text = cleanNullableText(message);
      cropError.hidden = !text;
      cropError.textContent = text || "";
    };
    let naturalW = 0;
    let naturalH = 0;
    let zoom = 1;
    let offsetX = 0;
    let offsetY = 0;
    let dragging = false;
    let dragStartX = 0;
    let dragStartY = 0;
    let dragOriginX = 0;
    let dragOriginY = 0;

    const computeScaled = () => {
      const baseScale = Math.min(viewportWidth / naturalW, viewportHeight / naturalH);
      const scale = baseScale * zoom;
      return {
        width: naturalW * scale,
        height: naturalH * scale,
      };
    };
    const clampOffsets = () => {
      const scaled = computeScaled();
      const maxX = Math.max(0, (scaled.width - viewportWidth) / 2);
      const maxY = Math.max(0, (scaled.height - viewportHeight) / 2);
      offsetX = Math.min(maxX, Math.max(-maxX, offsetX));
      offsetY = Math.min(maxY, Math.max(-maxY, offsetY));
    };
    const renderTransform = () => {
      if (!naturalW || !naturalH) return;
      clampOffsets();
      const scaled = computeScaled();
      img.style.width = `${scaled.width}px`;
      img.style.height = `${scaled.height}px`;
      img.style.left = `${Math.round((viewportWidth - scaled.width) / 2 + offsetX)}px`;
      img.style.top = `${Math.round((viewportHeight - scaled.height) / 2 + offsetY)}px`;
    };
    const close = (result = null) => {
      document.removeEventListener("keydown", onKeyDown, true);
      root.remove();
      resolve(result);
    };
    const onKeyDown = (ev) => {
      if (ev.key !== "Escape") return;
      ev.preventDefault();
      ev.stopPropagation();
      close(null);
    };

    const applyCrop = async () => {
      setCropError("");
      if (applyBtn instanceof HTMLButtonElement) applyBtn.disabled = true;
      try {
        if (!naturalW || !naturalH) {
          setCropError("Image non chargée. Réessaie dans une seconde.");
          return;
        }
        const outW = REPAIR_THUMBNAIL_OUTPUT_WIDTH;
        const outH = Math.max(120, Math.round(outW / ratio));
        const canvas = document.createElement("canvas");
        canvas.width = outW;
        canvas.height = outH;
        const ctx = canvas.getContext("2d");
        if (!ctx) {
          setCropError("Impossible de générer la miniature.");
          return;
        }
        const baseScale = Math.min(viewportWidth / naturalW, viewportHeight / naturalH);
        const scale = baseScale * zoom;
        const dxStage = (viewportWidth - naturalW * scale) / 2 + offsetX;
        const dyStage = (viewportHeight - naturalH * scale) / 2 + offsetY;
        const ratioX = outW / viewportWidth;
        const ratioY = outH / viewportHeight;
        const drawX = dxStage * ratioX;
        const drawY = dyStage * ratioY;
        const drawW = naturalW * scale * ratioX;
        const drawH = naturalH * scale * ratioY;
        ctx.clearRect(0, 0, outW, outH);
        ctx.drawImage(img, drawX, drawY, drawW, drawH);
        const blob = await new Promise((resolveBlob) => canvas.toBlob(resolveBlob, "image/jpeg", 0.92));
        if (!(blob instanceof Blob)) {
          setCropError("Impossible d'appliquer le cadrage de la miniature.");
          return;
        }
        close({ blob });
      } catch (err) {
        const detail = String(err?.message || err || "");
        if (/tainted|security|cross-origin/i.test(detail)) {
          setCropError("Cette photo ne peut pas être cadrée (CORS). Vérifie les règles Storage.");
        } else {
          setCropError("Impossible d'appliquer le cadrage.");
        }
      } finally {
        if (applyBtn instanceof HTMLButtonElement) applyBtn.disabled = false;
      }
    };

    img.addEventListener("load", () => {
      naturalW = Number(img.naturalWidth || 0);
      naturalH = Number(img.naturalHeight || 0);
      renderTransform();
    }, { once: true });
    if (img.complete) {
      naturalW = Number(img.naturalWidth || 0);
      naturalH = Number(img.naturalHeight || 0);
      renderTransform();
    }

    zoomInput.addEventListener("input", () => {
      zoom = Math.max(1, Math.min(3, Number(zoomInput.value || 1)));
      renderTransform();
    });

    stage.addEventListener("pointerdown", (ev) => {
      dragging = true;
      dragStartX = ev.clientX;
      dragStartY = ev.clientY;
      dragOriginX = offsetX;
      dragOriginY = offsetY;
      stage.setPointerCapture?.(ev.pointerId);
    });
    stage.addEventListener("pointermove", (ev) => {
      if (!dragging) return;
      offsetX = dragOriginX + (ev.clientX - dragStartX);
      offsetY = dragOriginY + (ev.clientY - dragStartY);
      renderTransform();
    });
    const stopDrag = () => { dragging = false; };
    stage.addEventListener("pointerup", stopDrag);
    stage.addEventListener("pointercancel", stopDrag);
    stage.addEventListener("pointerleave", stopDrag);

    root.addEventListener("click", async (ev) => {
      if (ev.target === root || ev.target?.closest?.(".repair-thumb-crop-close") || ev.target?.closest?.("[data-crop-cancel]")) {
        close(null);
        return;
      }
      if (ev.target?.closest?.("[data-crop-apply]")) {
        ev.preventDefault();
        ev.stopPropagation();
        await applyCrop();
      }
    });
    document.addEventListener("keydown", onKeyDown, true);
    document.body.appendChild(root);
  });

  const setPendingThumbnailSelection = (blob, source) => {
    if (draftPhotoThumbnailPendingPreviewUrl) {
      revokeRepairPhotoObjectUrl(draftPhotoThumbnailPendingPreviewUrl);
      draftPhotoThumbnailPendingPreviewUrl = null;
    }
    draftPhotoThumbnailPendingBlob = blob instanceof Blob ? blob : null;
    draftPhotoThumbnailPendingSource = source || null;
    if (draftPhotoThumbnailPendingBlob) {
      draftPhotoThumbnailPendingPreviewUrl = URL.createObjectURL(draftPhotoThumbnailPendingBlob);
      repairPhotoObjectUrls.add(draftPhotoThumbnailPendingPreviewUrl);
    }
  };

  const getStoredRepairPhotoPreviewUrl = async (photo) => {
    if (!photo) throw new Error("Photo introuvable.");
    const key = `${photo.bucket || REPAIRS_PHOTOS_BUCKET}/${photo.filename}`;
    let previewUrl = repairPhotoPreviewUrlByFilename[key];
    if (!previewUrl) {
      if (!state.authClient) throw new Error("Connexion Supabase indisponible.");
      previewUrl = await getRepairPhotoUrl(state.authClient, photo);
      repairPhotoPreviewUrlByFilename[key] = previewUrl;
      renderRepairPhotos();
    }
    return previewUrl;
  };

  const buildRepairPhotoGalleryEntries = (phase = null) => {
    const phaseFilter = phase ? normalizeRepairPhotoPhase(phase, "before") : null;
    const storedEntries = normalizeRepairPhotos(draftRepairPhotos)
      .map((photo, index) => ({ ...photo, index }))
      .filter((photo) => !phaseFilter || normalizeRepairPhotoPhase(photo.phase, "before") === phaseFilter)
      .map((photo) => ({
        kind: "stored",
        index: photo.index,
        title: photo.original_name || photo.filename.split("/").pop() || "Photo",
        getUrl: () => getStoredRepairPhotoPreviewUrl(photo),
      }));
    const pendingEntries = draftPendingPhotoEntries
      .map((photo, index) => ({ ...photo, index }))
      .filter((photo) => !phaseFilter || normalizeRepairPhotoPhase(photo.phase, "before") === phaseFilter)
      .map((photo) => ({
        kind: "pending",
        index: photo.index,
        title: photo.original_name || photo.filename || "Photo",
        getUrl: () => Promise.resolve(cleanNullableText(photo.preview_url)),
      }));
    return [...storedEntries, ...pendingEntries];
  };

  const openRepairPhotoGallery = async (kind, index, phase = null) => {
    const entries = buildRepairPhotoGalleryEntries(phase);
    const startIndex = entries.findIndex((entry) => entry.kind === kind && entry.index === index);
    if (startIndex < 0) return;
    if (photosMessage) photosMessage.textContent = "";
    openImagePreviewGallery(entries, startIndex, "Photos de réparation");
  };

  const selectStoredThumbnail = async (index) => {
    const photo = normalizeRepairPhotos(draftRepairPhotos)[index];
    if (!photo) return;
    const key = `${photo.bucket || REPAIRS_PHOTOS_BUCKET}/${photo.filename}`;
    let previewUrl = repairPhotoPreviewUrlByFilename[key];
    if (!previewUrl) {
      if (!state.authClient) {
        if (photosMessage) photosMessage.textContent = "Connexion Supabase indisponible.";
        return;
      }
      try {
        previewUrl = await getRepairPhotoUrl(state.authClient, photo);
        repairPhotoPreviewUrlByFilename[key] = previewUrl;
      } catch (err) {
        if (photosMessage) photosMessage.textContent = err?.message || "Impossible de charger cette photo.";
        return;
      }
    }
    const label = photo.original_name || photo.filename.split("/").pop() || "Miniature";
    const cropped = await openRepairThumbnailCropper(previewUrl, label);
    if (!cropped?.blob) {
      if (cropped?.error && photosMessage) photosMessage.textContent = String(cropped.error);
      return;
    }
    setPendingThumbnailSelection(cropped.blob, {
      kind: "stored",
      index,
      filename: photo.filename,
      bucket: photo.bucket || REPAIRS_PHOTOS_BUCKET,
      phase: normalizeRepairPhotoPhase(photo.phase, "before"),
    });
    if (photosMessage) photosMessage.textContent = "Miniature prête à être enregistrée.";
    renderRepairPhotos();
    refreshRepairDirtyState();
  };

  const selectPendingThumbnail = async (index) => {
    const pending = draftPendingPhotoEntries[index];
    if (!pending?.preview_url) return;
    const label = pending.original_name || pending.filename || "Miniature";
    const cropped = await openRepairThumbnailCropper(pending.preview_url, label);
    if (!cropped?.blob) {
      if (cropped?.error && photosMessage) photosMessage.textContent = String(cropped.error);
      return;
    }
    setPendingThumbnailSelection(cropped.blob, {
      kind: "pending",
      index,
      filename: pending.filename,
      bucket: pending.bucket || REPAIRS_PHOTOS_BUCKET,
      phase: normalizeRepairPhotoPhase(pending.phase, "before"),
    });
    if (photosMessage) photosMessage.textContent = "Miniature prête à être enregistrée.";
    renderRepairPhotos();
    refreshRepairDirtyState();
  };

  const renderClientPickerList = (query = "") => {
    if (!clientList) return;
    const q = normalizeText(query);
    const filtered = clientPickerRows.filter((row) => {
      if (!q) return true;
      return normalizeText(`${row.name} ${row.company ?? ""}`).includes(q);
    });
    const createClientRow = `
      <button type="button" class="repair-client-picker-item repair-client-picker-item-create" data-client-action="create">
        <span class="repair-client-picker-name">+ Créer un nouveau client</span>
      </button>
    `;

    if (!filtered.length) {
      clientList.innerHTML = `${createClientRow}<p class="repair-client-picker-empty">Aucun client trouvé.</p>`;
      return;
    }

    clientList.innerHTML = `${createClientRow}${filtered.map((row) => {
      const isCurrent = String(selectedClientId) === row.id;
      return `
        <button type="button" class="repair-client-picker-item${isCurrent ? " is-current" : ""}" data-client-id="${escapeHtml(row.id)}">
          <span class="repair-client-picker-name">${escapeHtml(row.name)}</span>${row.company ? `<span class="repair-client-picker-company"> (${escapeHtml(row.company)})</span>` : ""}
        </button>
      `;
    }).join("")}`;
  };

  const renderCreateClientBadge = () => {
    if (!createClientBadge || !createClientBadgeName || !createClientBadgeCompany) return;
    const hasClient = hasCurrentClientLink();
    if (clientAddBtn) {
      clientAddBtn.hidden = hasClient || isPersonalRepairValue;
      clientAddBtn.disabled = hasClient || isPersonalRepairValue;
    }
    createClientBadge.hidden = !hasClient;
    if (createClientBadgeRemoveBtn) {
      createClientBadgeRemoveBtn.hidden = !hasClient || isPersonalRepairValue;
      createClientBadgeRemoveBtn.disabled = isPersonalRepairValue;
    }
    if (!hasClient) {
      createClientBadgeName.textContent = "";
      createClientBadgeCompany.hidden = true;
      createClientBadgeCompany.textContent = "";
      return;
    }
    createClientBadgeName.textContent = cleanNullableText(selectedClientTitle) || "Sans client";
    const company = cleanNullableText(selectedClientCompany);
    createClientBadgeCompany.hidden = !company;
    createClientBadgeCompany.textContent = company ? `(${company})` : "";
  };

  const setClientPickerOpen = (open) => {
    if (!clientPicker) return;
    const canOpen = isClientEditMode && !hasCurrentClientLink() && !isPersonalRepairValue;
    const next = Boolean(open) && canOpen;
    clientPicker.hidden = !next;
    if (next) {
      renderClientPickerList(clientSearchInput?.value ?? "");
      if (clientSearchInput) {
        clientSearchInput.focus();
        clientSearchInput.select();
      }
    }
  };

  let isClientEditMode = isCreatingRepair ? true : Boolean(options?.startInEditMode);
  const syncClientEditMode = () => {
    modal.classList.toggle("is-client-editing", isClientEditMode);
    renderRepairModalTitle(isClientEditMode);
    renderRepairLines();
    renderRepairPhotos();
    renderRepairFilesList();
    requestAnimationFrame(refreshAutoTextareasHeight);
    if (personalSection instanceof HTMLElement) {
      const showPersonalToggle = isCreatingRepair || isClientEditMode;
      personalSection.hidden = !showPersonalToggle;
      personalSection.style.display = showPersonalToggle ? "" : "none";
    }
    if (personalInput) {
      personalInput.checked = isPersonalRepairValue;
      personalInput.disabled = isRepairSaving() || (!isCreatingRepair && !isClientEditMode);
    }
    if (clientEditBtn) {
      clientEditBtn.hidden = isCreatingRepair;
      clientEditBtn.classList.toggle("is-active", isClientEditMode);
      clientEditBtn.setAttribute("aria-pressed", String(isClientEditMode));
    }
    if (!isClientEditMode) setClientPickerOpen(false);
  };
  if (clientEditBtn && !isCreatingRepair) {
    clientEditBtn.onclick = () => {
      isClientEditMode = !isClientEditMode;
      syncClientEditMode();
    };
  }
  syncClientEditMode();
  renderCreateClientBadge();

  const setClientControlsDisabled = (disabled) => {
    if (clientAddBtn) clientAddBtn.disabled = disabled || hasCurrentClientLink() || isPersonalRepairValue;
    if (createClientBadgeRemoveBtn) createClientBadgeRemoveBtn.disabled = disabled || isPersonalRepairValue;
    if (clientSearchInput) clientSearchInput.disabled = disabled || isPersonalRepairValue;
    if (clientList) {
      clientList.querySelectorAll("button").forEach((btn) => {
        btn.disabled = disabled;
      });
    }
  };
  setClientControlsDisabled(false);

  if (clientAddBtn) {
    clientAddBtn.addEventListener("click", () => {
      const isOpen = clientPicker ? !clientPicker.hidden : false;
      setClientPickerOpen(!isOpen);
      if (clientMessage) clientMessage.textContent = "";
    });
  }

  const openCreateClientFromRepairPicker = () => {
    openCreateClientModal({
      onCreated: (createdClient) => {
        rebuildRepairClientPickerRows();
        let selected = false;

        const createdItemId = getClientReferenceId(createdClient);
        if (Number.isInteger(createdItemId) && createdItemId !== 0) {
          selected = setSelectedRepairClientById(String(createdItemId));
        }
        if (!selected && createdClient?.id != null) {
          const createdRow = getClientById(createdClient.id);
          const createdRowItemId = getClientReferenceId(createdRow);
          if (Number.isInteger(createdRowItemId)) {
            selected = setSelectedRepairClientById(String(createdRowItemId));
          }
        }
        if (!selected) {
          const fallbackName = normalizeText(formatClientPrimaryName(createdClient));
          const fallbackCompany = normalizeText(cleanNullableText(createdClient?.company) || "");
          const fallbackRow = clientPickerRows.find((row) => (
            normalizeText(row.name) === fallbackName
            && normalizeText(row.company || "") === fallbackCompany
          ));
          if (fallbackRow) {
            selected = setSelectedRepairClientById(fallbackRow.id);
          }
        }

        renderCreateClientBadge();
        renderClientPickerList(clientSearchInput?.value ?? "");
        setClientPickerOpen(false);
        if (clientMessage) clientMessage.textContent = "";
        refreshRepairDirtyState();
      }
    });
  };

  if (clientSearchInput) {
    clientSearchInput.addEventListener("input", () => {
      renderClientPickerList(clientSearchInput.value);
    });
  }

  if (clientList && clientMessage) {
    clientList.addEventListener("click", async (ev) => {
      const actionButton = ev.target?.closest?.("button[data-client-action]");
      if (actionButton && actionButton.getAttribute("data-client-action") === "create") {
        openCreateClientFromRepairPicker();
        return;
      }
      const button = ev.target?.closest?.("button[data-client-id]");
      if (!button) return;
      const selectedId = String(button.getAttribute("data-client-id") ?? "").trim();
      if (!selectedId) return;

      const selectedClient = state.clientsByPodioItemId[selectedId];
      if (!selectedClient) {
        clientMessage.textContent = "Client introuvable.";
        return;
      }

      setSelectedRepairClientById(selectedId);
      renderCreateClientBadge();
      renderClientPickerList(clientSearchInput?.value ?? "");
      setClientPickerOpen(false);
      clientMessage.textContent = "";
      refreshRepairDirtyState();
    });
  }

  if (createClientBadgeRemoveBtn && clientMessage) {
    createClientBadgeRemoveBtn.addEventListener("click", () => {
      if (isPersonalRepairValue) return;
      if (!hasCurrentClientLink()) {
        clientMessage.textContent = "Aucun client à retirer.";
        return;
      }
      selectedClientId = "";
      selectedClientItemId = null;
      selectedClientTitle = null;
      selectedClientCompany = null;
      renderCreateClientBadge();
      renderClientPickerList(clientSearchInput?.value ?? "");
      clientMessage.textContent = "";
      refreshRepairDirtyState();
    });
  }

  const applyPersonalRepairSelection = () => {
    const personalOwner = getPersonalOwnerSelection();
    selectedClientId = personalOwner.itemId != null ? String(personalOwner.itemId) : "";
    selectedClientItemId = personalOwner.itemId;
    selectedClientTitle = personalOwner.title;
    selectedClientCompany = personalOwner.company;
  };

  if (personalInput) {
    personalInput.addEventListener("change", () => {
      isPersonalRepairValue = personalInput.checked;
      if (isPersonalRepairValue) {
        applyPersonalRepairSelection();
      }
      draftRepairLines = ensureRepairInvoiceLines(draftRepairLines, {
        durationMinutes: repair?.duration_minutes,
        partsCost: repair?.parts_cost,
        hourlyRateSnapshot: repair?.hourly_rate_snapshot,
        defaultLaborRate: getConfiguredFactureLaborRate(),
        fallbackLaborQuantity: isCreatingRepair ? 1 : 0,
        includeLabor: !isPersonalRepairValue,
      });
      renderCreateClientBadge();
      setClientPickerOpen(false);
      renderRepairLines();
      refreshRepairDirtyState();
    });
  }

  let pendingPhotoPhase = "before";
  if (photosAddBeforeBtn && photosInput) {
    photosAddBeforeBtn.addEventListener("click", () => {
      if (!isClientEditMode) return;
      pendingPhotoPhase = "before";
      photosInput.click();
    });
  }
  if (photosAddAfterBtn && photosInput) {
    photosAddAfterBtn.addEventListener("click", () => {
      if (!isClientEditMode) return;
      pendingPhotoPhase = "after";
      photosInput.click();
    });
  }

  if (photosInput) {
    photosInput.addEventListener("change", () => {
      const files = Array.from(photosInput.files || []);
      photosInput.value = "";
      if (!files.length) return;

      let addedCount = 0;
      const stamp = Date.now();
      files.forEach((file, idx) => {
        if (!(file instanceof File)) return;
        const contentType = cleanNullableText(file.type) || "application/octet-stream";
        if (!contentType.toLowerCase().startsWith("image/")) return;

        const safeName = sanitizeFilenameToken(file.name || `photo-${stamp}-${idx + 1}.jpg`, `photo-${stamp}-${idx + 1}.jpg`);
        const previewUrl = URL.createObjectURL(file);
        repairPhotoObjectUrls.add(previewUrl);
        draftPendingPhotoEntries.push({
          file,
          filename: `pending/${stamp}-${idx + 1}_${safeName}`,
          bucket: REPAIRS_PHOTOS_BUCKET,
          original_name: cleanNullableText(file.name),
          content_type: contentType,
          size: Number.isFinite(Number(file.size)) ? Number(file.size) : null,
          created_at: new Date().toISOString(),
          phase: normalizeRepairPhotoPhase(pendingPhotoPhase, "before"),
          preview_url: previewUrl,
        });
        addedCount += 1;
      });

      if (!addedCount) {
        if (photosMessage) photosMessage.textContent = "Ajoute des fichiers image valides.";
        return;
      }
      if (photosMessage) photosMessage.textContent = "";
      renderRepairPhotos();
      refreshRepairDirtyState();
    });
  }

  if (repairFilesAddBtn && repairFilesInput) {
    repairFilesAddBtn.addEventListener("click", () => {
      if (isRepairSaving() || isCreatingRepair || isClientEditMode) return;
      repairFilesInput.click();
    });
  }

  if (repairFilesInput) {
    repairFilesInput.addEventListener("change", async () => {
      if (isRepairSaving() || isCreatingRepair || isClientEditMode) return;
      const files = Array.from(repairFilesInput.files || []).filter((file) => file instanceof File);
      repairFilesInput.value = "";
      if (!files.length) return;
      const developerMode = isDeveloperModeEnabled();
      const client = state.authClient;
      if (!developerMode && !client) {
        if (repairFilesMessage) repairFilesMessage.textContent = "Connexion Supabase indisponible.";
        return;
      }
      const uploadedFiles = [];
      try {
        let seq = existingRepairFiles.length + 1;
        for (const file of files) {
          let uploaded = null;
          if (developerMode) {
            let previewUrl = "";
            try {
              previewUrl = URL.createObjectURL(file);
            } catch {
              previewUrl = "";
            }
            uploaded = {
              filename: buildRepairFileObjectPath(repair, file.name || "fichier", seq++),
              bucket: PROJECT_FILES_BUCKET,
              original_name: cleanNullableText(file.name),
              content_type: cleanNullableText(file.type),
              size: Number.isFinite(Number(file.size)) ? Number(file.size) : null,
              created_at: new Date().toISOString(),
              preview_url: previewUrl || null,
            };
          } else {
            uploaded = await uploadRepairFileAttachment(client, repair, file, seq++);
          }
          uploadedFiles.push(uploaded);
        }
      } catch (err) {
        if (repairFilesMessage) repairFilesMessage.textContent = err?.message || "Impossible de televerser le fichier.";
        return;
      }
      await saveRepairFiles(
        [...existingRepairFiles, ...uploadedFiles],
        uploadedFiles.length > 1 ? "Fichiers ajoutés." : "Fichier ajouté.",
      );
    });
  }

  if (repairFilesList) {
    repairFilesList.addEventListener("click", async (ev) => {
      if (!(ev.target instanceof Element)) return;
      const downloadBtn = ev.target.closest("[data-repair-file-download]");
      if (downloadBtn) {
        const index = Number(downloadBtn.getAttribute("data-repair-file-download"));
        if (!Number.isInteger(index) || index < 0) return;
        try {
          if (repairFilesMessage) repairFilesMessage.textContent = "";
          await downloadRepairFile(index);
        } catch (err) {
          if (repairFilesMessage) repairFilesMessage.textContent = err?.message || "Impossible de télécharger le fichier.";
        }
        return;
      }
      const openBtn = ev.target.closest("[data-repair-file-open]");
      if (!openBtn) return;
      const index = Number(openBtn.getAttribute("data-repair-file-open"));
      if (!Number.isInteger(index) || index < 0) return;
      try {
        if (repairFilesMessage) repairFilesMessage.textContent = "";
        await openRepairFile(index);
      } catch (err) {
        if (repairFilesMessage) repairFilesMessage.textContent = err?.message || "Impossible d'ouvrir le fichier.";
      }
    });
  }

  const handlePhotosGridClick = async (ev) => {
      const removeBtn = ev.target?.closest?.("[data-photo-remove]");
      if (removeBtn && isClientEditMode) {
        const kind = removeBtn.getAttribute("data-photo-remove");
        const index = Number(removeBtn.getAttribute("data-photo-index"));
        if (!Number.isInteger(index) || index < 0) return;

        if (kind === "stored") {
          const next = normalizeRepairPhotos(draftRepairPhotos);
          if (index >= next.length) return;
          const removed = next[index];
          const currentThumb = normalizeRepairPhotoThumbnail(draftPhotoThumbnail);
          if (currentThumb && removed && String(currentThumb.filename) === String(removed.filename) && String(currentThumb.bucket || REPAIRS_PHOTOS_BUCKET) === String(removed.bucket || REPAIRS_PHOTOS_BUCKET)) {
            draftPhotoThumbnail = null;
          }
          next.splice(index, 1);
          draftRepairPhotos = normalizeRepairPhotos(next);
        }
        if (kind === "pending") {
          const pending = draftPendingPhotoEntries[index];
          if (!pending) return;
          if (draftPhotoThumbnailPendingSource?.kind === "pending" && Number(draftPhotoThumbnailPendingSource.index) === index) {
            setPendingThumbnailSelection(null, null);
          }
          revokeRepairPhotoObjectUrl(pending.preview_url);
          draftPendingPhotoEntries.splice(index, 1);
          if (draftPhotoThumbnailPendingSource?.kind === "pending" && Number(draftPhotoThumbnailPendingSource.index) > index) {
            draftPhotoThumbnailPendingSource.index -= 1;
          }
        }
        if (photosMessage) photosMessage.textContent = "";
        renderRepairPhotos();
        refreshRepairDirtyState();
        return;
      }

      const thumbBtn = ev.target?.closest?.("[data-photo-thumb]");
      if (thumbBtn && isClientEditMode) {
        const thumbKind = thumbBtn.getAttribute("data-photo-thumb");
        const thumbIndex = Number(thumbBtn.getAttribute("data-photo-index"));
        if (!Number.isInteger(thumbIndex) || thumbIndex < 0) return;
        if (thumbKind === "stored") {
          await selectStoredThumbnail(thumbIndex);
          return;
        }
        if (thumbKind === "pending") {
          await selectPendingThumbnail(thumbIndex);
          return;
        }
      }
      const openBtn = ev.target?.closest?.("[data-photo-open]");
      if (!openBtn) return;
      const kind = openBtn.getAttribute("data-photo-open");
      const index = Number(openBtn.getAttribute("data-photo-index"));
      const phase = openBtn.closest("[data-photo-phase]")?.getAttribute("data-photo-phase") || null;
      if (!Number.isInteger(index) || index < 0) return;

      if (kind === "stored" || kind === "pending") {
        await openRepairPhotoGallery(kind, index, phase);
      }
  };
  if (photosBeforeGrid) photosBeforeGrid.addEventListener("click", handlePhotosGridClick);
  if (photosAfterGrid) photosAfterGrid.addEventListener("click", handlePhotosGridClick);

  if (photoThumbClearBtn) {
    photoThumbClearBtn.addEventListener("click", () => {
      if (!isClientEditMode) return;
      draftPhotoThumbnail = null;
      setPendingThumbnailSelection(null, null);
      if (photosMessage) photosMessage.textContent = "Miniature retirée.";
      renderRepairPhotos();
      refreshRepairDirtyState();
    });
  }

  const clientSection = body.querySelector(".repair-plain-client");
  const workDoneField = $("repairWorkDoneInput");
  const partsShopBtn = $("repairPartsShopBtn");
  const requiredMarkClass = "repair-required-missing";
  const clearCompletionHighlights = () => {
    [clientSection, workDoneField, repairCostLinesSection].forEach((el) => {
      if (!el) return;
      el.classList.remove(requiredMarkClass);
    });
  };
  const highlightCompletionMissing = ({ client = false, workDone = false, lines = false } = {}) => {
    clearCompletionHighlights();
    if (client && clientSection) clientSection.classList.add(requiredMarkClass);
    if (workDone && workDoneField) workDoneField.classList.add(requiredMarkClass);
    if (lines && repairCostLinesSection) repairCostLinesSection.classList.add(requiredMarkClass);
  };
  const repairDirtyFieldIds = [
    "repairProblemInput",
    "repairObservationInput",
    "repairWorkDoneInput",
    "repairManufacturerInput",
    "repairModelInput",
    "repairSerialInput",
  ];
  if (partsShopBtn) {
    partsShopBtn.addEventListener("click", () => {
      openRepairPartsSuppliersDialog();
    });
  }
  if (repairAddLineBtn) {
    repairAddLineBtn.addEventListener("click", () => {
      if (isRepairSaving() || quickRepairCostDraft) return;
      quickRepairCostDraft = buildEmptyRepairQuickCostDraft();
      renderRepairLines();
      requestAnimationFrame(() => {
        const input = body.querySelector(
          isClientEditMode ? ".repair-cost-draft-desc-input" : ".repair-cost-draft-desc-input",
        );
        if (input instanceof HTMLElement) input.focus();
      });
      refreshRepairDirtyState();
    });
  }
  if (repairLinesTbody) {
    repairLinesTbody.addEventListener("input", (ev) => {
      if (isRepairSaving() || !(ev.target instanceof Element)) return;
      if (quickRepairCostDraft) {
        const draftRow = ev.target.closest("tr[data-repair-cost-draft-row]");
        if (draftRow) {
          const target = ev.target;
          if (!(target instanceof HTMLInputElement)) return;
          if (target.classList.contains("repair-cost-draft-desc-input")) quickRepairCostDraft.description = target.value;
          if (target.classList.contains("repair-cost-draft-unit-input")) quickRepairCostDraft.unit_price = Math.max(0, parseRepairLineNumber(target.value));
          if (target.classList.contains("repair-cost-draft-qty-input")) quickRepairCostDraft.quantity = Math.max(0, parseRepairLineNumber(target.value));
          if (target.classList.contains("repair-cost-draft-fx-input")) {
            quickRepairCostDraft.exchange_rate_to_cad = normalizeProjectCostExchangeRate(target.value, quickRepairCostDraft.currency);
          }
          syncRepairCostLineDerivedValues(quickRepairCostDraft);
          const totalEl = draftRow.querySelector(".facture-line-total");
          if (totalEl) totalEl.textContent = formatMoney(getRepairLineAmount(quickRepairCostDraft));
          refreshRepairDirtyState();
          return;
        }
      }
      if (!isClientEditMode) return;
      const row = ev.target?.closest?.("tr[data-repair-line-index]");
      if (!row) return;
      const index = Number(row.getAttribute("data-repair-line-index"));
      if (!Number.isInteger(index) || index < 0) return;
      const line = draftRepairLines[index];
      if (!line) return;
      const target = ev.target;
      if (!(target instanceof HTMLInputElement)) return;
      if (target.classList.contains("repair-line-desc-input") && !isRepairLaborLine(line)) {
        line.description = target.value;
      } else if (target.classList.contains("repair-line-qty-input")) {
        line.quantity = Math.max(0, parseRepairLineNumber(target.value));
      } else if (target.classList.contains("repair-line-unit-input")) {
        line.unit_price = Math.max(0, parseRepairLineNumber(target.value));
        if (isRepairLaborLine(line)) line.labor_rate_key = customRepairLaborRateKey;
      } else if (target.classList.contains("repair-line-fx-input")) {
        line.exchange_rate_to_cad = normalizeProjectCostExchangeRate(target.value, line.currency);
      }
      syncRepairCostLineDerivedValues(line);
      const totalEl = row.querySelector(".facture-line-total");
      if (totalEl) totalEl.textContent = formatMoney(line.amount);
      refreshRepairDirtyState();
    });
    repairLinesTbody.addEventListener("change", async (ev) => {
      if (isRepairSaving() || !(ev.target instanceof Element)) return;
      if (quickRepairCostDraft) {
        const draftRow = ev.target.closest("tr[data-repair-cost-draft-row]");
        if (draftRow) {
          const target = ev.target;
          if (target instanceof HTMLSelectElement && target.classList.contains("repair-cost-draft-currency-select")) {
            quickRepairCostDraft.currency = normalizeProjectCostCurrency(target.value);
            if (quickRepairCostDraft.currency === PROJECT_COST_DEFAULT_CURRENCY) {
              quickRepairCostDraft.exchange_rate_to_cad = 1;
              syncRepairCostLineDerivedValues(quickRepairCostDraft);
            } else {
              await fetchAndApplyRepairCostExchangeRate(quickRepairCostDraft);
            }
            renderRepairLines();
            refreshRepairDirtyState();
            return;
          }
          if (target instanceof HTMLInputElement && target.classList.contains("repair-cost-draft-date-input")) {
            quickRepairCostDraft.expense_date = normalizeProjectCostExpenseDate(target.value) || formatDateYmd(new Date());
            if (normalizeProjectCostCurrency(quickRepairCostDraft.currency) !== PROJECT_COST_DEFAULT_CURRENCY) {
              await fetchAndApplyRepairCostExchangeRate(quickRepairCostDraft);
            } else {
              syncRepairCostLineDerivedValues(quickRepairCostDraft);
            }
            renderRepairLines();
            refreshRepairDirtyState();
            return;
          }
          if (target instanceof HTMLInputElement && target.classList.contains("repair-cost-draft-paypal-checkbox")) {
            quickRepairCostDraft.paypal_purchase = target.checked;
            syncRepairCostLineDerivedValues(quickRepairCostDraft);
            renderRepairLines();
            refreshRepairDirtyState();
            return;
          }
        }
      }
      if (!isClientEditMode) return;
      const row = ev.target?.closest?.("tr[data-repair-line-index]");
      if (!row) return;
      const index = Number(row.getAttribute("data-repair-line-index"));
      if (!Number.isInteger(index) || index < 0) return;
      const line = draftRepairLines[index];
      if (!line) return;
      const target = ev.target;
      if (target instanceof HTMLSelectElement && target.classList.contains("repair-line-currency-select")) {
        line.currency = normalizeProjectCostCurrency(target.value);
        if (line.currency === PROJECT_COST_DEFAULT_CURRENCY) {
          line.exchange_rate_to_cad = 1;
          syncRepairCostLineDerivedValues(line);
        } else {
          await fetchAndApplyRepairCostExchangeRate(line);
        }
        renderRepairLines();
        refreshRepairDirtyState();
        return;
      }
      if (target instanceof HTMLSelectElement && target.classList.contains("repair-line-rate-select")) {
        const nextRateKey = String(target.value ?? "").trim();
        if (!(repairLaborRateOptionByKey[nextRateKey] || nextRateKey === customRepairLaborRateKey)) return;
        line.labor_rate_key = nextRateKey;
        if (nextRateKey !== customRepairLaborRateKey) {
          line.unit_price = getRepairLaborRateValueByKey(nextRateKey);
        }
        syncRepairCostLineDerivedValues(line);
        renderRepairLines();
        refreshRepairDirtyState();
        return;
      }
      if (target instanceof HTMLInputElement && target.classList.contains("repair-line-date-input")) {
        line.expense_date = normalizeProjectCostExpenseDate(target.value) || formatDateYmd(new Date());
        if (normalizeProjectCostCurrency(line.currency) !== PROJECT_COST_DEFAULT_CURRENCY) {
          await fetchAndApplyRepairCostExchangeRate(line);
        } else {
          syncRepairCostLineDerivedValues(line);
        }
        renderRepairLines();
        refreshRepairDirtyState();
        return;
      }
      if (target instanceof HTMLInputElement && target.classList.contains("repair-line-paypal-checkbox")) {
        line.paypal_purchase = target.checked;
        syncRepairCostLineDerivedValues(line);
        renderRepairLines();
        refreshRepairDirtyState();
      }
    });
    repairLinesTbody.addEventListener("click", (ev) => {
      if (isRepairSaving() || !(ev.target instanceof Element)) return;
      const draftActionBtn = ev.target.closest("[data-repair-cost-draft-action]");
      if (draftActionBtn) {
        const action = String(draftActionBtn.getAttribute("data-repair-cost-draft-action") || "");
        if (action === "cancel") {
          quickRepairCostDraft = null;
          if (saveMessage) saveMessage.textContent = "";
          renderRepairLines();
          refreshRepairDirtyState();
          return;
        }
        if (action === "save") {
          void persistInlineRepairCostDraft();
        }
        return;
      }
      if (!isClientEditMode) return;
      const btn = ev.target?.closest?.("[data-repair-line-remove]");
      if (!btn) return;
      const index = Number(btn.getAttribute("data-repair-line-remove"));
      if (!Number.isInteger(index) || index < 0) return;
      const line = draftRepairLines[index];
      if (!line || isRepairLaborLine(line)) return;
      draftRepairLines = draftRepairLines.filter((_, lineIndex) => lineIndex !== index);
      renderRepairLines();
      refreshRepairDirtyState();
    });
  }
  if (repairCostLinesView) {
    repairCostLinesView.addEventListener("input", (ev) => {
      if (isRepairSaving() || !(ev.target instanceof Element) || !quickRepairCostDraft || isClientEditMode) return;
      const draftRow = ev.target.closest("tr[data-repair-cost-draft-row]");
      if (!draftRow) return;
      const target = ev.target;
      if (!(target instanceof HTMLInputElement)) return;
      if (target.classList.contains("repair-cost-draft-desc-input")) quickRepairCostDraft.description = target.value;
      if (target.classList.contains("repair-cost-draft-unit-input")) quickRepairCostDraft.unit_price = Math.max(0, parseRepairLineNumber(target.value));
      if (target.classList.contains("repair-cost-draft-qty-input")) quickRepairCostDraft.quantity = Math.max(0, parseRepairLineNumber(target.value));
      if (target.classList.contains("repair-cost-draft-fx-input")) {
        quickRepairCostDraft.exchange_rate_to_cad = normalizeProjectCostExchangeRate(target.value, quickRepairCostDraft.currency);
      }
      syncRepairCostLineDerivedValues(quickRepairCostDraft);
      const totalEl = draftRow.querySelector(".facture-line-total");
      if (totalEl) totalEl.textContent = formatMoney(getRepairLineAmount(quickRepairCostDraft));
      refreshRepairDirtyState();
    });
    repairCostLinesView.addEventListener("change", async (ev) => {
      if (isRepairSaving() || !(ev.target instanceof Element) || !quickRepairCostDraft || isClientEditMode) return;
      const draftRow = ev.target.closest("tr[data-repair-cost-draft-row]");
      if (!draftRow) return;
      const target = ev.target;
      if (target instanceof HTMLSelectElement && target.classList.contains("repair-cost-draft-currency-select")) {
        quickRepairCostDraft.currency = normalizeProjectCostCurrency(target.value);
        if (quickRepairCostDraft.currency === PROJECT_COST_DEFAULT_CURRENCY) {
          quickRepairCostDraft.exchange_rate_to_cad = 1;
          syncRepairCostLineDerivedValues(quickRepairCostDraft);
        } else {
          await fetchAndApplyRepairCostExchangeRate(quickRepairCostDraft);
        }
        renderRepairLines();
        refreshRepairDirtyState();
        return;
      }
      if (target instanceof HTMLInputElement && target.classList.contains("repair-cost-draft-date-input")) {
        quickRepairCostDraft.expense_date = normalizeProjectCostExpenseDate(target.value) || formatDateYmd(new Date());
        if (normalizeProjectCostCurrency(quickRepairCostDraft.currency) !== PROJECT_COST_DEFAULT_CURRENCY) {
          await fetchAndApplyRepairCostExchangeRate(quickRepairCostDraft);
        } else {
          syncRepairCostLineDerivedValues(quickRepairCostDraft);
        }
        renderRepairLines();
        refreshRepairDirtyState();
        return;
      }
      if (target instanceof HTMLInputElement && target.classList.contains("repair-cost-draft-paypal-checkbox")) {
        quickRepairCostDraft.paypal_purchase = target.checked;
        syncRepairCostLineDerivedValues(quickRepairCostDraft);
        renderRepairLines();
        refreshRepairDirtyState();
      }
    });
    repairCostLinesView.addEventListener("click", (ev) => {
      if (isRepairSaving() || !(ev.target instanceof Element) || !quickRepairCostDraft || isClientEditMode) return;
      const draftActionBtn = ev.target.closest("[data-repair-cost-draft-action]");
      if (!draftActionBtn) return;
      const action = String(draftActionBtn.getAttribute("data-repair-cost-draft-action") || "");
      if (action === "cancel") {
        quickRepairCostDraft = null;
        if (saveMessage) saveMessage.textContent = "";
        renderRepairLines();
        refreshRepairDirtyState();
        return;
      }
      if (action === "save") {
        void persistInlineRepairCostDraft();
      }
    });
  }
  for (const fieldId of repairDirtyFieldIds) {
    const field = $(fieldId);
    if (!field) continue;
    field.addEventListener("input", refreshRepairDirtyState);
    field.addEventListener("change", refreshRepairDirtyState);
  }

  const restoreInitialRepairFormValues = () => {
    const setVal = (id, value) => {
      const field = $(id);
      if (!field) return;
      field.value = value ?? "";
    };

    draftStatusValue = initialFormValues.status_label;
    const statusSelect = $("repairStatusSelect");
    if (statusSelect) {
      const target = initialFormValues.status_label ?? "";
      const hasOption = Array.from(statusSelect.options).some((opt) => normalizeText(opt.value) === normalizeText(target));
      if (hasOption) statusSelect.value = target;
      const nextTone = repairStatusToneKey(statusSelect.value || target);
      statusSelect.className = `repair-status-select repair-status-select-inline repair-modal-status repair-modal-status-${nextTone}`;
    }
    draftFacturePaidValue = initialFacturePaidValue;
    isPersonalRepairValue = initialIsPersonalRepair;
    if (personalInput) personalInput.checked = initialIsPersonalRepair;
    renderRepairModalTitle(isClientEditMode);

    setVal("repairProblemInput", initialFormValues.desc_problem ?? "");
    setVal("repairObservationInput", initialFormValues.observations ?? "");
    setVal("repairWorkDoneInput", initialFormValues.desc_done ?? "");
    setVal("repairManufacturerInput", initialFormValues.manufacturer ?? "");
    setVal("repairModelInput", initialFormValues.model ?? "");
    setVal("repairSerialInput", initialFormValues.serial ?? "");
    draftRepairLines = normalizeRepairInvoiceLines(initialFormValues.lines);
    quickRepairCostDraft = null;
    renderRepairLines();
    selectedClientItemId = initialClientState.itemId == null ? null : Number(initialClientState.itemId);
    selectedClientId = selectedClientItemId == null ? "" : String(selectedClientItemId);
    selectedClientTitle = initialClientState.title;
    selectedClientCompany = initialClientState.company;
    draftPhotoThumbnail = normalizeRepairPhotoThumbnail(initialFormValues.photo_thumbnail);
    cleanupRepairPhotoResources();
    draftRepairPhotos = normalizeRepairPhotos(initialFormValues.photos);
    renderRepairPhotos();
    void hydrateStoredRepairPhotoPreviews();
    if (photosInput) photosInput.value = "";
    if (photosMessage) photosMessage.textContent = "";
    renderCreateClientBadge();
    renderClientPickerList(clientSearchInput?.value ?? "");
    requestAnimationFrame(refreshAutoTextareasHeight);
    if (saveMessage) saveMessage.textContent = "";
    clearCompletionHighlights();
    refreshRepairDirtyState();
  };

  if (cancelChangesBtn) {
    cancelChangesBtn.addEventListener("click", () => {
      if (saveBtn?.disabled) return;
      if (isCreatingRepair) {
        closeRepairModal();
        return;
      }
      restoreInitialRepairFormValues();
    });
  }

  const buildRepairPayload = (nextValues) => {
    const payload = {};
    if (normalizeText(nextValues.status_label) !== normalizeText(initialFormValues.status_label)) payload.status_label = nextValues.status_label;
    if (nextValues.desc_problem !== initialFormValues.desc_problem) payload.desc_problem = nextValues.desc_problem;
    if (nextValues.observations !== initialFormValues.observations) payload.observations = nextValues.observations;
    if (nextValues.desc_done !== initialFormValues.desc_done) payload.desc_done = nextValues.desc_done;
    if (nextValues.manufacturer !== initialFormValues.manufacturer) payload.manufacturer = nextValues.manufacturer;
    if (nextValues.model !== initialFormValues.model) payload.model = nextValues.model;
    if (nextValues.serial !== initialFormValues.serial) payload.serial = nextValues.serial;
    if (serializeRepairInvoiceLines(nextValues.lines) !== serializeRepairInvoiceLines(initialFormValues.lines)) payload.lines = nextValues.lines;
    if (!sameNullableNumber(nextValues.duration_minutes, initialFormValues.duration_minutes)) payload.duration_minutes = nextValues.duration_minutes;
    if (!sameNullableNumber(nextValues.parts_cost, initialFormValues.parts_cost)) payload.parts_cost = nextValues.parts_cost;
    if (Boolean(nextValues.is_personal) !== Boolean(initialFormValues.is_personal)) payload.is_personal = Boolean(nextValues.is_personal);
    const currentClientId = selectedClientItemId == null ? null : Number(selectedClientItemId);
    const initialClientId = initialClientState.itemId == null ? null : Number(initialClientState.itemId);
    if (currentClientId !== initialClientId) payload.client_item_id = currentClientId;
    const currentClientTitle = cleanNullableText(selectedClientTitle);
    const initialClientTitle = cleanNullableText(initialClientState.title);
    if (currentClientTitle !== initialClientTitle) payload.client_title = currentClientTitle;
    return payload;
  };

  const buildRepairCreatePayload = (nextValues) => {
    const payload = {
      status_label: nextValues.status_label || "Dossier ouvert",
      manufacturer: nextValues.manufacturer,
      model: nextValues.model,
      serial: nextValues.serial,
      desc_problem: nextValues.desc_problem,
      observations: nextValues.observations,
      desc_done: nextValues.desc_done,
      lines: nextValues.lines,
      duration_minutes: nextValues.duration_minutes,
      parts_cost: nextValues.parts_cost,
      is_personal: Boolean(nextValues.is_personal),
    };
    if (selectedClientItemId != null && Number.isFinite(Number(selectedClientItemId))) payload.client_item_id = Number(selectedClientItemId);
    if (cleanNullableText(selectedClientTitle)) payload.client_title = cleanNullableText(selectedClientTitle);
    return payload;
  };

  if (editForm && saveBtn && saveMessage) {
    const submitRepairChanges = async ({ closeAfterSave = false } = {}) => {
      saveMessage.textContent = "";
      if (quickRepairCostDraft && !isClientEditMode) {
        const savedInline = await persistInlineRepairCostDraft();
        if (savedInline && closeAfterSave) finalizeCloseRepairModal();
        return savedInline;
      }
      if (quickRepairCostDraft) {
        saveMessage.textContent = "Enregistrez ou annulez la ligne ajoutée avant de sauvegarder le bon.";
        renderRepairLines();
        return null;
      }
      const nextLines = normalizeRepairInvoiceLines(draftRepairLines)
        .map((line) => ({
          ...line,
          description: String(line.description ?? "").trim(),
          quantity: Math.max(0, parseRepairLineNumber(line.quantity)),
          unit_price: Math.max(0, parseRepairLineNumber(line.unit_price)),
          currency: normalizeProjectCostCurrency(line.currency),
          exchange_rate_to_cad: getProjectCostLineBaseExchangeRate(line),
          expense_date: normalizeProjectCostExpenseDate(line.expense_date) || null,
          paypal_purchase: normalizeProjectCostPaypalPurchase(line.paypal_purchase),
        }))
        .filter((line, index) => (
          index === 0
          || cleanNullableText(line.description)
          || getProjectCostLineCadAmount(line) > 0
        ))
        .map((line, index) => ({
          ...line,
          description: !isPersonalRepairValue && index === 0 && isRepairLaborLine(line)
            ? "Main d'oeuvre"
            : String(line.description ?? "").trim(),
          amount: getProjectCostLineCadAmount(line),
          source_kind: cleanNullableText(line.source_kind) || (!isPersonalRepairValue && index === 0 ? "repair_labor" : "repair_parts"),
        }));
      draftRepairLines = ensureRepairInvoiceLines(nextLines, {
        defaultLaborRate: getConfiguredFactureLaborRate(),
        hourlyRateSnapshot: repair?.hourly_rate_snapshot,
        fallbackLaborQuantity: isCreatingRepair ? 1 : 0,
        includeLabor: !isPersonalRepairValue,
      });
      const derivedCosts = deriveRepairLegacyValuesFromLines(draftRepairLines);

      const nextValues = {
        status_label: cleanNullableText(draftStatusValue),
        desc_problem: cleanNullableText($("repairProblemInput")?.value),
        observations: cleanNullableText($("repairObservationInput")?.value),
        desc_done: cleanNullableText($("repairWorkDoneInput")?.value),
        manufacturer: cleanNullableText($("repairManufacturerInput")?.value),
        model: cleanNullableText($("repairModelInput")?.value),
        serial: cleanNullableText($("repairSerialInput")?.value),
        lines: draftRepairLines,
        duration_minutes: derivedCosts.duration_minutes,
        parts_cost: derivedCosts.parts_cost,
        is_personal: isPersonalRepairValue,
      };
      const invalidRepairFxLine = draftRepairLines.find((line) => {
        const currency = normalizeProjectCostCurrency(line.currency);
        if (currency === PROJECT_COST_DEFAULT_CURRENCY) return false;
        return !(normalizeProjectCostExchangeRate(line.exchange_rate_to_cad, currency) > 0);
      });
      if (invalidRepairFxLine) {
        saveMessage.textContent = "Definissez un taux de change CAD valide pour chaque depense en devise etrangere.";
        return repair;
      }

      const payload = isCreatingRepair
        ? buildRepairCreatePayload(nextValues)
        : buildRepairPayload(nextValues);
      const facturePaymentDirty = !isCreatingRepair && hasFacturePaymentDirty();
      const photosDirty = serializeRepairPhotos(getRepairPhotosState()) !== serializeRepairPhotos(initialFormValues.photos);
      const thumbnailDirty = serializeRepairPhotoThumbnail(draftPhotoThumbnail) !== serializeRepairPhotoThumbnail(initialFormValues.photo_thumbnail)
        || Boolean(draftPhotoThumbnailPendingBlob);

      if (!Object.keys(payload).length && !facturePaymentDirty && !photosDirty && !thumbnailDirty) {
        saveMessage.textContent = "Aucune modification à sauvegarder.";
        if (closeAfterSave) {
          finalizeCloseRepairModal();
        }
        return repair;
      }

      saveBtn.disabled = true;
      if (cancelChangesBtn) cancelChangesBtn.disabled = true;
      if (completeBtn) completeBtn.disabled = true;
      if (handoverBtn) handoverBtn.disabled = true;
      saveMessage.textContent = "Sauvegarde...";
      try {
        let updatedRepair = repair;
        if (Object.keys(payload).length) {
          updatedRepair = isCreatingRepair
            ? await createRepairRecord(payload)
            : await updateRepairRecord(repair.id, payload);
        }
        if (photosDirty || thumbnailDirty) {
          const mergedPhotos = normalizeRepairPhotos(draftRepairPhotos);
          const uploadedByPendingFilename = new Map();
          if (draftPendingPhotoEntries.length > 0) {
            if (!isDeveloperModeEnabled() && !state.authClient) {
              throw new Error("Connexion Supabase indisponible.");
            }
            for (const pending of draftPendingPhotoEntries) {
              if (!(pending?.file instanceof File)) continue;
              const photoPhase = normalizeRepairPhotoPhase(pending.phase, "before");
              const nextSeq = normalizeRepairPhotos(mergedPhotos)
                .filter((photo) => normalizeRepairPhotoPhase(photo.phase, "before") === photoPhase)
                .length + 1;
              if (isDeveloperModeEnabled()) {
                const uploaded = {
                  filename: buildRepairPhotoObjectPath(
                    updatedRepair,
                    photoPhase,
                    nextSeq,
                    pending.file.name || pending.filename || "photo.jpg",
                    pending.file.type || pending.content_type || "",
                  ),
                  bucket: REPAIRS_PHOTOS_BUCKET,
                  original_name: cleanNullableText(pending.file.name),
                  content_type: cleanNullableText(pending.file.type),
                  size: Number.isFinite(Number(pending.file.size)) ? Number(pending.file.size) : null,
                  created_at: new Date().toISOString(),
                  phase: photoPhase,
                };
                mergedPhotos.push(uploaded);
                uploadedByPendingFilename.set(String(pending.filename), uploaded);
                continue;
              }
              const uploaded = await uploadRepairPhotoFile(
                state.authClient,
                updatedRepair,
                pending.file,
                photoPhase,
                nextSeq,
              );
              mergedPhotos.push(uploaded);
              uploadedByPendingFilename.set(String(pending.filename), uploaded);
            }
          }

          let nextThumbnail = normalizeRepairPhotoThumbnail(draftPhotoThumbnail);
          if (draftPhotoThumbnailPendingBlob && draftPhotoThumbnailPendingSource) {
            if (!isDeveloperModeEnabled() && !state.authClient) {
              throw new Error("Connexion Supabase indisponible.");
            }
            let sourcePhoto = null;
            if (draftPhotoThumbnailPendingSource.kind === "stored") {
              sourcePhoto = normalizeRepairPhotos(draftRepairPhotos)[Number(draftPhotoThumbnailPendingSource.index)] || null;
            }
            if (draftPhotoThumbnailPendingSource.kind === "pending") {
              const pendingFileName = String(draftPhotoThumbnailPendingSource.filename || "");
              sourcePhoto = uploadedByPendingFilename.get(pendingFileName) || null;
              if (!sourcePhoto) {
                const fallbackPending = draftPendingPhotoEntries[Number(draftPhotoThumbnailPendingSource.index)];
                if (fallbackPending) {
                  sourcePhoto = {
                    filename: fallbackPending.filename,
                    bucket: fallbackPending.bucket || REPAIRS_PHOTOS_BUCKET,
                    phase: normalizeRepairPhotoPhase(fallbackPending.phase, "before"),
                  };
                }
              }
            }
            if (isDeveloperModeEnabled()) {
              nextThumbnail = {
                filename: buildRepairThumbnailObjectPath(updatedRepair),
                bucket: REPAIRS_PHOTOS_BUCKET,
                content_type: "image/jpeg",
                source_filename: cleanNullableText(sourcePhoto?.filename),
                created_at: new Date().toISOString(),
                size: Number.isFinite(Number(draftPhotoThumbnailPendingBlob.size)) ? Number(draftPhotoThumbnailPendingBlob.size) : null,
                phase: normalizeRepairPhotoPhase(sourcePhoto?.phase, "before"),
              };
            } else {
              nextThumbnail = await uploadRepairThumbnailBlob(
                state.authClient,
                updatedRepair,
                draftPhotoThumbnailPendingBlob,
                sourcePhoto,
              );
            }
          }

          updatedRepair = await updateRepairRecord(updatedRepair.id, {
            photos: normalizeRepairPhotos(mergedPhotos),
            photo_thumbnail: nextThumbnail,
          });
          dropRepairThumbnailCacheEntry(initialFormValues.photo_thumbnail);
          dropRepairThumbnailCacheEntry(nextThumbnail);
          dropRepairThumbnailCacheEntry(updatedRepair.photo_thumbnail);
        }
        if (facturePaymentDirty && initialFacture) {
          await updateFacturePaymentStatus(initialFacture, Boolean(draftFacturePaidValue));
        }
        renderAccueilCards();
        renderBonsBoard();
        renderFacturesList();
        if (closeAfterSave) {
          finalizeCloseRepairModal();
        } else {
          openRepairModal(updatedRepair);
        }
        return updatedRepair;
      } catch (err) {
        saveMessage.textContent = err?.message || "Impossible de sauvegarder la réparation.";
        return null;
      } finally {
        saveBtn.disabled = false;
        if (cancelChangesBtn) cancelChangesBtn.disabled = false;
        if (completeBtn) completeBtn.disabled = false;
        if (handoverBtn) handoverBtn.disabled = false;
      }
    };

    if (completeBtn) {
      completeBtn.addEventListener("click", async () => {
        if (saveBtn.disabled) return;

        const current = readRepairFormValues();
        const missingClient = !hasCurrentClientLink();
        const missingWorkDone = !cleanNullableText(current.desc_done);
        const missingDuration = !isPersonalRepairValue && !(Number.isFinite(Number(current.duration_minutes)) && Number(current.duration_minutes) > 0);

        if (missingClient || missingWorkDone || missingDuration) {
          if (!isClientEditMode) {
            isClientEditMode = true;
            syncClientEditMode();
          }
          highlightCompletionMissing({
            client: missingClient,
            workDone: missingWorkDone,
            lines: missingDuration,
          });
          saveMessage.textContent = "Complètez les champs requis avant de terminer.";
          refreshRepairDirtyState();
          return;
        }

        clearCompletionHighlights();
        draftStatusValue = "Terminée";
        const statusSelect = $("repairStatusSelect");
        if (statusSelect) {
          statusSelect.value = "Terminée";
          statusSelect.className = `repair-status-select repair-status-select-inline repair-modal-status repair-modal-status-${repairStatusToneKey("Terminée")}`;
        }
        refreshRepairDirtyState();
        const savedRepair = await submitRepairChanges({ closeAfterSave: true });
        if (!savedRepair) return;

        if (isPersonalRepairValue) return;
        const wantsInvoice = window.confirm("Ce bon de réparation est au statut Terminé. Voulez-vous émettre une facture ?");
        if (!wantsInvoice) return;

        const prefillRepairItemId = Number(savedRepair?.podio_item_id ?? repair?.podio_item_id);
        if (Number.isInteger(prefillRepairItemId) && prefillRepairItemId > 0) {
          openCreateFactureModal({ prefillRepairItemId });
        } else {
          openCreateFactureModal();
        }
      });
    }

    if (handoverBtn) {
      handoverBtn.addEventListener("click", async () => {
        if (saveBtn.disabled) return;
        clearCompletionHighlights();
        draftStatusValue = "Remis au client";
        const statusSelect = $("repairStatusSelect");
        if (statusSelect) {
          statusSelect.value = "Remis au client";
          statusSelect.className = `repair-status-select repair-status-select-inline repair-modal-status repair-modal-status-${repairStatusToneKey("Remis au client")}`;
        }
        refreshRepairDirtyState();
        await submitRepairChanges({ closeAfterSave: true });
      });
    }

    modal.__repairSave = submitRepairChanges;
    modal.__repairIsDirty = isRepairFormDirty;
    modal.__repairIsSaving = () => Boolean(saveBtn.disabled);

    editForm.addEventListener("submit", async (ev) => {
      ev.preventDefault();
      await submitRepairChanges();
    });
  }

  refreshRepairDirtyState();
  renderRepairPhotos();
  renderRepairFilesList();
  void hydrateStoredRepairPhotoPreviews();
  modal.__repairCleanup = cleanupRepairPhotoResources;
  modal.hidden = false;
  document.body.style.overflow = "hidden";
  requestAnimationFrame(refreshAutoTextareasHeight);
  setTimeout(refreshAutoTextareasHeight, 60);
}

function openCreateRepairModal() {
  const draft = createEmptyRepairDraft();
  openRepairModal(draft, { isCreating: true, startInEditMode: true });
}

function finalizeCloseRepairModal() {
  const modal = $("repairModal");
  if (!modal) return;
  if (typeof modal.__repairCleanup === "function") {
    try {
      modal.__repairCleanup();
    } catch {
      // ignore
    }
  }
  modal.hidden = true;
  modal.classList.remove("is-client-editing");
  modal.__repairSave = null;
  modal.__repairIsDirty = null;
  modal.__repairIsSaving = null;
  modal.__repairCleanup = null;
  const body = $("repairModalBody");
  if (body) body.innerHTML = "";
  const clientEditBtn = $("repairModalClientEditBtn");
  if (clientEditBtn) {
    clientEditBtn.classList.remove("is-active");
    clientEditBtn.setAttribute("aria-pressed", "false");
  }
  if (
    (!$("clientModal") || $("clientModal").hidden)
    && (!$("factureModal") || $("factureModal").hidden)
    && (!$("projectModal") || $("projectModal").hidden)
    && (!$("expenseModal") || $("expenseModal").hidden)
  ) {
    document.body.style.overflow = "";
  }
}

function showRepairUnsavedChangesDialog() {
  if (repairUnsavedDialogPromise) return repairUnsavedDialogPromise;

  repairUnsavedDialogPromise = new Promise((resolve) => {
    const previousActive = document.activeElement instanceof HTMLElement ? document.activeElement : null;
    const root = document.createElement("div");
    root.id = "repairUnsavedConfirm";
    root.className = "repair-confirm-overlay";
    root.innerHTML = `
      <div class="repair-confirm-panel" role="dialog" aria-modal="true" aria-labelledby="repairConfirmTitle">
        <h4 id="repairConfirmTitle" class="repair-confirm-title">Modifications non sauvegardées</h4>
        <p class="repair-confirm-text">
          Des modifications ont été apportées à ce bon de réparation.
          Que voulez-vous faire ?
        </p>
        <div class="repair-confirm-actions">
          <button type="button" class="repair-confirm-btn repair-confirm-btn-primary" data-choice="save">Enregistrer</button>
          <button type="button" class="repair-confirm-btn repair-confirm-btn-danger" data-choice="discard">Abandonner les modifications</button>
          <button type="button" class="repair-confirm-btn" data-choice="cancel">Annuler</button>
        </div>
      </div>
    `;

    const finalize = (choice) => {
      document.removeEventListener("keydown", onKeyDown, true);
      root.remove();
      repairUnsavedDialogPromise = null;
      if (previousActive && document.contains(previousActive)) previousActive.focus();
      resolve(choice);
    };

    const onKeyDown = (ev) => {
      if (ev.key !== "Escape") return;
      ev.preventDefault();
      ev.stopPropagation();
      finalize("cancel");
    };

    root.addEventListener("click", (ev) => {
      const btn = ev.target?.closest?.("[data-choice]");
      if (btn) {
        finalize(btn.getAttribute("data-choice"));
        return;
      }
      if (ev.target === root) {
        finalize("cancel");
      }
    });

    document.addEventListener("keydown", onKeyDown, true);
    document.body.appendChild(root);
    root.querySelector("[data-choice='cancel']")?.focus();
  });

  return repairUnsavedDialogPromise;
}

async function closeRepairModal() {
  const modal = $("repairModal");
  if (!modal || modal.hidden) return true;

  const isSaving = typeof modal.__repairIsSaving === "function" ? Boolean(modal.__repairIsSaving()) : false;
  if (isSaving) return false;

  const isDirty = typeof modal.__repairIsDirty === "function" ? Boolean(modal.__repairIsDirty()) : false;
  if (isDirty) {
    const choice = await showRepairUnsavedChangesDialog();
    if (choice === "cancel") return false;
    if (choice === "discard") {
      finalizeCloseRepairModal();
      return true;
    }

    const saveFn = modal.__repairSave;
    if (choice === "save" && typeof saveFn === "function") {
      const saved = await saveFn({ closeAfterSave: true });
      return Boolean(saved);
    }
    return false;
  }

  finalizeCloseRepairModal();
  return true;
}

function finalizeCloseFactureModal() {
  const modal = $("factureModal");
  if (!modal) return;
  modal.hidden = true;
  modal.__factureIsDirty = null;
  modal.__factureIsSaving = null;
  modal.__factureFinalizeClose = null;
  const body = $("factureModalBody");
  if (body) body.innerHTML = "";
  factureModalState.open = false;
  if (
    (!$("repairModal") || $("repairModal").hidden)
    && (!$("clientModal") || $("clientModal").hidden)
    && (!$("projectModal") || $("projectModal").hidden)
    && (!$("expenseModal") || $("expenseModal").hidden)
  ) {
    document.body.style.overflow = "";
  }
}

async function closeFactureModal() {
  const modal = $("factureModal");
  if (!modal || modal.hidden) return true;
  const isSaving = typeof modal.__factureIsSaving === "function" ? Boolean(modal.__factureIsSaving()) : false;
  if (isSaving) return false;

  const isDirty = typeof modal.__factureIsDirty === "function" ? Boolean(modal.__factureIsDirty()) : false;
  if (isDirty) {
    const ok = window.confirm("Des modifications ont été apportées à la facture. Fermer sans enregistrer ?");
    if (!ok) return false;
  }

  finalizeCloseFactureModal();
  return true;
}

function showProjectProceedToFactureDialog() {
  if (projectProceedToFactureDialogPromise) return projectProceedToFactureDialogPromise;

  projectProceedToFactureDialogPromise = new Promise((resolve) => {
    const previousActive = document.activeElement instanceof HTMLElement ? document.activeElement : null;
    const root = document.createElement("div");
    root.id = "projectProceedToFactureDialog";
    root.className = "repair-confirm-overlay";
    root.innerHTML = `
      <div class="repair-confirm-panel" role="dialog" aria-modal="true" aria-labelledby="projectProceedToFactureTitle">
        <h4 id="projectProceedToFactureTitle" class="repair-confirm-title">Projet modifié</h4>
        <p class="repair-confirm-text">
          Des modifications ont été apportées au projet. Enregistrer et poursuivre vers la facture ?
        </p>
        <div class="repair-confirm-actions">
          <button type="button" class="repair-confirm-btn repair-confirm-btn-primary" data-choice="save_continue">Enregistrer et poursuivre</button>
          <button type="button" class="repair-confirm-btn" data-choice="cancel">Annuler</button>
        </div>
      </div>
    `;

    const finalize = (choice) => {
      document.removeEventListener("keydown", onKeyDown, true);
      root.remove();
      projectProceedToFactureDialogPromise = null;
      if (previousActive && document.contains(previousActive)) previousActive.focus();
      resolve(choice);
    };

    const onKeyDown = (ev) => {
      if (ev.key !== "Escape") return;
      ev.preventDefault();
      ev.stopPropagation();
      finalize("cancel");
    };

    root.addEventListener("click", (ev) => {
      const btn = ev.target?.closest?.("[data-choice]");
      if (btn) {
        finalize(btn.getAttribute("data-choice"));
        return;
      }
      if (ev.target === root) finalize("cancel");
    });

    document.addEventListener("keydown", onKeyDown, true);
    document.body.appendChild(root);
    root.querySelector("[data-choice='cancel']")?.focus();
  });

  return projectProceedToFactureDialogPromise;
}

function finalizeCloseProjectModal() {
  const modal = $("projectModal");
  if (!modal) return;
  if (typeof modal.__projectCleanup === "function") {
    try {
      modal.__projectCleanup();
    } catch {
      // ignore cleanup failures
    }
  }
  modal.hidden = true;
  modal.__projectCleanup = null;
  modal.__projectIsDirty = null;
  modal.__projectIsSaving = null;
  modal.__projectSave = null;
  const body = $("projectModalBody");
  if (body) body.innerHTML = "";
  projectModalState.open = false;
  if (
    (!$("repairModal") || $("repairModal").hidden)
    && (!$("clientModal") || $("clientModal").hidden)
    && (!$("factureModal") || $("factureModal").hidden)
    && (!$("expenseModal") || $("expenseModal").hidden)
  ) {
    document.body.style.overflow = "";
  }
}

async function closeProjectModal() {
  const modal = $("projectModal");
  if (!modal || modal.hidden) return true;
  const isSaving = typeof modal.__projectIsSaving === "function" ? Boolean(modal.__projectIsSaving()) : false;
  if (isSaving) return false;
  const isDirty = typeof modal.__projectIsDirty === "function" ? Boolean(modal.__projectIsDirty()) : false;
  if (isDirty) {
    const ok = window.confirm("Des modifications ont été apportées au projet. Fermer sans enregistrer ?");
    if (!ok) return false;
  }
  finalizeCloseProjectModal();
  return true;
}

function finalizeCloseExpenseModal() {
  const modal = $("expenseModal");
  if (!modal) return;
  if (typeof modal.__expenseCleanup === "function") {
    try {
      modal.__expenseCleanup();
    } catch {
      // ignore cleanup failures
    }
  }
  modal.hidden = true;
  modal.__expenseCleanup = null;
  modal.__expenseIsDirty = null;
  modal.__expenseIsSaving = null;
  modal.__expenseSave = null;
  const body = $("expenseModalBody");
  if (body) body.innerHTML = "";
  expenseModalState.open = false;
  if (
    (!$("repairModal") || $("repairModal").hidden)
    && (!$("clientModal") || $("clientModal").hidden)
    && (!$("factureModal") || $("factureModal").hidden)
    && (!$("projectModal") || $("projectModal").hidden)
  ) {
    document.body.style.overflow = "";
  }
}

async function closeExpenseModal() {
  const modal = $("expenseModal");
  if (!modal || modal.hidden) return true;
  const isSaving = typeof modal.__expenseIsSaving === "function" ? Boolean(modal.__expenseIsSaving()) : false;
  if (isSaving) return false;
  const isDirty = typeof modal.__expenseIsDirty === "function" ? Boolean(modal.__expenseIsDirty()) : false;
  if (isDirty) {
    const ok = window.confirm("Des modifications ont été apportées à la dépense. Fermer sans enregistrer ?");
    if (!ok) return false;
  }
  finalizeCloseExpenseModal();
  return true;
}

function openExpenseModal(options = {}) {
  const modal = $("expenseModal");
  const body = $("expenseModalBody");
  const title = $("expenseModalTitle");
  if (!modal || !body || !title) return;

  if (typeof modal.__expenseCleanup === "function") {
    try {
      modal.__expenseCleanup();
    } catch {
      // ignore
    }
  }
  modal.__expenseCleanup = null;

  let row = options?.row && typeof options.row === "object" ? normalizeExpenseRow(options.row) : null;
  let hasPersistedExpense = Number.isInteger(Number(row?.id)) && Number(row?.id) > 0;
  let saving = false;
  let draft = {
    title: cleanNullableText(row?.title) || "",
    description: cleanNullableText(row?.description) || "",
    purchase_date: /^\d{4}-\d{2}-\d{2}$/.test(String(row?.purchase_date || "")) ? String(row.purchase_date) : formatDateYmd(new Date()),
    amount: Number.isFinite(Number(row?.amount)) ? Math.max(0, Number(row.amount)) : 0,
  };
  let invoiceMeta = {
    invoice_filename: cleanNullableText(row?.invoice_filename),
    invoice_bucket: cleanNullableText(row?.invoice_bucket) || EXPENSES_STORAGE_BUCKET,
    invoice_path: cleanNullableText(row?.invoice_path),
    invoice_content_type: cleanNullableText(row?.invoice_content_type),
    preview_url: cleanNullableText(row?.preview_url),
  };
  let pendingInvoiceFile = null;
  let pendingInvoicePreviewUrl = null;
  const objectUrlsToRevoke = new Set();

  const readSnapshot = () => JSON.stringify({
    title: draft.title,
    description: draft.description,
    purchase_date: draft.purchase_date,
    amount: Math.round(Number(draft.amount || 0) * 100) / 100,
    invoice_filename: invoiceMeta.invoice_filename || "",
    invoice_path: invoiceMeta.invoice_path || "",
    pending_invoice_name: cleanNullableText(pendingInvoiceFile?.name) || "",
  });
  let initialSnapshot = readSnapshot();

  body.innerHTML = `
    <form id="expenseForm" class="facture-form expense-form">
      <section class="facture-section">
        <h4 class="facture-section-title">Titre</h4>
        <input id="expenseTitleInput" class="facture-input" type="text" maxlength="160" placeholder="Ex.: Composants DigiKey">
      </section>

      <section class="facture-section">
        <h4 class="facture-section-title">Description</h4>
        <textarea id="expenseDescriptionInput" class="facture-textarea" placeholder="Détails de la dépense"></textarea>
      </section>

      <div class="expense-form-grid">
        <section class="facture-section">
          <h4 class="facture-section-title">Date d'achat</h4>
          <input id="expensePurchaseDateInput" class="facture-input" type="date">
        </section>

        <section class="facture-section">
          <h4 class="facture-section-title">Prix</h4>
          <input id="expenseAmountInput" class="facture-input" type="number" min="0" step="0.01" inputmode="decimal" placeholder="0,00">
        </section>
      </div>

      <section class="facture-section">
        <input id="expenseInvoiceInput" type="file" accept=".pdf,image/*" hidden>
        <div class="repair-photos-head expense-file-head">
          <h4 class="facture-section-title">Facture</h4>
          <button id="expenseInvoiceAddBtn" type="button" class="repair-client-add-btn repair-photos-add-btn">Téléverser</button>
        </div>
        <div id="expenseInvoiceCard" class="expense-invoice-card"></div>
        <p id="expenseInvoiceMessage" class="facture-message"></p>
      </section>

      <div class="facture-actions">
        <p id="expenseMessage" class="facture-message"></p>
        <div class="facture-actions-right">
          <button id="expenseCancelBtn" type="button" class="repair-btn">Annuler</button>
          <button id="expenseSaveBtn" type="submit" class="repair-btn repair-btn-primary">Enregistrer</button>
        </div>
      </div>
    </form>
  `;

  const form = $("expenseForm");
  const titleInput = $("expenseTitleInput");
  const descriptionInput = $("expenseDescriptionInput");
  const purchaseDateInput = $("expensePurchaseDateInput");
  const amountInput = $("expenseAmountInput");
  const invoiceInput = $("expenseInvoiceInput");
  const invoiceAddBtn = $("expenseInvoiceAddBtn");
  const invoiceCard = $("expenseInvoiceCard");
  const invoiceMessage = $("expenseInvoiceMessage");
  const cancelBtn = $("expenseCancelBtn");
  const saveBtn = $("expenseSaveBtn");
  const messageEl = $("expenseMessage");

  if (titleInput) titleInput.value = draft.title;
  if (descriptionInput) descriptionInput.value = draft.description;
  if (purchaseDateInput) purchaseDateInput.value = draft.purchase_date;
  if (amountInput) amountInput.value = draft.amount > 0 ? String(draft.amount) : "";

  const syncDraftFromInputs = () => {
    draft = {
      title: cleanNullableText(titleInput?.value) || "",
      description: cleanNullableText(descriptionInput?.value) || "",
      purchase_date: /^\d{4}-\d{2}-\d{2}$/.test(String(purchaseDateInput?.value || "")) ? String(purchaseDateInput.value) : formatDateYmd(new Date()),
      amount: Math.max(0, normalizeNonNegativeNumber(amountInput?.value, 0)),
    };
  };

  const renderInvoiceCard = () => {
    if (!invoiceCard) return;
    const pendingName = cleanNullableText(pendingInvoiceFile?.name);
    const fileName = pendingName || invoiceMeta.invoice_filename;
    if (!fileName) {
      invoiceCard.innerHTML = `<p class="expense-invoice-empty">Aucune facture téléversée.</p>`;
      return;
    }
    const hasPreview = Boolean(pendingInvoicePreviewUrl || invoiceMeta.invoice_path || invoiceMeta.preview_url);
    invoiceCard.innerHTML = `
      <div class="expense-invoice-file">
        <div class="expense-invoice-file-main">
          <p class="expense-invoice-file-name">${escapeHtml(fileName)}</p>
          <p class="expense-invoice-file-meta">${escapeHtml(pendingName ? "Fichier en attente d'enregistrement" : "Fichier enregistré")}</p>
        </div>
        <div class="expense-invoice-file-actions">
          <button id="expenseInvoicePreviewBtn" type="button" class="project-file-action-btn is-preview" ${hasPreview ? "" : "disabled"} title="Voir la facture" aria-label="Voir la facture">${FACTURE_PREVIEW_EYE_SVG}</button>
        </div>
      </div>
    `;
    const previewBtn = $("expenseInvoicePreviewBtn");
    if (previewBtn) {
      previewBtn.addEventListener("click", async () => {
        try {
          if (pendingInvoicePreviewUrl) {
            const lowerName = String(fileName || "").toLowerCase();
            if (lowerName.endsWith(".pdf")) {
              const openedInModal = await openFacturePdfModal(pendingInvoicePreviewUrl, `Aperçu PDF - ${fileName}`);
              if (!openedInModal) window.open(pendingInvoicePreviewUrl, "_blank", "noopener,noreferrer");
            } else {
              openImagePreviewGallery([{ title: fileName, getUrl: () => Promise.resolve(pendingInvoicePreviewUrl) }], 0, fileName);
            }
            return;
          }
          const currentExpense = row ? { ...row, ...invoiceMeta } : { ...invoiceMeta, id: 0, title: draft.title };
          await openExpenseInvoice(currentExpense);
        } catch (err) {
          if (invoiceMessage) invoiceMessage.textContent = err?.message || "Impossible d'ouvrir la facture.";
        }
      });
    }
  };

  const upsertExpenseInState = (nextRow) => {
    const normalized = normalizeExpenseRow(nextRow);
    if (!normalized) return null;
    const hasExpense = (state.expenses || []).some((expense) => String(expense?.id) === String(normalized.id));
    const nextExpenses = (state.expenses || []).map((expense) => (String(expense?.id) === String(normalized.id) ? normalized : expense));
    state.expenses = hasExpense ? nextExpenses : [normalized, ...nextExpenses];
    state.expensesLoadError = null;
    renderExpensesList();
    return normalized;
  };

  const saveExpenseChanges = async () => {
    syncDraftFromInputs();
    const expenseTitle = cleanNullableText(draft.title);
    if (!expenseTitle) {
      if (messageEl) messageEl.textContent = "Le titre de la dépense est requis.";
      return false;
    }
    if (!/^\d{4}-\d{2}-\d{2}$/.test(String(draft.purchase_date || ""))) {
      if (messageEl) messageEl.textContent = "La date d'achat est requise.";
      return false;
    }
    if (!(Number(draft.amount) >= 0)) {
      if (messageEl) messageEl.textContent = "Le prix doit être valide.";
      return false;
    }

    const payload = {
      title: expenseTitle,
      description: cleanNullableText(draft.description),
      purchase_date: draft.purchase_date,
      amount: Math.round(Number(draft.amount || 0) * 100) / 100,
      updated_at: new Date().toISOString(),
    };

    try {
      saving = true;
      if (saveBtn) saveBtn.disabled = true;
      if (cancelBtn) cancelBtn.disabled = true;
      if (messageEl) messageEl.textContent = "Enregistrement...";
      if (invoiceMessage) invoiceMessage.textContent = "";

      const developerMode = isDeveloperModeEnabled();
      const client = state.authClient;
      if (!developerMode && !client) throw new Error("Connexion Supabase indisponible.");

      let saved = null;
      if (developerMode) {
        const now = new Date().toISOString();
        saved = {
          ...(row || {}),
          id: Number.isInteger(Number(row?.id)) ? Number(row.id) : nextLocalEntityId(),
          ...payload,
          created_at: cleanNullableText(row?.created_at) || now,
          updated_at: now,
        };
      } else if (hasPersistedExpense && row) {
        const { data, error } = await client
          .from("expenses")
          .update(payload)
          .eq("id", Number(row.id))
          .select("*")
          .single();
        if (error) throw new Error(error.message || "Impossible d'enregistrer la dépense.");
        saved = data;
      } else {
        const { data, error } = await client
          .from("expenses")
          .insert(payload)
          .select("*")
          .single();
        if (error) throw new Error(error.message || "Impossible de créer la dépense.");
        saved = data;
      }

      if (!saved) throw new Error("Dépense introuvable après enregistrement.");

      if (pendingInvoiceFile) {
        if (developerMode) {
          invoiceMeta = {
            invoice_filename: cleanNullableText(pendingInvoiceFile.name),
            invoice_bucket: EXPENSES_STORAGE_BUCKET,
            invoice_path: `local/${sanitizeFilenameToken(pendingInvoiceFile.name, "facture")}`,
            invoice_content_type: cleanNullableText(pendingInvoiceFile.type),
            preview_url: pendingInvoicePreviewUrl,
          };
          if (pendingInvoicePreviewUrl) objectUrlsToRevoke.delete(pendingInvoicePreviewUrl);
          saved = { ...saved, ...invoiceMeta };
        } else {
          const stamp = Date.now();
          const safeName = sanitizeFilenameToken(pendingInvoiceFile.name || `facture-${stamp}`, `facture-${stamp}`);
          const storagePath = `${String(saved.id)}/${stamp}-${safeName}`;
          const { error: uploadError } = await client
            .storage
            .from(EXPENSES_STORAGE_BUCKET)
            .upload(storagePath, pendingInvoiceFile, {
              contentType: pendingInvoiceFile.type || "application/octet-stream",
              upsert: true,
            });
          if (uploadError) throw new Error(uploadError.message || "Impossible de téléverser la facture.");
          const invoicePayload = {
            invoice_filename: cleanNullableText(pendingInvoiceFile.name),
            invoice_bucket: EXPENSES_STORAGE_BUCKET,
            invoice_path: storagePath,
            invoice_content_type: cleanNullableText(pendingInvoiceFile.type),
            updated_at: new Date().toISOString(),
          };
          const { data: savedWithInvoice, error: invoiceError } = await client
            .from("expenses")
            .update(invoicePayload)
            .eq("id", Number(saved.id))
            .select("*")
            .single();
          if (invoiceError) throw new Error(invoiceError.message || "Impossible d'associer la facture à la dépense.");
          saved = savedWithInvoice;
          invoiceMeta = {
            invoice_filename: cleanNullableText(saved?.invoice_filename),
            invoice_bucket: cleanNullableText(saved?.invoice_bucket) || EXPENSES_STORAGE_BUCKET,
            invoice_path: cleanNullableText(saved?.invoice_path),
            invoice_content_type: cleanNullableText(saved?.invoice_content_type),
            preview_url: null,
          };
        }
      } else {
        invoiceMeta = {
          invoice_filename: cleanNullableText(saved?.invoice_filename),
          invoice_bucket: cleanNullableText(saved?.invoice_bucket) || EXPENSES_STORAGE_BUCKET,
          invoice_path: cleanNullableText(saved?.invoice_path),
          invoice_content_type: cleanNullableText(saved?.invoice_content_type),
          preview_url: cleanNullableText(saved?.preview_url),
        };
      }

      const normalized = upsertExpenseInState(saved);
      row = normalized;
      hasPersistedExpense = true;
      pendingInvoiceFile = null;
      pendingInvoicePreviewUrl = null;
      if (invoiceInput) invoiceInput.value = "";
      renderInvoiceCard();
      if (messageEl) messageEl.textContent = "Dépense enregistrée.";
      initialSnapshot = readSnapshot();
      title.textContent = `Dépense | ${expenseTitle}`;
      return true;
    } catch (err) {
      if (messageEl) messageEl.textContent = err?.message || "Impossible d'enregistrer la dépense.";
      return false;
    } finally {
      saving = false;
      if (saveBtn) saveBtn.disabled = false;
      if (cancelBtn) cancelBtn.disabled = false;
    }
  };

  if (form) {
    form.addEventListener("input", () => {
      syncDraftFromInputs();
      if (messageEl) messageEl.textContent = "";
    });
    form.addEventListener("submit", async (ev) => {
      ev.preventDefault();
      await saveExpenseChanges();
    });
  }

  if (invoiceAddBtn && invoiceInput) {
    invoiceAddBtn.addEventListener("click", () => invoiceInput.click());
    invoiceInput.addEventListener("change", () => {
      const file = invoiceInput.files?.[0] || null;
      pendingInvoiceFile = file;
      if (pendingInvoicePreviewUrl && objectUrlsToRevoke.has(pendingInvoicePreviewUrl)) {
        URL.revokeObjectURL(pendingInvoicePreviewUrl);
        objectUrlsToRevoke.delete(pendingInvoicePreviewUrl);
      }
      pendingInvoicePreviewUrl = file ? URL.createObjectURL(file) : null;
      if (pendingInvoicePreviewUrl) objectUrlsToRevoke.add(pendingInvoicePreviewUrl);
      if (invoiceMessage) invoiceMessage.textContent = "";
      renderInvoiceCard();
    });
  }

  if (cancelBtn) {
    cancelBtn.addEventListener("click", async () => {
      await closeExpenseModal();
    });
  }

  renderInvoiceCard();
  title.textContent = hasPersistedExpense && row ? `Dépense | ${row.title}` : "Nouvelle dépense";
  modal.hidden = false;
  expenseModalState.open = true;
  document.body.style.overflow = "hidden";

  modal.__expenseIsDirty = () => readSnapshot() !== initialSnapshot;
  modal.__expenseIsSaving = () => saving;
  modal.__expenseSave = saveExpenseChanges;
  modal.__expenseCleanup = () => {
    objectUrlsToRevoke.forEach((url) => {
      try {
        URL.revokeObjectURL(url);
      } catch {
        // ignore
      }
    });
    objectUrlsToRevoke.clear();
  };
}

function openProjectModal(options = {}) {
  const modal = $("projectModal");
  const body = $("projectModalBody");
  const title = $("projectModalTitle");
  if (!modal || !body || !title) return;

  if (typeof modal.__projectCleanup === "function") {
    try {
      modal.__projectCleanup();
    } catch {
      // ignore
    }
  }
  modal.__projectCleanup = null;

  let row = options?.row && typeof options.row === "object" ? options.row : null;
  let hasPersistedProject = Number.isInteger(Number(row?.id)) && Number(row?.id) > 0;
  let isEditing = !hasPersistedProject;
  let projectCode = (cleanNullableText(row?.numero) || nextProjectNumeroFromState()).toUpperCase();
  let projectTitle = cleanNullableText(row?.title) || "";
  let projectDescription = cleanNullableText(row?.description) || "";
  let projectNotes = cleanNullableText(row?.notes) || "";
  let isPersonalProjectValue = isPersonalProject(row);
  const initialIsPersonalProject = isPersonalProjectValue;
  const initialBillableHours = Number.isFinite(Number(row?.billable_hours)) ? Math.max(0, Number(row.billable_hours)) : 0;
  let stages = normalizeProjectStages(row?.stages);
  if (!stages.length) stages = [{ id: "stage-1", label: "Ouverture", description: "", done: false }];
  let costLines = normalizeProjectCostLines(row?.cost_lines);
  let quickCostDraft = null;
  let existingFiles = normalizeProjectFiles(row?.files);
  let existingPhotos = normalizeProjectPhotos(row?.photos);
  let pendingPhotos = [];
  let pendingSeq = 1;
  let stageSeq = stages.length + 1;
  let saving = false;
  let clientPickerOpen = false;
  const photoPreviewUrlByKey = {};
  const pendingPhotoPreviewUrls = new Set();

  let sortedClients = [];
  const clientById = {};
  const rebuildClients = () => {
    sortedClients = [...(state.clients || [])]
      .filter((client) => Number.isInteger(getClientReferenceId(client)))
      .sort((a, b) => normalizeText(formatClientPrimaryName(a)).localeCompare(normalizeText(formatClientPrimaryName(b)), "fr-CA"));
    Object.keys(clientById).forEach((k) => delete clientById[k]);
    for (const client of sortedClients) {
      clientById[String(getClientReferenceId(client))] = client;
    }
  };
  rebuildClients();

  let selectedClientId = "";
  let selectedClientTitle = cleanNullableText(row?.client_title) || "";
  let selectedClientCompany = "";
  const initialClientItemId = Number(row?.client_item_id);
  if (Number.isInteger(initialClientItemId) && initialClientItemId > 0) {
    selectedClientId = String(initialClientItemId);
    const selectedClient = clientById[selectedClientId] || null;
    if (selectedClient) {
      selectedClientTitle = formatClientPrimaryName(selectedClient);
      selectedClientCompany = cleanNullableText(selectedClient?.company) || "";
    }
  }
  if (isPersonalProjectValue) {
    const personalOwner = getPersonalOwnerSelection();
    selectedClientId = personalOwner.itemId != null ? String(personalOwner.itemId) : "";
    selectedClientTitle = personalOwner.title;
    selectedClientCompany = personalOwner.company || "";
  }

  const parseNumberOrZero = (value) => {
    const raw = String(value ?? "").trim().replace(",", ".");
    const n = Number(raw);
    return Number.isFinite(n) ? n : 0;
  };
  const round2 = (value) => {
    const n = Number(value);
    if (!Number.isFinite(n)) return 0;
    return Math.round(n * 100) / 100;
  };
  const photoKey = (photo) => `${photo?.bucket || REPAIRS_PHOTOS_BUCKET}/${photo?.filename || ""}`;
  const laborRateOptions = (() => {
    const cfg = state.appConfig || DEFAULT_APP_CONFIG;
    return [
      {
        key: "labor_rate_diagnostic",
        label: "Taux reduit",
        rate: normalizeNonNegativeNumber(cfg.labor_rate_diagnostic, DEFAULT_APP_CONFIG.labor_rate_diagnostic),
      },
      {
        key: "labor_rate_standard",
        label: "Taux standard",
        rate: normalizeNonNegativeNumber(cfg.labor_rate_standard, DEFAULT_APP_CONFIG.labor_rate_standard),
      },
      {
        key: "labor_rate_secondary",
        label: "Taux accelere",
        rate: normalizeNonNegativeNumber(cfg.labor_rate_secondary, DEFAULT_APP_CONFIG.labor_rate_secondary),
      },
      {
        key: "labor_rate_estimation",
        label: "Taux estimation",
        rate: normalizeNonNegativeNumber(cfg.labor_rate_estimation, DEFAULT_APP_CONFIG.labor_rate_estimation),
      },
    ];
  })();
  const laborRateOptionByKey = Object.fromEntries(laborRateOptions.map((opt) => [opt.key, opt]));
  const defaultLaborRateKey = "labor_rate_standard";
  const customLaborRateKey = "custom";
  const getLaborRateValueByKey = (key) => {
    const option = laborRateOptionByKey[String(key ?? "").trim()];
    if (!option) return normalizeNonNegativeNumber((state.appConfig || DEFAULT_APP_CONFIG).labor_rate_standard, DEFAULT_APP_CONFIG.labor_rate_standard);
    return round2(option.rate);
  };
  const findLaborRateKeyByValue = (value) => {
    const target = round2(parseNumberOrZero(value));
    for (const option of laborRateOptions) {
      if (Math.abs(target - round2(option.rate)) < 0.005) return option.key;
    }
    return customLaborRateKey;
  };
  const isLaborDescription = (value) => {
    const normalized = normalizeText(value)
      .replace(/[’'`]/g, "")
      .replace(/\s+/g, " ")
      .trim();
    return normalized.includes("main doeuvre") || normalized.includes("main d oeuvre");
  };
  const isProjectLaborLine = (line) => {
    if (!line || typeof line !== "object") return false;
    return String(line.source_kind ?? "").trim() === "project_labor";
  };
  const isLegacyLaborLine = (line) => {
    if (isProjectLaborLine(line)) return true;
    return isLaborDescription(line?.description);
  };
  const buildLaborRateOptionsHtml = (selectedKey) => {
    const selected = String(selectedKey || "").trim();
    const isKnown = Boolean(laborRateOptionByKey[selected]) || selected === customLaborRateKey;
    const currentKey = isKnown ? selected : defaultLaborRateKey;
    const opts = laborRateOptions.map((option) => {
      const selected = option.key === currentKey ? "selected" : "";
      const label = `${option.label} (${formatMoneyCompact(option.rate)}/h)`;
      return `<option value="${escapeHtml(option.key)}" ${selected}>${escapeHtml(label)}</option>`;
    });
    const customSelected = currentKey === customLaborRateKey ? "selected" : "";
    opts.push(`<option value="${escapeHtml(customLaborRateKey)}" ${customSelected}>${escapeHtml("Taux personnalise")}</option>`);
    return opts.join("");
  };
  const buildDefaultLaborLine = (qty = 1) => {
    const quantity = Math.max(0, round2(parseNumberOrZero(qty)));
    const unit = getLaborRateValueByKey(defaultLaborRateKey);
    return {
      description: "Main d'oeuvre",
      quantity: quantity > 0 ? quantity : 1,
      unit_price: unit,
      amount: round2((quantity > 0 ? quantity : 1) * unit),
      currency: PROJECT_COST_DEFAULT_CURRENCY,
      exchange_rate_to_cad: 1,
      expense_date: "",
      labor_rate_key: defaultLaborRateKey,
      paypal_purchase: false,
      source_kind: "project_labor",
    };
  };
  const buildEmptyQuickCostDraft = () => ({
    description: "",
    quantity: 1,
    unit_price: 0,
    currency: PROJECT_COST_DEFAULT_CURRENCY,
    exchange_rate_to_cad: 1,
    expense_date: formatDateYmd(new Date()),
    paypal_purchase: false,
  });
  const buildProjectCurrencyOptionsHtml = (selectedCurrency) => {
    const current = normalizeProjectCostCurrency(selectedCurrency);
    return PROJECT_COST_CURRENCIES.map((currencyCode) => {
      const selected = currencyCode === current ? "selected" : "";
      return `<option value="${escapeHtml(currencyCode)}" ${selected}>${escapeHtml(currencyCode)}</option>`;
    }).join("");
  };
  const syncProjectCostLineDerivedValues = (line) => {
    if (!line || typeof line !== "object") return line;
    line.currency = normalizeProjectCostCurrency(line.currency);
    if (isProjectLaborLine(line)) {
      line.currency = PROJECT_COST_DEFAULT_CURRENCY;
      line.exchange_rate_to_cad = 1;
      line.expense_date = "";
      line.paypal_purchase = false;
    } else if (line.currency === PROJECT_COST_DEFAULT_CURRENCY) {
      line.exchange_rate_to_cad = 1;
      line.paypal_purchase = false;
    } else {
      line.exchange_rate_to_cad = normalizeProjectCostExchangeRate(line.exchange_rate_to_cad, line.currency);
      line.paypal_purchase = normalizeProjectCostPaypalPurchase(line.paypal_purchase);
    }
    line.expense_date = normalizeProjectCostExpenseDate(line.expense_date);
    line.amount = getProjectCostLineCadAmount(line);
    return line;
  };
  const fetchAndApplyProjectCostExchangeRate = async (line, { silent = false } = {}) => {
    if (!line || typeof line !== "object") return false;
    const currency = normalizeProjectCostCurrency(line.currency);
    if (currency === PROJECT_COST_DEFAULT_CURRENCY) {
      line.exchange_rate_to_cad = 1;
      line.amount = getProjectCostLineCadAmount(line);
      return true;
    }
    clearProjectCostLineExchangeRate(line);
    const expenseDate = normalizeProjectCostExpenseDate(line.expense_date) || formatDateYmd(new Date());
    line.expense_date = expenseDate;
    const requestToken = `fx-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`;
    line.__fx_request_token = requestToken;
    try {
      const nextRate = await fetchHistoricalCadExchangeRate(currency, expenseDate);
      if (line.__fx_request_token !== requestToken) return false;
      line.exchange_rate_to_cad = nextRate;
      line.amount = getProjectCostLineCadAmount(line);
      return true;
    } catch (err) {
      if (!silent && messageEl) {
        messageEl.textContent = err?.message || "Impossible de recuperer le taux de change.";
      }
      return false;
    } finally {
      if (line.__fx_request_token === requestToken) delete line.__fx_request_token;
    }
  };
  const normalizeProjectCostsForSave = (lines) => {
    const normalizedCosts = lines
      .map((line) => ({
        description: cleanNullableText(line.description) || "",
        quantity: round2(Math.max(0, parseNumberOrZero(line.quantity))),
        unit_price: round2(Math.max(0, parseNumberOrZero(line.unit_price))),
        currency: normalizeProjectCostCurrency(line.currency),
        exchange_rate_to_cad: getProjectCostLineBaseExchangeRate(line),
        expense_date: normalizeProjectCostExpenseDate(line.expense_date) || null,
        labor_rate_key: cleanNullableText(line.labor_rate_key),
        paypal_purchase: normalizeProjectCostPaypalPurchase(line.paypal_purchase),
        source_kind: cleanNullableText(line.source_kind),
      }))
      .filter((line) => line.description || line.quantity > 0 || line.unit_price > 0)
      .map((line) => {
        if (String(line.source_kind || "").trim() === "project_labor") {
          line.currency = PROJECT_COST_DEFAULT_CURRENCY;
          line.exchange_rate_to_cad = 1;
          line.expense_date = null;
          line.paypal_purchase = false;
        }
        return {
          ...line,
          amount: getProjectCostLineCadAmount(line),
        };
      });
    const laborLine = normalizedCosts.find((line) => String(line.source_kind || "").trim() === "project_labor")
      || normalizedCosts.find((line) => isLaborDescription(line.description));
    const billableHours = laborLine
      ? round2(Math.max(0, parseNumberOrZero(laborLine.quantity)))
      : 0;
    return {
      normalizedCosts,
      billableHours,
    };
  };
  const applySavedProjectState = (saved, options = {}) => {
    const {
      successMessage = "Projet enregistré.",
      nextEditing = false,
      previousRow = row,
    } = options;
    if (!saved) return;
    if (previousRow?.photo_thumbnail) dropProjectThumbnailCacheEntry(previousRow.photo_thumbnail);
    if (saved?.photo_thumbnail) dropProjectThumbnailCacheEntry(saved.photo_thumbnail);
    const hasExisting = (state.projects || []).some((project) => String(project?.id) === String(saved.id));
    state.projects = hasExisting
      ? (state.projects || []).map((project) => (String(project?.id) === String(saved.id) ? saved : project))
      : [saved, ...(state.projects || [])];
    state.projectsLoadError = null;
    renderProjectsList();

    row = saved;
    hasPersistedProject = true;
    isEditing = nextEditing;
    clientPickerOpen = false;
    quickCostDraft = null;
    projectCode = (cleanNullableText(saved?.numero) || projectCode).toUpperCase();
    projectTitle = cleanNullableText(saved?.title) || "";
    projectDescription = cleanNullableText(saved?.description) || "";
    projectNotes = cleanNullableText(saved?.notes) || "";
    isPersonalProjectValue = isPersonalProject(saved);
    if (personalInput) personalInput.checked = isPersonalProjectValue;
    stages = normalizeProjectStages(saved?.stages);
    if (!stages.length) stages = [{ id: "stage-1", label: "Ouverture", description: "", done: false }];
    stageSeq = stages.length + 1;
    costLines = normalizeProjectCostLines(saved?.cost_lines);
    {
      const laborIndex = costLines.findIndex((line) => isLegacyLaborLine(line));
      if (isPersonalProjectValue) {
        costLines = costLines.filter((line) => !isLegacyLaborLine(line));
      } else if (laborIndex < 0) {
        const savedHours = Number.isFinite(Number(saved?.billable_hours)) ? Math.max(0, Number(saved.billable_hours)) : 0;
        costLines.unshift(buildDefaultLaborLine(savedHours || 1));
      } else {
        const laborLine = costLines[laborIndex];
        laborLine.source_kind = "project_labor";
        laborLine.description = "Main d'oeuvre";
        laborLine.currency = PROJECT_COST_DEFAULT_CURRENCY;
        laborLine.exchange_rate_to_cad = 1;
        laborLine.expense_date = "";
        laborLine.paypal_purchase = false;
        if (laborIndex > 0) {
          costLines.splice(laborIndex, 1);
          costLines.unshift(laborLine);
        }
      }
    }
    if (!costLines.length && !isPersonalProjectValue) costLines = [buildDefaultLaborLine(1)];
    costLines = costLines.map((line) => syncProjectCostLineDerivedValues({ ...line }));
    existingFiles = normalizeProjectFiles(saved?.files);
    existingPhotos = normalizeProjectPhotos(saved?.photos);
    const savedThumb = normalizeProjectPhotoThumbnail(saved?.photo_thumbnail);
    resetPendingThumbSelection();
    selectedThumbKey = "";
    if (savedThumb) {
      const thumbKey = `existing:${photoKey(savedThumb)}`;
      if (existingPhotos.some((photo) => `existing:${photoKey(photo)}` === thumbKey)) {
        selectedThumbKey = thumbKey;
      }
    }
    ensureThumbSelection();
    if (titleInput) titleInput.value = projectTitle;
    if (descriptionInput) descriptionInput.value = projectDescription;
    if (notesInput) notesInput.value = projectNotes;
    if (messageEl) messageEl.textContent = successMessage;
  };
  const persistInlineProjectCostDraft = async () => {
    if (saving || !hasPersistedProject || !row || !quickCostDraft) return false;
    const description = cleanNullableText(quickCostDraft.description);
    const quantity = round2(Math.max(0, parseNumberOrZero(quickCostDraft.quantity)));
    const unitPrice = round2(Math.max(0, parseNumberOrZero(quickCostDraft.unit_price)));
    const currency = normalizeProjectCostCurrency(quickCostDraft.currency);
    const expenseDate = normalizeProjectCostExpenseDate(quickCostDraft.expense_date) || formatDateYmd(new Date());
    const exchangeRate = normalizeProjectCostExchangeRate(quickCostDraft.exchange_rate_to_cad, currency);
    const paypalPurchase = normalizeProjectCostPaypalPurchase(quickCostDraft.paypal_purchase);
    if (!description) {
      if (messageEl) messageEl.textContent = "La description de la dépense est requise.";
      renderAll();
      return false;
    }
    if (currency !== PROJECT_COST_DEFAULT_CURRENCY && !(exchangeRate > 0)) {
      if (messageEl) messageEl.textContent = "Definissez un taux de change CAD valide pour cette dépense.";
      renderAll();
      return false;
    }
    const draftLine = {
      description,
      quantity,
      unit_price: unitPrice,
      amount: getProjectCostLineCadAmount({
        quantity,
        unit_price: unitPrice,
        currency,
        exchange_rate_to_cad: currency === PROJECT_COST_DEFAULT_CURRENCY ? 1 : exchangeRate,
        paypal_purchase: paypalPurchase,
      }),
      currency,
      exchange_rate_to_cad: currency === PROJECT_COST_DEFAULT_CURRENCY ? 1 : exchangeRate,
      expense_date: currency === PROJECT_COST_DEFAULT_CURRENCY ? expenseDate : expenseDate,
      labor_rate_key: null,
      paypal_purchase: paypalPurchase,
      source_kind: "custom",
    };
    const { normalizedCosts, billableHours } = normalizeProjectCostsForSave([...costLines, draftLine]);
    const developerMode = isDeveloperModeEnabled();
    const client = state.authClient;
    if (!developerMode && !client) {
      if (messageEl) messageEl.textContent = "Connexion Supabase indisponible.";
      renderAll();
      return false;
    }
    let saveSucceeded = false;
    try {
      saving = true;
      if (messageEl) messageEl.textContent = "Ajout de la dépense...";
      renderAll();
      const payload = {
        cost_lines: normalizedCosts,
        billable_hours: isPersonalProjectValue ? 0 : round2(billableHours),
        updated_at: new Date().toISOString(),
      };
      let saved = null;
      if (developerMode) {
        saved = {
          ...(row && typeof row === "object" ? row : {}),
          ...payload,
          id: Number(row.id),
        };
      } else {
        const { data, error } = await client
          .from("projects")
          .update(payload)
          .eq("id", Number(row.id))
          .select("*")
          .single();
        if (error) throw new Error(error.message || "Impossible d'ajouter la dépense.");
        saved = data;
      }
      if (!saved) throw new Error("Projet introuvable après enregistrement.");
      applySavedProjectState(saved, { successMessage: "Dépense ajoutée.", nextEditing: false, previousRow: row });
      saveSucceeded = true;
      return true;
    } catch (err) {
      if (messageEl) messageEl.textContent = err?.message || "Impossible d'ajouter la dépense.";
      return false;
    } finally {
      saving = false;
      renderAll();
      if (saveSucceeded) initialSnapshot = readDraftSnapshot();
      if (hasPersistedProject) void hydrateStoredProjectPhotoPreviews();
    }
  };
  {
    const laborIndex = costLines.findIndex((line) => isLegacyLaborLine(line));
    if (isPersonalProjectValue) {
      costLines = costLines.filter((line) => !isLegacyLaborLine(line));
    } else if (laborIndex < 0) {
      costLines.unshift(buildDefaultLaborLine(initialBillableHours || 1));
    } else {
      const laborLine = costLines[laborIndex];
      laborLine.source_kind = "project_labor";
      laborLine.description = "Main d'oeuvre";
      laborLine.currency = PROJECT_COST_DEFAULT_CURRENCY;
      laborLine.exchange_rate_to_cad = 1;
      laborLine.expense_date = "";
      const lineQty = Math.max(0, round2(parseNumberOrZero(laborLine.quantity)));
      laborLine.quantity = lineQty > 0 ? lineQty : (initialBillableHours > 0 ? initialBillableHours : 1);
      const currentRateKey = String(laborLine.labor_rate_key || "").trim();
      laborLine.labor_rate_key = laborRateOptionByKey[currentRateKey]
        ? currentRateKey
        : findLaborRateKeyByValue(laborLine.unit_price);
      if (laborLine.labor_rate_key !== customLaborRateKey) {
        laborLine.unit_price = getLaborRateValueByKey(laborLine.labor_rate_key);
      } else {
        laborLine.unit_price = Math.max(0, round2(parseNumberOrZero(laborLine.unit_price)));
      }
      laborLine.amount = getProjectCostLineCadAmount(laborLine);
      if (laborIndex > 0) {
        costLines.splice(laborIndex, 1);
        costLines.unshift(laborLine);
      }
    }
  }
  if (!costLines.length && !isPersonalProjectValue) costLines = [buildDefaultLaborLine(1)];
  costLines = costLines.map((line) => syncProjectCostLineDerivedValues({ ...line }));
  const normalizedThumb = normalizeProjectPhotoThumbnail(row?.photo_thumbnail);
  let selectedThumbKey = "";
  if (normalizedThumb) {
    const initialThumbKey = `existing:${photoKey(normalizedThumb)}`;
    if (existingPhotos.some((photo) => `existing:${photoKey(photo)}` === initialThumbKey)) {
      selectedThumbKey = initialThumbKey;
    }
  }
  let pendingThumbBlob = null;
  let pendingThumbSource = null;

  const revokePendingPhotoPreview = (url) => {
    const value = cleanNullableText(url);
    if (!value || !pendingPhotoPreviewUrls.has(value)) return;
    try {
      URL.revokeObjectURL(value);
    } catch {
      // ignore
    }
    pendingPhotoPreviewUrls.delete(value);
  };

  const openProjectPhotoPreviewByUrl = (url, titleText = "Photo du projet") => {
    const safeUrl = cleanNullableText(url);
    if (!safeUrl) return;
    openImagePreviewGallery([{ title: titleText, url: safeUrl }], 0, titleText);
  };

  const openProjectThumbnailCropper = (imageUrl, titleText = "Miniature") => new Promise((resolve) => {
    const ratio = getBonsCardAspectRatio();
    const viewportWidth = Math.min(window.innerWidth - 120, 780);
    const viewportHeight = Math.max(160, Math.round(viewportWidth / ratio));
    const root = document.createElement("div");
    root.className = "repair-thumb-crop-overlay";
    root.innerHTML = `
      <div class="repair-thumb-crop-panel" role="dialog" aria-modal="true" aria-label="Choisir la miniature">
        <div class="repair-thumb-crop-head">
          <p class="repair-thumb-crop-title">${escapeHtml(titleText)}</p>
          <button type="button" class="repair-thumb-crop-close" aria-label="Fermer">×</button>
        </div>
        <p class="repair-thumb-crop-help">Déplace l'image pour cadrer la miniature.</p>
        <div class="repair-thumb-crop-stage-wrap">
          <div class="repair-thumb-crop-stage" style="--crop-stage-w:${Math.round(viewportWidth)}px;--crop-stage-h:${Math.round(viewportHeight)}px;">
            <img class="repair-thumb-crop-image" crossorigin="anonymous" src="${escapeHtml(imageUrl)}" alt="${escapeHtml(titleText)}">
            <span class="repair-thumb-crop-frame" aria-hidden="true"></span>
          </div>
        </div>
        <div class="repair-thumb-crop-zoom-row">
          <label for="projectThumbCropZoomInput">Zoom</label>
          <input id="projectThumbCropZoomInput" class="repair-thumb-crop-zoom" type="range" min="1" max="3" step="0.01" value="1">
        </div>
        <p class="repair-thumb-crop-error" hidden></p>
        <div class="repair-thumb-crop-actions">
          <button type="button" class="repair-btn" data-crop-cancel>Annuler</button>
          <button type="button" class="repair-btn repair-btn-primary" data-crop-apply>Appliquer</button>
        </div>
      </div>
    `;
    const img = root.querySelector(".repair-thumb-crop-image");
    const stage = root.querySelector(".repair-thumb-crop-stage");
    const zoomInput = root.querySelector("#projectThumbCropZoomInput");
    const cropError = root.querySelector(".repair-thumb-crop-error");
    const applyBtn = root.querySelector("[data-crop-apply]");
    if (!(img instanceof HTMLImageElement) || !(stage instanceof HTMLElement) || !(zoomInput instanceof HTMLInputElement)) {
      resolve(null);
      return;
    }
    const setCropError = (message = "") => {
      if (!(cropError instanceof HTMLElement)) return;
      const text = cleanNullableText(message);
      cropError.hidden = !text;
      cropError.textContent = text || "";
    };
    let naturalW = 0;
    let naturalH = 0;
    let zoom = 1;
    let offsetX = 0;
    let offsetY = 0;
    let dragging = false;
    let dragStartX = 0;
    let dragStartY = 0;
    let dragOriginX = 0;
    let dragOriginY = 0;

    const computeScaled = () => {
      const baseScale = Math.min(viewportWidth / naturalW, viewportHeight / naturalH);
      const scale = baseScale * zoom;
      return {
        width: naturalW * scale,
        height: naturalH * scale,
      };
    };
    const clampOffsets = () => {
      const scaled = computeScaled();
      const maxX = Math.max(0, (scaled.width - viewportWidth) / 2);
      const maxY = Math.max(0, (scaled.height - viewportHeight) / 2);
      offsetX = Math.min(maxX, Math.max(-maxX, offsetX));
      offsetY = Math.min(maxY, Math.max(-maxY, offsetY));
    };
    const renderTransform = () => {
      if (!naturalW || !naturalH) return;
      clampOffsets();
      const scaled = computeScaled();
      img.style.width = `${scaled.width}px`;
      img.style.height = `${scaled.height}px`;
      img.style.left = `${Math.round((viewportWidth - scaled.width) / 2 + offsetX)}px`;
      img.style.top = `${Math.round((viewportHeight - scaled.height) / 2 + offsetY)}px`;
    };
    const close = (result = null) => {
      document.removeEventListener("keydown", onKeyDown, true);
      root.remove();
      resolve(result);
    };
    const onKeyDown = (ev) => {
      if (ev.key !== "Escape") return;
      ev.preventDefault();
      ev.stopPropagation();
      close(null);
    };

    const applyCrop = async () => {
      setCropError("");
      if (applyBtn instanceof HTMLButtonElement) applyBtn.disabled = true;
      try {
        if (!naturalW || !naturalH) {
          setCropError("Image non chargée. Réessaie dans une seconde.");
          return;
        }
        const outW = REPAIR_THUMBNAIL_OUTPUT_WIDTH;
        const outH = Math.max(120, Math.round(outW / ratio));
        const canvas = document.createElement("canvas");
        canvas.width = outW;
        canvas.height = outH;
        const ctx = canvas.getContext("2d");
        if (!ctx) {
          setCropError("Impossible de générer la miniature.");
          return;
        }
        const baseScale = Math.min(viewportWidth / naturalW, viewportHeight / naturalH);
        const scale = baseScale * zoom;
        const dxStage = (viewportWidth - naturalW * scale) / 2 + offsetX;
        const dyStage = (viewportHeight - naturalH * scale) / 2 + offsetY;
        const ratioX = outW / viewportWidth;
        const ratioY = outH / viewportHeight;
        const drawX = dxStage * ratioX;
        const drawY = dyStage * ratioY;
        const drawW = naturalW * scale * ratioX;
        const drawH = naturalH * scale * ratioY;
        ctx.clearRect(0, 0, outW, outH);
        ctx.drawImage(img, drawX, drawY, drawW, drawH);
        const blob = await new Promise((resolveBlob) => canvas.toBlob(resolveBlob, "image/jpeg", 0.92));
        if (!(blob instanceof Blob)) {
          setCropError("Impossible d'appliquer le cadrage de la miniature.");
          return;
        }
        close({ blob });
      } catch (err) {
        const detail = String(err?.message || err || "");
        if (/tainted|security|cross-origin/i.test(detail)) {
          setCropError("Cette photo ne peut pas être cadrée (CORS). Vérifie les règles Storage.");
        } else {
          setCropError("Impossible d'appliquer le cadrage.");
        }
      } finally {
        if (applyBtn instanceof HTMLButtonElement) applyBtn.disabled = false;
      }
    };

    img.addEventListener("load", () => {
      naturalW = Number(img.naturalWidth || 0);
      naturalH = Number(img.naturalHeight || 0);
      renderTransform();
    }, { once: true });
    if (img.complete) {
      naturalW = Number(img.naturalWidth || 0);
      naturalH = Number(img.naturalHeight || 0);
      renderTransform();
    }

    zoomInput.addEventListener("input", () => {
      zoom = Math.max(1, Math.min(3, Number(zoomInput.value || 1)));
      renderTransform();
    });

    stage.addEventListener("pointerdown", (ev) => {
      dragging = true;
      dragStartX = ev.clientX;
      dragStartY = ev.clientY;
      dragOriginX = offsetX;
      dragOriginY = offsetY;
      stage.setPointerCapture?.(ev.pointerId);
    });
    stage.addEventListener("pointermove", (ev) => {
      if (!dragging) return;
      offsetX = dragOriginX + (ev.clientX - dragStartX);
      offsetY = dragOriginY + (ev.clientY - dragStartY);
      renderTransform();
    });
    const stopDrag = () => { dragging = false; };
    stage.addEventListener("pointerup", stopDrag);
    stage.addEventListener("pointercancel", stopDrag);
    stage.addEventListener("pointerleave", stopDrag);

    root.addEventListener("click", async (ev) => {
      if (ev.target === root || ev.target?.closest?.(".repair-thumb-crop-close") || ev.target?.closest?.("[data-crop-cancel]")) {
        close(null);
        return;
      }
      if (ev.target?.closest?.("[data-crop-apply]")) {
        ev.preventDefault();
        ev.stopPropagation();
        await applyCrop();
      }
    });
    document.addEventListener("keydown", onKeyDown, true);
    document.body.appendChild(root);
  });

  const ensureThumbSelection = () => {
    const existingKeys = new Set(existingPhotos.map((photo) => `existing:${photoKey(photo)}`));
    const pendingKeys = new Set(pendingPhotos.map((photo) => `pending:${photo.pending_id}`));
    if (selectedThumbKey && (existingKeys.has(selectedThumbKey) || pendingKeys.has(selectedThumbKey))) return;
    selectedThumbKey = "";
  };

  const resetPendingThumbSelection = () => {
    pendingThumbBlob = null;
    pendingThumbSource = null;
  };
  ensureThumbSelection();

  const progressSummary = () => getProjectProgressSummary(stages);
  const canEmitProjectInvoice = () => {
    if (isPersonalProjectValue) return false;
    if (!hasPersistedProject || !row) return false;
    return true;
  };
  const setClientFromId = (clientId) => {
    const key = String(clientId ?? "").trim();
    if (!key || !clientById[key]) {
      selectedClientId = "";
      selectedClientTitle = "";
      selectedClientCompany = "";
      return;
    }
    const selected = clientById[key];
    selectedClientId = key;
    selectedClientTitle = formatClientPrimaryName(selected);
    selectedClientCompany = cleanNullableText(selected?.company) || "";
  };

  const renderProjectFactureAction = () => {
    if (!factureActionBtn) return;
    if (isPersonalProjectValue) {
      factureActionBtn.hidden = true;
      factureActionBtn.disabled = true;
      return;
    }
    const facture = row ? getFactureForProject(row) : null;
    const canEmit = canEmitProjectInvoice();
    const isIssued = Boolean(facture);
    if (!isIssued && !canEmit) {
      factureActionBtn.hidden = true;
      factureActionBtn.disabled = true;
      factureActionBtn.textContent = "Émettre la facture";
      factureActionBtn.classList.remove("is-issued");
      factureActionBtn.classList.add("is-not-issued");
      return;
    }

    factureActionBtn.hidden = false;
    factureActionBtn.disabled = saving || isEditing;
    const factureNumber = formatFactureNumber(facture);
    const label = isIssued
      ? (factureNumber ? `Facture émise ${factureNumber}` : "Facture émise")
      : "Émettre la facture";
    factureActionBtn.textContent = label;
    factureActionBtn.classList.toggle("is-issued", isIssued);
    factureActionBtn.classList.toggle("is-not-issued", !isIssued);
  };

  const getStoredProjectPhotoPreviewUrl = async (photo) => {
    if (!photo) throw new Error("Photo introuvable.");
    const key = photoKey(photo);
    let previewUrl = photoPreviewUrlByKey[key];
    if (!previewUrl) {
      if (!state.authClient) throw new Error("Connexion Supabase indisponible.");
      previewUrl = await getProjectPhotoUrl(state.authClient, photo);
      photoPreviewUrlByKey[key] = previewUrl;
    }
    return previewUrl;
  };

  const buildProjectPhotoGalleryEntries = () => {
    const storedEntries = normalizeProjectPhotos(existingPhotos)
      .map((photo, index) => ({
        kind: "stored",
        index,
        title: cleanNullableText(photo.original_name) || photo.filename.split("/").pop() || "Photo du projet",
        getUrl: () => getStoredProjectPhotoPreviewUrl(photo),
      }));
    const pendingEntries = pendingPhotos
      .map((photo, index) => ({
        kind: "pending",
        index,
        title: cleanNullableText(photo.original_name) || `Photo ${index + 1}`,
        getUrl: () => Promise.resolve(cleanNullableText(photo.preview_url)),
      }));
    return [...storedEntries, ...pendingEntries];
  };

  const openProjectPhotoGallery = (kind, index) => {
    const entries = buildProjectPhotoGalleryEntries();
    const startIndex = entries.findIndex((entry) => entry.kind === kind && entry.index === index);
    if (startIndex < 0) return;
    openImagePreviewGallery(entries, startIndex, "Photos du projet");
  };

  const selectStoredProjectThumbnail = async (index) => {
    const photo = normalizeProjectPhotos(existingPhotos)[index];
    if (!photo) return false;
    const key = photoKey(photo);
    let previewUrl = photoPreviewUrlByKey[key];
    if (!previewUrl) {
      if (!state.authClient) throw new Error("Connexion Supabase indisponible.");
      previewUrl = await getProjectPhotoUrl(state.authClient, photo);
      photoPreviewUrlByKey[key] = previewUrl;
    }
    const label = cleanNullableText(photo.original_name) || photo.filename.split("/").pop() || "Miniature";
    const cropped = await openProjectThumbnailCropper(previewUrl, label);
    if (!cropped?.blob) return false;
    selectedThumbKey = `existing:${key}`;
    pendingThumbBlob = cropped.blob;
    pendingThumbSource = {
      kind: "stored",
      key: selectedThumbKey,
      filename: photo.filename,
      bucket: photo.bucket || REPAIRS_PHOTOS_BUCKET,
      original_name: cleanNullableText(photo.original_name),
      content_type: cleanNullableText(photo.content_type),
    };
    return true;
  };

  const selectPendingProjectThumbnail = async (index) => {
    const pending = pendingPhotos[index];
    if (!pending?.preview_url) return false;
    const label = cleanNullableText(pending.original_name) || `Miniature ${index + 1}`;
    const cropped = await openProjectThumbnailCropper(pending.preview_url, label);
    if (!cropped?.blob) return false;
    selectedThumbKey = `pending:${pending.pending_id}`;
    pendingThumbBlob = cropped.blob;
    pendingThumbSource = {
      kind: "pending",
      key: selectedThumbKey,
      filename: cleanNullableText(pending.original_name) || cleanNullableText(pending.file?.name),
      bucket: REPAIRS_PHOTOS_BUCKET,
      original_name: cleanNullableText(pending.original_name),
      content_type: cleanNullableText(pending.file?.type),
    };
    return true;
  };

  const hydrateStoredProjectPhotoPreviews = async () => {
    if (!state.authClient) return;
    const rows = normalizeProjectPhotos(existingPhotos);
    let updated = false;
    for (const photo of rows) {
      const key = photoKey(photo);
      if (photoPreviewUrlByKey[key]) continue;
      try {
        const url = await getProjectPhotoUrl(state.authClient, photo);
        if (url) {
          photoPreviewUrlByKey[key] = url;
          updated = true;
        }
      } catch {
        // ignore
      }
    }
    if (updated) renderPhotos();
  };
  const openProjectFile = async (index) => {
    const fileRow = normalizeProjectFiles(existingFiles)[index];
    if (!fileRow) return;
    const developerMode = isDeveloperModeEnabled();
    if (!developerMode && !state.authClient) throw new Error("Connexion Supabase indisponible.");
    const url = developerMode && fileRow.preview_url
      ? fileRow.preview_url
      : await getProjectFileUrl(state.authClient, fileRow);
    if (!url) throw new Error("Impossible d'ouvrir le fichier.");
    const name = cleanNullableText(fileRow.original_name) || fileRow.filename.split("/").pop() || "Fichier";
    const contentType = String(fileRow.content_type || "").toLowerCase();
    const isPdf = contentType.includes("pdf") || /\.pdf$/i.test(String(fileRow.filename || ""));
    const isImage = contentType.startsWith("image/") || /\.(png|jpe?g|gif|webp|bmp|svg)$/i.test(String(fileRow.filename || ""));
    if (isPdf) {
      const openedInModal = await openFacturePdfModal(url, `Apercu PDF - ${name}`);
      if (openedInModal) return;
      const opened = openPdfInNewTab(url);
      if (opened) return;
      throw new Error("Impossible d'ouvrir ce PDF.");
    }
    if (isImage) {
      openProjectPhotoPreviewByUrl(url, name);
      return;
    }
    window.open(url, "_blank", "noopener,noreferrer");
  };
  const downloadProjectFile = async (index) => {
    const fileRow = normalizeProjectFiles(existingFiles)[index];
    if (!fileRow) return;
    const developerMode = isDeveloperModeEnabled();
    if (!developerMode && !state.authClient) throw new Error("Connexion Supabase indisponible.");
    const fileName = cleanNullableText(fileRow.original_name) || fileRow.filename.split("/").pop() || "Fichier";
    const url = developerMode && fileRow.preview_url
      ? fileRow.preview_url
      : await getProjectFileDownloadUrl(state.authClient, fileRow);
    if (!url) throw new Error("Impossible de télécharger le fichier.");
    const anchor = document.createElement("a");
    anchor.href = url;
    anchor.download = fileName;
    anchor.rel = "noopener noreferrer";
    anchor.style.display = "none";
    document.body.appendChild(anchor);
    anchor.click();
    anchor.remove();
  };
  const saveProjectFiles = async (nextFiles, successMessage = "Fichier ajouté.") => {
    if (saving || !hasPersistedProject || !row) return false;
    const normalizedFiles = normalizeProjectFiles(nextFiles);
    const developerMode = isDeveloperModeEnabled();
    const client = state.authClient;
    if (!developerMode && !client) {
      if (filesMessage) filesMessage.textContent = "Connexion Supabase indisponible.";
      return false;
    }
    let saveSucceeded = false;
    try {
      saving = true;
      if (filesMessage) filesMessage.textContent = "Enregistrement des fichiers...";
      renderAll();
      const payload = {
        files: normalizedFiles,
        updated_at: new Date().toISOString(),
      };
      let saved = null;
      if (developerMode) {
        saved = {
          ...(row && typeof row === "object" ? row : {}),
          ...payload,
          id: Number(row.id),
        };
      } else {
        const { data, error } = await client
          .from("projects")
          .update(payload)
          .eq("id", Number(row.id))
          .select("*")
          .single();
        if (error) throw new Error(error.message || "Impossible d'enregistrer les fichiers.");
        saved = data;
      }
      if (!saved) throw new Error("Projet introuvable après enregistrement.");
      applySavedProjectState(saved, { successMessage, nextEditing: false, previousRow: row });
      if (filesMessage) filesMessage.textContent = "";
      saveSucceeded = true;
      return true;
    } catch (err) {
      if (filesMessage) filesMessage.textContent = err?.message || "Impossible d'enregistrer les fichiers.";
      return false;
    } finally {
      saving = false;
      renderAll();
      if (saveSucceeded) initialSnapshot = readDraftSnapshot();
      if (hasPersistedProject) void hydrateStoredProjectPhotoPreviews();
    }
  };

  body.innerHTML = `
    <form id="projectForm" class="facture-form project-form">
      <section id="projectTitleSection" class="facture-section project-title-section" hidden>
        <h4 class="facture-section-title">Titre du projet</h4>
        <input id="projectTitleInput" class="facture-input" type="text" value="${escapeHtml(projectTitle)}" placeholder="Ex.: Préamp 8 entrées custom" aria-label="Titre du projet">
      </section>

      <section class="facture-section project-client-section">
        <label class="personal-work-toggle">
          <input id="projectPersonalInput" type="checkbox" ${isPersonalProjectValue ? "checked" : ""}>
          <span>Projet personnel</span>
        </label>
        <h4 id="projectClientSectionTitle" class="facture-section-title">Client</h4>
        <div id="projectClientView" class="project-client-view" hidden>
          <p id="projectClientViewName" class="project-client-view-name">Sans client</p>
          <div id="projectClientViewDetails" class="project-client-view-details"></div>
        </div>
        <div id="projectClientEditWrap">
          <div class="facture-client-controls">
            <button id="projectClientAddBtn" type="button" class="repair-client-add-btn">Ajouter un client</button>
            <div id="projectClientBadge" class="facture-client-badge" hidden>
              <span id="projectClientBadgeName" class="facture-client-badge-name"></span>
              <span id="projectClientBadgeCompany" class="facture-client-badge-company" hidden></span>
              <button id="projectClientBadgeRemoveBtn" type="button" class="facture-client-badge-remove" aria-label="Retirer le client" title="Retirer le client" hidden>×</button>
            </div>
          </div>
          <div id="projectClientPicker" class="facture-picker" hidden>
            <input id="projectClientSearchInput" class="facture-input" type="search" placeholder="Rechercher un client">
            <div id="projectClientList" class="facture-picker-list"></div>
          </div>
        </div>
      </section>

      <section class="facture-section project-description-section">
        <h4 id="projectDescriptionSectionTitle" class="facture-section-title">Description du projet</h4>
        <p id="projectDescriptionView" class="project-description-view"></p>
        <textarea id="projectDescriptionInput" class="facture-textarea">${escapeHtml(projectDescription)}</textarea>
      </section>

      <section class="facture-section">
        <h4 class="facture-section-title">Progression</h4>
        <div class="project-progress-inline">
          <div class="project-thermometer project-thermometer-modal">
            <span id="projectProgressFill" class="project-thermometer-fill" style="width:0%;"></span>
            <span id="projectProgressPercent" class="project-progress-percent">0%</span>
          </div>
        </div>
        <div id="projectStageAddRow" class="project-stage-add-row">
          <input id="projectStageAddTitleInput" class="facture-input" type="text" placeholder="Titre de l'étape">
          <input id="projectStageAddDescriptionInput" class="facture-input project-stage-add-description-input" type="text" placeholder="Description de l'étape">
          <button id="projectStageAddBtn" type="button" class="facture-add-line-btn">Ajouter</button>
        </div>
        <div id="projectStagesList" class="project-stages-list"></div>
      </section>

      <section class="facture-section">
        <input id="projectPhotosInput" type="file" accept="image/*" multiple hidden>
        <div class="repair-photos-head">
          <h4 class="facture-section-title">Photos</h4>
          <button id="projectPhotosAddBtn" type="button" class="repair-client-add-btn repair-photos-add-btn">Ajouter</button>
        </div>
        <div id="projectPhotosGrid" class="repair-photos-grid"></div>
        <p id="projectPhotosMessage" class="facture-message"></p>
      </section>

      <section class="facture-section">
        <h4 class="facture-section-title">Coûts</h4>
        <div class="facture-lines-table-wrap repair-cost-lines-table-wrap project-cost-lines-table-wrap">
          <table class="facture-lines-table repair-cost-lines-table project-cost-table project-cost-lines-table">
            <thead>
              <tr>
                <th class="project-cost-header-desc">Description</th>
                <th class="project-cost-header-unit">Prix unitaire</th>
                <th class="project-cost-header-currency">Devise</th>
                <th class="project-cost-header-qty">Qté</th>
                <th class="project-cost-header-fx">Taux</th>
                <th class="project-cost-header-total">Total (CAD)</th>
              </tr>
            </thead>
            <tbody id="projectCostLinesTbody"></tbody>
          </table>
        </div>
        <div class="facture-lines-actions repair-cost-lines-actions">
          <button id="projectCostAddLineBtn" type="button" class="facture-add-line-btn">Ajouter</button>
          <p id="projectCostSubtotal" class="project-cost-lines-subtotal repair-cost-lines-subtotal">Sous-total: 0,00 $</p>
        </div>
      </section>

      <section class="facture-section project-notes-section">
        <h4 class="facture-section-title">Notes</h4>
        <p id="projectNotesView" class="project-description-view"></p>
        <textarea id="projectNotesInput" class="facture-textarea" placeholder="Notes internes">${escapeHtml(projectNotes)}</textarea>
      </section>

      <section class="facture-section project-files-section">
        <input id="projectFilesInput" type="file" multiple hidden>
        <div class="repair-photos-head">
          <h4 class="facture-section-title">Fichiers</h4>
          <button id="projectFilesAddBtn" type="button" class="repair-client-add-btn repair-photos-add-btn">Ajouter</button>
        </div>
        <div id="projectFilesList" class="project-files-list"></div>
        <p id="projectFilesMessage" class="facture-message"></p>
      </section>

      <div class="facture-actions">
        <p id="projectMessage" class="facture-message"></p>
        <div class="facture-actions-right">
          <button id="projectCancelBtn" type="button" class="repair-btn">Annuler</button>
          <button id="projectSaveBtn" type="submit" class="repair-btn repair-btn-primary">Enregistrer</button>
        </div>
      </div>
    </form>
  `;

  const progressFill = $("projectProgressFill");
  const progressPercent = $("projectProgressPercent");
  const clientSectionTitle = $("projectClientSectionTitle");
  const clientView = $("projectClientView");
  const clientViewName = $("projectClientViewName");
  const clientViewDetails = $("projectClientViewDetails");
  const clientEditWrap = $("projectClientEditWrap");
  const clientAddBtn = $("projectClientAddBtn");
  const clientBadge = $("projectClientBadge");
  const clientBadgeName = $("projectClientBadgeName");
  const clientBadgeCompany = $("projectClientBadgeCompany");
  const clientBadgeRemoveBtn = $("projectClientBadgeRemoveBtn");
  const clientPicker = $("projectClientPicker");
  const clientSearchInput = $("projectClientSearchInput");
  const clientList = $("projectClientList");
  const personalInput = $("projectPersonalInput");
  const personalToggle = personalInput?.closest?.(".personal-work-toggle") || null;
  const titleSection = $("projectTitleSection");
  const titleInput = $("projectTitleInput");
  const stageAddRow = $("projectStageAddRow");
  const stageAddTitleInput = $("projectStageAddTitleInput");
  const stageAddDescriptionInput = $("projectStageAddDescriptionInput");
  const stageAddBtn = $("projectStageAddBtn");
  const stagesList = $("projectStagesList");
  const photosInput = $("projectPhotosInput");
  const photosAddBtn = $("projectPhotosAddBtn");
  const photosGrid = $("projectPhotosGrid");
  const photosMessage = $("projectPhotosMessage");
  const filesInput = $("projectFilesInput");
  const filesAddBtn = $("projectFilesAddBtn");
  const filesList = $("projectFilesList");
  const filesMessage = $("projectFilesMessage");
  const descriptionSectionTitle = $("projectDescriptionSectionTitle");
  const descriptionView = $("projectDescriptionView");
  const descriptionInput = $("projectDescriptionInput");
  const notesView = $("projectNotesView");
  const notesInput = $("projectNotesInput");
  const costLinesTbody = $("projectCostLinesTbody");
  const addCostLineBtn = $("projectCostAddLineBtn");
  const costSubtotalEl = $("projectCostSubtotal");
  const cancelBtn = $("projectCancelBtn");
  const saveBtn = $("projectSaveBtn");
  const form = $("projectForm");
  const messageEl = $("projectMessage");
  const editBtn = $("projectModalEditBtn");
  const factureActionBtn = $("projectFactureActionBtn");

  const renderClientPicker = (query = "") => {
    if (!clientList) return;
    const q = normalizeText(query);
    const filtered = sortedClients.filter((client) => {
      if (!q) return true;
      const hay = `${formatClientPrimaryName(client)} ${client.company || ""} ${client.email || ""} ${client.phone || ""}`;
      return normalizeText(hay).includes(q);
    });
    const createRow = `
      <button type="button" class="facture-picker-item facture-picker-item-create" data-client-action="create">
        <span class="facture-picker-item-name">+ Créer un nouveau client</span>
      </button>
    `;
    clientList.innerHTML = filtered.length
      ? `${createRow}${filtered.map((client) => `
        <button type="button" class="facture-picker-item" data-client-id="${escapeHtml(String(getClientReferenceId(client) || ""))}">
          <span class="facture-picker-item-name">${escapeHtml(formatClientPrimaryName(client))}</span>
          ${cleanNullableText(client.company) ? `<span class="facture-picker-item-sub">(${escapeHtml(cleanNullableText(client.company))})</span>` : ""}
        </button>
      `).join("")}`
      : `${createRow}<p class="repair-client-picker-empty">Aucun client trouvé.</p>`;
  };

  const renderClientBadge = () => {
    const hasClient = Boolean(cleanNullableText(selectedClientId) || cleanNullableText(selectedClientTitle));
    const selectedClientRow = (() => {
      const key = String(selectedClientId || "").trim();
      if (!key) return null;
      return clientById[key] || state.clientsByPodioItemId?.[key] || null;
    })();
    const displayName = cleanNullableText(selectedClientTitle)
      || formatClientPrimaryName(selectedClientRow)
      || "Sans client";
    const displayCompany = cleanNullableText(selectedClientCompany)
      || cleanNullableText(selectedClientRow?.company)
      || "";
    const canEditClient = isEditing && !saving;
    if (clientAddBtn) {
      clientAddBtn.hidden = !isEditing || hasClient || isPersonalProjectValue;
      clientAddBtn.disabled = !canEditClient || hasClient || isPersonalProjectValue;
    }
    if (clientPicker) clientPicker.hidden = !(clientPickerOpen && canEditClient && !hasClient && !isPersonalProjectValue);
    if (clientSearchInput) clientSearchInput.disabled = !canEditClient || !clientPickerOpen || hasClient || isPersonalProjectValue;
    if (clientBadgeRemoveBtn) {
      clientBadgeRemoveBtn.hidden = !isEditing || !hasClient || isPersonalProjectValue;
      clientBadgeRemoveBtn.disabled = !canEditClient || !hasClient || isPersonalProjectValue;
    }
    if (!clientBadge || !clientBadgeName || !clientBadgeCompany) return;
    clientBadge.hidden = !hasClient;
    clientBadgeName.textContent = hasClient ? displayName : "";
    const company = hasClient ? displayCompany : null;
    clientBadgeCompany.hidden = !company;
    clientBadgeCompany.textContent = company ? `(${company})` : "";

    if (clientViewName) clientViewName.textContent = hasClient ? displayName : "Sans client";
    if (clientViewDetails) {
      if (!hasClient) {
        clientViewDetails.innerHTML = "";
      } else {
        const details = [
          company || null,
          cleanNullableText(selectedClientRow?.address),
          cleanNullableText(selectedClientRow?.phone),
          cleanNullableText(selectedClientRow?.email),
        ].filter(Boolean);
        clientViewDetails.innerHTML = details
          .map((line) => `<p class="project-client-view-line">${escapeHtml(String(line))}</p>`)
          .join("");
      }
    }
  };

  const renderHeader = () => {
    const progress = progressSummary();
    const percent = Math.max(0, Math.min(100, Number(progress.percent || 0)));
    const phase = cleanNullableText(progress.phaseLabel) || "Ouverture";
    const personalLabelHtml = isPersonalProjectValue
      ? `<span class="modal-personal-kicker">Projet personnel</span>`
      : "";
    if (progressPercent) progressPercent.textContent = `${percent}%`;
    if (progressFill) progressFill.style.width = `${percent}%`;
    const titleLabel = projectTitle || "Projet";
    title.innerHTML = `
      ${personalLabelHtml}
      <span class="project-modal-title-mainline">
        <span class="project-modal-title-code">${escapeHtml(projectCode)}</span>
        <span class="repair-modal-title-sep">|</span>
        <span class="project-modal-title-name">${escapeHtml(titleLabel)}</span>
        <span class="repair-modal-title-sep">|</span>
        <span class="project-modal-title-phase">${escapeHtml(`${phase} ${percent}%`)}</span>
      </span>
    `;
    if (editBtn) {
      editBtn.hidden = !hasPersistedProject;
      editBtn.disabled = saving || !hasPersistedProject;
      editBtn.classList.toggle("is-active", isEditing);
      editBtn.setAttribute("aria-pressed", String(isEditing));
    }
    renderProjectFactureAction();
  };

  const renderStages = () => {
    if (!stagesList) return;
    stagesList.innerHTML = stages.map((stage, stageIndex) => {
      const isLockedStage = stageIndex === 0 || stageIndex === stages.length - 1
        || isProjectOpeningStageLabel(stage?.label)
        || isProjectEmitInvoiceStageLabel(stage?.label);
      const description = cleanNullableText(stage?.description) || "";
      return `
      <div class="project-stage-item${stage.done ? " is-done" : ""}" data-stage-id="${escapeHtml(stage.id)}">
        <div class="project-stage-main">
          <div class="project-stage-copy${isEditing ? " is-editing" : ""}">
            ${isEditing
              ? `
                <input type="text" class="facture-input project-stage-label-input" value="${escapeHtml(stage.label)}" placeholder="Titre" ${saving || isLockedStage ? "disabled" : ""}>
                <input type="text" class="facture-input project-stage-description-input" value="${escapeHtml(description)}" placeholder="Description de l'étape" ${saving ? "disabled" : ""}>
              `
              : `
                <span class="project-stage-title-view">${escapeHtml(stage.label || "Etape")}</span>
                ${description ? `<span class="project-stage-description-view">${escapeHtml(description)}</span>` : ""}
              `}
          </div>
          <span class="project-stage-guide" aria-hidden="true"></span>
          ${isEditing
            ? `
              <div class="project-stage-actions">
                <button type="button" class="project-stage-move-btn" data-stage-move-id="${escapeHtml(stage.id)}" data-stage-move-dir="up" ${saving || isLockedStage || stageIndex <= 1 ? "disabled" : ""} aria-label="Monter l'étape" title="Monter">▲</button>
                <button type="button" class="project-stage-move-btn" data-stage-move-id="${escapeHtml(stage.id)}" data-stage-move-dir="down" ${saving || isLockedStage || stageIndex >= (stages.length - 2) ? "disabled" : ""} aria-label="Descendre l'étape" title="Descendre">▼</button>
                <button type="button" class="facture-line-remove" data-stage-remove-id="${escapeHtml(stage.id)}" ${saving || stages.length <= 2 || isLockedStage ? "disabled" : ""}>×</button>
              </div>
            `
            : ""}
          <label class="project-stage-check" aria-label="Étape complétée">
            <input type="checkbox" class="project-stage-done-input" ${stage.done ? "checked" : ""} ${saving ? "disabled" : ""}>
          </label>
        </div>
      </div>
    `;
    }).join("");
  };

  const renderPhotos = () => {
    if (!photosGrid) return;
    ensureThumbSelection();
    const rows = [
      ...existingPhotos.map((photo, index) => ({
        kind: "stored",
        index,
        key: `existing:${photoKey(photo)}`,
        label: cleanNullableText(photo.original_name) || photo.filename,
        previewUrl: photoPreviewUrlByKey[photoKey(photo)] || "",
      })),
      ...pendingPhotos.map((photo, index) => ({
        kind: "pending",
        index,
        key: `pending:${photo.pending_id}`,
        label: cleanNullableText(photo.original_name) || `Nouvelle photo ${photo.pending_id}`,
        previewUrl: cleanNullableText(photo.preview_url) || "",
      })),
    ];
    if (!rows.length) {
      photosGrid.innerHTML = `<p class="repair-client-picker-empty">Aucune photo.</p>`;
      return;
    }
    photosGrid.innerHTML = rows.map((entry) => `
      <div class="repair-photo-item${entry.kind === "pending" ? " repair-photo-item-pending" : ""}${selectedThumbKey === entry.key ? " is-selected-thumb" : ""}">
        <button type="button" class="repair-photo-thumb-btn" data-project-photo-open="1" data-project-photo-kind="${escapeHtml(entry.kind)}" data-project-photo-index="${escapeHtml(String(entry.index))}" title="Voir la photo">
          ${entry.previewUrl
            ? `<img class="repair-photo-thumb" src="${escapeHtml(entry.previewUrl)}" alt="${escapeHtml(entry.label)}">`
            : `<span class="repair-photo-thumb-placeholder" aria-hidden="true">${REPAIR_THUMB_ICON_SVG}</span>`}
        </button>
        ${isEditing
          ? `
            <div class="repair-photo-actions">
              <button type="button" class="repair-photo-thumb-select-btn${selectedThumbKey === entry.key ? " is-active" : ""}" data-project-photo-thumb="1" data-project-photo-kind="${escapeHtml(entry.kind)}" data-project-photo-index="${escapeHtml(String(entry.index))}" ${saving ? "disabled" : ""} title="Choisir comme miniature">${REPAIR_THUMB_ICON_SVG}</button>
            </div>
            <button type="button" class="repair-photo-remove-btn" data-project-photo-remove="1" data-project-photo-kind="${escapeHtml(entry.kind)}" data-project-photo-index="${escapeHtml(String(entry.index))}" ${saving ? "disabled" : ""}>×</button>
          `
          : ""}
      </div>
    `).join("");
  };

  const renderFiles = () => {
    if (!filesList) return;
    const rows = normalizeProjectFiles(existingFiles);
    if (!rows.length) {
      filesList.innerHTML = `<p class="repair-client-picker-empty">Aucun fichier.</p>`;
      return;
    }
    filesList.innerHTML = rows.map((file, index) => {
      const label = cleanNullableText(file.original_name) || file.filename.split("/").pop() || "Fichier";
      const sizeLabel = Number.isFinite(Number(file.size)) && Number(file.size) > 0 ? formatBytes(Number(file.size)) : "";
      return `
        <div class="project-file-item">
          <span class="project-file-item-main">
            <span class="project-file-item-head">
              <span class="project-file-item-name">${escapeHtml(label)}</span>
              ${sizeLabel ? `<span class="project-file-item-meta">${escapeHtml(sizeLabel)}</span>` : ""}
            </span>
          </span>
          <span class="project-file-item-guide" aria-hidden="true"></span>
          <span class="project-file-item-actions">
            <button type="button" class="project-file-action-btn is-preview" data-project-file-open="${escapeHtml(String(index))}" title="Visualiser le fichier" aria-label="Visualiser le fichier">${FACTURE_PREVIEW_EYE_SVG}</button>
            <button type="button" class="project-file-action-btn" data-project-file-download="${escapeHtml(String(index))}" title="Télécharger le fichier" aria-label="Télécharger le fichier">${PROJECT_FILE_DOWNLOAD_SVG}</button>
          </span>
        </div>
      `;
    }).join("");
  };

  const getProjectCostLinesSubtotal = () => {
    const currentSubtotal = costLines.reduce((sum, line) => sum + getProjectCostLineCadAmount(line), 0);
    const draftSubtotal = quickCostDraft ? getProjectCostLineCadAmount(quickCostDraft) : 0;
    return round2(currentSubtotal + draftSubtotal);
  };

  const renderProjectCostSubtotal = () => {
    if (!costSubtotalEl) return;
    costSubtotalEl.textContent = `Sous-total: ${formatMoney(getProjectCostLinesSubtotal())}`;
  };

  const renderCostLines = () => {
    if (!costLinesTbody) return;
    const rowsHtml = costLines.map((line, idx) => {
      const laborLine = isProjectLaborLine(line);
      syncProjectCostLineDerivedValues(line);
      if (laborLine) {
        line.source_kind = "project_labor";
        line.description = "Main d'oeuvre";
        const currentRateKey = String(line.labor_rate_key || "").trim();
        const selectedRateKey = (Boolean(laborRateOptionByKey[currentRateKey]) || currentRateKey === customLaborRateKey)
          ? currentRateKey
          : findLaborRateKeyByValue(line.unit_price);
        line.labor_rate_key = selectedRateKey;
        if (selectedRateKey !== customLaborRateKey) {
          line.unit_price = getLaborRateValueByKey(selectedRateKey);
        }
      }
      syncProjectCostLineDerivedValues(line);
      const currency = normalizeProjectCostCurrency(line.currency);
      const expenseDate = normalizeProjectCostExpenseDate(line.expense_date);
      const exchangeRate = getProjectCostLineExchangeRate(line);
      const paypalPurchase = normalizeProjectCostPaypalPurchase(line.paypal_purchase);
      const fxStamp = formatProjectCostFxStamp(currency, exchangeRate, expenseDate, { paypalPurchase });
      const total = getProjectCostLineCadAmount(line);
      const rateControl = laborLine
        ? `
          <div class="project-cost-rate-wrap repair-cost-rate-wrap">
            <div class="project-cost-rate-select-wrap repair-cost-rate-select-wrap">
              <span class="project-cost-rate-collapsed repair-cost-rate-collapsed">${escapeHtml(`${formatMoneyCompact(line.unit_price)}/h`)}</span>
              <select class="facture-input project-cost-rate-select repair-line-rate-select" ${saving || !isEditing ? "disabled" : ""}>
                ${buildLaborRateOptionsHtml(String(line.labor_rate_key || defaultLaborRateKey))}
              </select>
            </div>
            ${String(line.labor_rate_key || "") === customLaborRateKey
              ? `<input type="number" min="0" step="0.01" class="facture-input project-cost-unit-input repair-line-unit-input project-cost-number-input" value="${escapeHtml(round2(line.unit_price))}" ${saving || !isEditing ? "disabled" : ""}>`
              : ""}
          </div>
        `
        : (isEditing
          ? `<input type="number" min="0" step="0.01" class="facture-input project-cost-unit-input repair-line-unit-input project-cost-number-input" value="${escapeHtml(round2(line.unit_price))}" ${saving ? "disabled" : ""}>`
          : `<span class="project-cost-cell-value project-cost-cell-value-right">${escapeHtml(round2(line.unit_price))}</span>`);
      const currencyControl = laborLine
        ? `<span class="project-cost-currency-static">CAD</span>`
        : (isEditing
          ? `
            <div class="project-cost-currency-wrap">
              <div class="project-cost-currency-main">
                <select class="facture-input project-cost-currency-select repair-line-currency-select" ${saving ? "disabled" : ""}>
                  ${buildProjectCurrencyOptionsHtml(currency)}
                </select>
                <label class="project-cost-date-picker${expenseDate ? " is-set" : ""}" title="${escapeHtml(expenseDate ? `Date de depense: ${expenseDate}` : "Choisir la date de depense")}">
                  <span class="project-cost-date-btn" aria-hidden="true">${PROJECT_COST_DATE_SVG}</span>
                  <input type="date" class="project-cost-date-input" value="${escapeHtml(expenseDate)}" ${saving ? "disabled" : ""}>
                </label>
              </div>
            </div>
          `
          : `<span class="project-cost-currency-static">${escapeHtml(currency)}</span>`);
      const fxControl = laborLine
        ? `<span class="project-cost-fx-empty"></span>`
        : (isEditing
          ? `
            <div class="project-cost-fx-wrap">
              <div class="project-cost-fx-stamp">${escapeHtml(fxStamp)}</div>
              ${currency !== PROJECT_COST_DEFAULT_CURRENCY
                ? buildProjectCostPaypalToggleHtml("project-cost-paypal-checkbox", paypalPurchase, saving)
                : ""}
            </div>
          `
          : `<div class="project-cost-fx-stamp">${escapeHtml(fxStamp)}</div>`);
      const descriptionControl = isEditing
        ? `
          <div class="project-cost-desc-wrap repair-line-desc-wrap">
            ${laborLine
              ? ""
              : `<button type="button" class="facture-line-remove project-cost-inline-remove" data-cost-line-remove-id="${escapeHtml(String(idx))}" ${saving ? "disabled" : ""}>×</button>`}
            <input type="text" class="facture-input project-cost-desc-input" value="${escapeHtml(line.description || "")}" ${saving || laborLine ? "disabled" : ""}>
          </div>
        `
        : `<span class="project-cost-cell-value project-cost-cell-value-left">${escapeHtml(line.description || "")}</span>`;
      return `
        <tr data-cost-line-id="${escapeHtml(String(idx))}">
          <td class="project-cost-cell-desc">${descriptionControl}</td>
          <td class="project-cost-cell-unit">${rateControl}</td>
          <td class="project-cost-cell-currency">${currencyControl}</td>
          <td class="project-cost-cell-qty">${isEditing
            ? `<input type="number" min="0" step="0.01" class="facture-input project-cost-qty-input repair-line-qty-input project-cost-number-input" value="${escapeHtml(round2(line.quantity))}" ${saving ? "disabled" : ""}>`
            : `<span class="project-cost-cell-value project-cost-cell-value-right">${escapeHtml(round2(line.quantity))}</span>`}</td>
          <td class="project-cost-cell-fx">${fxControl}</td>
          <td class="project-cost-cell-total"><span class="facture-line-total">${escapeHtml(formatMoney(total))}</span></td>
        </tr>
      `;
    });
    if (quickCostDraft) {
      syncProjectCostLineDerivedValues(quickCostDraft);
      const quantity = round2(Math.max(0, parseNumberOrZero(quickCostDraft.quantity)));
      const unitPrice = round2(Math.max(0, parseNumberOrZero(quickCostDraft.unit_price)));
      const currency = normalizeProjectCostCurrency(quickCostDraft.currency);
      const expenseDate = normalizeProjectCostExpenseDate(quickCostDraft.expense_date);
      const exchangeRate = getProjectCostLineExchangeRate(quickCostDraft);
      const paypalPurchase = normalizeProjectCostPaypalPurchase(quickCostDraft.paypal_purchase);
      const fxStamp = formatProjectCostFxStamp(currency, exchangeRate, expenseDate, { paypalPurchase });
      const total = getProjectCostLineCadAmount(quickCostDraft);
      rowsHtml.push(`
        <tr class="project-cost-draft-row" data-cost-draft-row="1">
          <td class="project-cost-cell-desc">
            <div class="project-cost-desc-wrap repair-cost-draft-desc-wrap">
              <div class="project-cost-draft-actions repair-cost-draft-actions-left">
                <button type="button" class="project-cost-draft-save-btn" data-cost-draft-action="save" ${saving ? "disabled" : ""} aria-label="Enregistrer la ligne" title="Enregistrer">✓</button>
                <button type="button" class="facture-line-remove" data-cost-draft-action="cancel" ${saving ? "disabled" : ""}>×</button>
              </div>
              <input type="text" class="facture-input project-cost-draft-desc-input" value="${escapeHtml(quickCostDraft.description || "")}" placeholder="Description" ${saving ? "disabled" : ""}>
            </div>
          </td>
          <td class="project-cost-cell-unit"><input type="number" min="0" step="0.01" class="facture-input project-cost-draft-unit-input repair-cost-draft-unit-input project-cost-number-input" value="${escapeHtml(unitPrice)}" ${saving ? "disabled" : ""}></td>
          <td class="project-cost-cell-currency">
            <div class="project-cost-currency-wrap">
              <div class="project-cost-currency-main">
                <select class="facture-input project-cost-draft-currency-select repair-cost-draft-currency-select" ${saving ? "disabled" : ""}>
                  ${buildProjectCurrencyOptionsHtml(currency)}
                </select>
                <label class="project-cost-date-picker${expenseDate ? " is-set" : ""}" title="${escapeHtml(expenseDate ? `Date de depense: ${expenseDate}` : "Choisir la date de depense")}">
                  <span class="project-cost-date-btn" aria-hidden="true">${PROJECT_COST_DATE_SVG}</span>
                  <input type="date" class="project-cost-date-input project-cost-draft-date-input" value="${escapeHtml(expenseDate)}" ${saving ? "disabled" : ""}>
                </label>
              </div>
            </div>
          </td>
          <td class="project-cost-cell-qty"><input type="number" min="0" step="0.01" class="facture-input project-cost-draft-qty-input repair-cost-draft-qty-input project-cost-number-input" value="${escapeHtml(quantity)}" ${saving ? "disabled" : ""}></td>
          <td class="project-cost-cell-fx">
            <div class="project-cost-fx-wrap">
              <div class="project-cost-fx-stamp">${escapeHtml(fxStamp)}</div>
              ${currency !== PROJECT_COST_DEFAULT_CURRENCY
                ? buildProjectCostPaypalToggleHtml("project-cost-draft-paypal-checkbox", paypalPurchase, saving)
                : ""}
            </div>
          </td>
          <td class="project-cost-cell-total"><span class="facture-line-total">${escapeHtml(formatMoney(total))}</span></td>
        </tr>
      `);
    }
    costLinesTbody.innerHTML = rowsHtml.join("");
    renderProjectCostSubtotal();
  };

  const readDraftSnapshot = () => JSON.stringify({
    client: selectedClientId,
    is_personal: isPersonalProjectValue,
    title: cleanNullableText(titleInput?.value) || "",
    description: cleanNullableText(descriptionInput?.value) || "",
    notes: cleanNullableText(notesInput?.value) || "",
    stages: stages.map((stage) => ({
      label: cleanNullableText(stage.label) || "",
      description: cleanNullableText(stage.description) || "",
      done: stage.done === true,
    })),
    costs: costLines.map((line) => ({
      description: cleanNullableText(line.description) || "",
      quantity: round2(parseNumberOrZero(line.quantity)),
      unit_price: round2(parseNumberOrZero(line.unit_price)),
      currency: normalizeProjectCostCurrency(line.currency),
      exchange_rate_to_cad: getProjectCostLineBaseExchangeRate(line),
      expense_date: normalizeProjectCostExpenseDate(line.expense_date),
      labor_rate_key: cleanNullableText(line.labor_rate_key),
      paypal_purchase: normalizeProjectCostPaypalPurchase(line.paypal_purchase),
      source_kind: cleanNullableText(line.source_kind),
    })),
    files: serializeProjectFiles(existingFiles),
    photos: existingPhotos.map((photo) => photoKey(photo)),
    pending: pendingPhotos.map((photo) => cleanNullableText(photo.original_name) || photo.pending_id),
    thumb: selectedThumbKey,
    thumb_pending: Boolean(pendingThumbBlob),
    thumb_source: cleanNullableText(pendingThumbSource?.key) || "",
    quick_cost_draft: quickCostDraft
      ? {
        description: cleanNullableText(quickCostDraft.description) || "",
        quantity: round2(parseNumberOrZero(quickCostDraft.quantity)),
        unit_price: round2(parseNumberOrZero(quickCostDraft.unit_price)),
        currency: normalizeProjectCostCurrency(quickCostDraft.currency),
        exchange_rate_to_cad: getProjectCostLineBaseExchangeRate(quickCostDraft),
        expense_date: normalizeProjectCostExpenseDate(quickCostDraft.expense_date),
        paypal_purchase: normalizeProjectCostPaypalPurchase(quickCostDraft.paypal_purchase),
      }
      : null,
  });
  let initialSnapshot = readDraftSnapshot();

  const renderAll = () => {
    projectTitle = cleanNullableText(titleInput?.value) || "";
    projectDescription = cleanNullableText(descriptionInput?.value) || "";
    projectNotes = cleanNullableText(notesInput?.value) || "";
    modal.classList.toggle("is-project-editing", isEditing);
    if (titleSection) titleSection.hidden = !isEditing;
    if (clientSectionTitle) clientSectionTitle.hidden = false;
    if (clientView) {
      clientView.hidden = isEditing;
      clientView.style.display = isEditing ? "none" : "";
    }
    if (personalInput) {
      personalInput.checked = isPersonalProjectValue;
      personalInput.disabled = saving || !isEditing;
    }
    if (personalToggle instanceof HTMLElement) {
      personalToggle.hidden = !isEditing;
      personalToggle.style.display = isEditing ? "" : "none";
    }
    if (clientEditWrap) clientEditWrap.hidden = !isEditing;
    if (descriptionSectionTitle) descriptionSectionTitle.hidden = false;
    if (descriptionView) {
      descriptionView.hidden = isEditing;
      descriptionView.textContent = projectDescription || "-";
    }
    if (descriptionInput) descriptionInput.hidden = !isEditing;
    if (notesView) {
      notesView.hidden = isEditing;
      notesView.textContent = projectNotes || "-";
    }
    if (notesInput) {
      notesInput.hidden = !isEditing;
      notesInput.disabled = saving || !isEditing;
    }
    renderHeader();
    renderClientBadge();
    if (clientPickerOpen && clientSearchInput && !selectedClientId) renderClientPicker(clientSearchInput.value);
    renderStages();
    renderPhotos();
    renderFiles();
    renderCostLines();
    if (stageAddRow) {
      stageAddRow.hidden = !isEditing;
      stageAddRow.style.display = isEditing ? "" : "none";
    }
    if (titleInput) titleInput.disabled = saving || !isEditing;
    if (stageAddTitleInput) stageAddTitleInput.disabled = saving || !isEditing;
    if (stageAddDescriptionInput) stageAddDescriptionInput.disabled = saving || !isEditing;
    if (stageAddBtn) stageAddBtn.disabled = saving || !isEditing;
    if (photosAddBtn) {
      photosAddBtn.hidden = !isEditing;
      photosAddBtn.disabled = saving || !isEditing;
    }
    if (photosInput) photosInput.disabled = saving || !isEditing;
    if (filesAddBtn) {
      filesAddBtn.hidden = !hasPersistedProject;
      filesAddBtn.disabled = saving || !hasPersistedProject || isEditing;
    }
    if (filesInput) filesInput.disabled = saving || !hasPersistedProject || isEditing;
    if (descriptionInput) descriptionInput.disabled = saving || !isEditing;
    if (addCostLineBtn) {
      const allowQuickAdd = hasPersistedProject && !isEditing;
      addCostLineBtn.hidden = !(isEditing || allowQuickAdd);
      addCostLineBtn.disabled = saving || Boolean(quickCostDraft);
      addCostLineBtn.textContent = isEditing ? "Ajouter un coût" : "Ajouter";
    }
    if (cancelBtn) {
      cancelBtn.disabled = saving;
      cancelBtn.textContent = isEditing ? "Annuler" : "Fermer";
    }
    if (saveBtn) saveBtn.disabled = saving;
  };

  const applyPersonalProjectSelection = () => {
    const personalOwner = getPersonalOwnerSelection();
    selectedClientId = personalOwner.itemId != null ? String(personalOwner.itemId) : "";
    selectedClientTitle = personalOwner.title;
    selectedClientCompany = personalOwner.company || "";
  };

  if (personalInput) {
    personalInput.addEventListener("change", () => {
      isPersonalProjectValue = personalInput.checked;
      if (isPersonalProjectValue) {
        applyPersonalProjectSelection();
        costLines = ensureProjectLaborCostLines(costLines, 0, getLaborRateValueByKey(defaultLaborRateKey), { includeLabor: false });
      } else {
        costLines = ensureProjectLaborCostLines(costLines, 0, getLaborRateValueByKey(defaultLaborRateKey), { includeLabor: true });
      }
      clientPickerOpen = false;
      renderAll();
    });
  }

  if (editBtn) {
    editBtn.onclick = () => {
      if (saving || !hasPersistedProject) return;
      if (quickCostDraft) {
        if (messageEl) messageEl.textContent = "Enregistrez ou annulez la ligne ajoutée avant de changer de mode.";
        return;
      }
      isEditing = !isEditing;
      if (!isEditing) clientPickerOpen = false;
      if (messageEl) messageEl.textContent = "";
      renderAll();
    };
  }

  if (factureActionBtn) {
    factureActionBtn.onclick = async (ev) => {
      ev.preventDefault();
      ev.stopPropagation();
      if (saving || !row || !hasPersistedProject) return;
      const existingFacture = getFactureForProject(row);
      if (existingFacture) {
        try {
          await openFacturePdfPreview(existingFacture);
        } catch (err) {
          window.alert(err?.message || "Impossible d'ouvrir le PDF.");
        }
        return;
      }
      if (!canEmitProjectInvoice()) return;
      await promptProjectFactureCreation(row);
      renderAll();
    };
  }

  if (cancelBtn) {
    cancelBtn.addEventListener("click", async () => {
      await closeProjectModal();
    });
  }
  if (clientAddBtn) {
    clientAddBtn.addEventListener("click", () => {
      if (saving || !isEditing || isPersonalProjectValue) return;
      clientPickerOpen = !clientPickerOpen;
      if (clientPickerOpen && clientSearchInput) renderClientPicker(clientSearchInput.value);
      renderAll();
    });
  }
  if (clientSearchInput) clientSearchInput.addEventListener("input", () => renderClientPicker(clientSearchInput.value));
  if (clientBadgeRemoveBtn) {
    clientBadgeRemoveBtn.addEventListener("click", () => {
      if (saving || !isEditing || isPersonalProjectValue) return;
      setClientFromId("");
      clientPickerOpen = false;
      renderAll();
    });
  }
  if (clientList) {
    clientList.addEventListener("click", (ev) => {
      if (saving || !isEditing || !(ev.target instanceof Element)) return;
      const createBtn = ev.target.closest("[data-client-action='create']");
      if (createBtn) {
        openCreateClientModal({
          onCreated: (createdClient) => {
            rebuildClients();
            const createdItemId = getClientReferenceId(createdClient);
            if (Number.isInteger(createdItemId)) setClientFromId(String(createdItemId));
            clientPickerOpen = false;
            renderAll();
          },
        });
        return;
      }
      const btnClient = ev.target.closest("[data-client-id]");
      if (!btnClient) return;
      setClientFromId(String(btnClient.getAttribute("data-client-id") || ""));
      clientPickerOpen = false;
      renderAll();
    });
  }
  if (stageAddBtn && stageAddTitleInput) {
    const addStage = () => {
      if (saving || !isEditing) return;
      const label = cleanNullableText(stageAddTitleInput.value);
      const description = cleanNullableText(stageAddDescriptionInput?.value);
      if (!label) return;
      const insertAt = Math.max(1, stages.length - 1);
      stages.splice(insertAt, 0, {
        id: `stage-${stageSeq++}`,
        label,
        description: description || "",
        done: false,
      });
      stageAddTitleInput.value = "";
      if (stageAddDescriptionInput) stageAddDescriptionInput.value = "";
      renderAll();
    };
    stageAddBtn.addEventListener("click", addStage);
    const handleStageAddKeydown = (ev) => {
      if (ev.key !== "Enter") return;
      ev.preventDefault();
      addStage();
    };
    stageAddTitleInput.addEventListener("keydown", handleStageAddKeydown);
    if (stageAddDescriptionInput) stageAddDescriptionInput.addEventListener("keydown", handleStageAddKeydown);
  }
  if (stagesList) {
    stagesList.addEventListener("input", (ev) => {
      if (saving) return;
      if (!(ev.target instanceof Element)) return;
      const rowEl = ev.target.closest("[data-stage-id]");
      if (!rowEl) return;
      const stageId = String(rowEl.getAttribute("data-stage-id") || "");
      const stage = stages.find((item) => String(item.id) === stageId);
      if (!stage) return;
      const stageIndex = stages.findIndex((item) => String(item.id) === stageId);
      const isLockedStage = stageIndex === 0 || stageIndex === stages.length - 1
        || isProjectOpeningStageLabel(stage?.label)
        || isProjectEmitInvoiceStageLabel(stage?.label);
      const target = ev.target;
      if (!(target instanceof HTMLInputElement)) return;
      if (target.classList.contains("project-stage-done-input")) {
        stage.done = Boolean(target.checked);
        rowEl.classList.toggle("is-done", stage.done === true);
        renderHeader();
        return;
      }
      if (target.classList.contains("project-stage-label-input")) {
        if (!isEditing || isLockedStage) return;
        stage.label = target.value;
        renderHeader();
        return;
      }
      if (target.classList.contains("project-stage-description-input")) {
        if (!isEditing) return;
        stage.description = target.value;
      }
    });
    stagesList.addEventListener("click", (ev) => {
      if (saving || !isEditing || !(ev.target instanceof Element)) return;
      const removeBtn = ev.target.closest("[data-stage-remove-id]");
      if (!removeBtn) return;
      const stageId = String(removeBtn.getAttribute("data-stage-remove-id") || "");
      const stageIndex = stages.findIndex((stage) => String(stage.id) === stageId);
      if (stageIndex < 0) return;
      const stage = stages[stageIndex];
      const isLockedStage = stageIndex === 0 || stageIndex === stages.length - 1
        || isProjectOpeningStageLabel(stage?.label)
        || isProjectEmitInvoiceStageLabel(stage?.label);
      if (stages.length <= 2 || isLockedStage) return;
      stages = stages.filter((stage) => String(stage.id) !== stageId);
      renderAll();
    });
    stagesList.addEventListener("click", (ev) => {
      if (saving || !isEditing || !(ev.target instanceof Element)) return;
      const moveBtn = ev.target.closest("[data-stage-move-id]");
      if (!moveBtn) return;
      const stageId = String(moveBtn.getAttribute("data-stage-move-id") || "");
      const dir = String(moveBtn.getAttribute("data-stage-move-dir") || "");
      const index = stages.findIndex((stage) => String(stage.id) === stageId);
      if (index < 0) return;
      const stage = stages[index];
      const isLockedStage = index === 0 || index === stages.length - 1
        || isProjectOpeningStageLabel(stage?.label)
        || isProjectEmitInvoiceStageLabel(stage?.label);
      if (isLockedStage) return;
      const nextIndex = dir === "up" ? index - 1 : index + 1;
      if (nextIndex < 1 || nextIndex >= (stages.length - 1)) return;
      const [moved] = stages.splice(index, 1);
      stages.splice(nextIndex, 0, moved);
      renderAll();
    });
  }
  if (photosAddBtn && photosInput) {
    photosAddBtn.addEventListener("click", () => {
      if (saving || !isEditing) return;
      photosInput.click();
    });
  }
  if (photosInput) {
    photosInput.addEventListener("change", () => {
      if (saving || !isEditing) return;
      const files = Array.from(photosInput.files || []).filter((file) => file instanceof File);
      for (const file of files) {
        let previewUrl = "";
        try {
          previewUrl = URL.createObjectURL(file);
        } catch {
          previewUrl = "";
        }
        pendingPhotos.push({
          pending_id: `pending-${Date.now()}-${pendingSeq++}`,
          file,
          original_name: cleanNullableText(file.name),
          preview_url: previewUrl,
        });
        if (previewUrl) pendingPhotoPreviewUrls.add(previewUrl);
      }
      photosInput.value = "";
      ensureThumbSelection();
      renderAll();
    });
  }
  if (filesAddBtn && filesInput) {
    filesAddBtn.addEventListener("click", () => {
      if (saving || !hasPersistedProject || isEditing) return;
      if (filesMessage) filesMessage.textContent = "";
      filesInput.click();
    });
  }
  if (filesInput) {
    filesInput.addEventListener("change", async () => {
      if (saving || !hasPersistedProject || isEditing) return;
      const files = Array.from(filesInput.files || []).filter((file) => file instanceof File);
      filesInput.value = "";
      if (!files.length) return;
      const developerMode = isDeveloperModeEnabled();
      const client = state.authClient;
      if (!developerMode && !client) {
        if (filesMessage) filesMessage.textContent = "Connexion Supabase indisponible.";
        return;
      }
      const uploadedFiles = [];
      try {
        let seq = existingFiles.length + 1;
        for (const file of files) {
          let uploaded = null;
          if (developerMode) {
            let previewUrl = "";
            try {
              previewUrl = URL.createObjectURL(file);
            } catch {
              previewUrl = "";
            }
            uploaded = {
              filename: buildProjectFileObjectPath(projectCode, file.name || "fichier", seq++),
              bucket: PROJECT_FILES_BUCKET,
              original_name: cleanNullableText(file.name),
              content_type: cleanNullableText(file.type),
              size: Number.isFinite(Number(file.size)) ? Number(file.size) : null,
              created_at: new Date().toISOString(),
              preview_url: previewUrl || null,
            };
          } else {
            uploaded = await uploadProjectFileAttachment(client, projectCode, file, seq++);
          }
          uploadedFiles.push(uploaded);
        }
      } catch (err) {
        if (filesMessage) filesMessage.textContent = err?.message || "Impossible de televerser le fichier.";
        return;
      }
      await saveProjectFiles([...existingFiles, ...uploadedFiles], uploadedFiles.length > 1 ? "Fichiers ajoutés." : "Fichier ajouté.");
    });
  }
  if (photosGrid) {
    photosGrid.addEventListener("click", async (ev) => {
      if (saving) return;
      if (!(ev.target instanceof Element)) return;
      const removeBtn = ev.target.closest("[data-project-photo-remove]");
      if (removeBtn && isEditing) {
        const kind = String(removeBtn.getAttribute("data-project-photo-kind") || "");
        const index = Number(removeBtn.getAttribute("data-project-photo-index"));
        if (!Number.isInteger(index) || index < 0) return;
        let removedThumbKey = "";
        if (kind === "stored") {
          const removed = existingPhotos[index];
          if (!removed) return;
          removedThumbKey = `existing:${photoKey(removed)}`;
          existingPhotos.splice(index, 1);
          dropProjectThumbnailCacheEntry(removed);
        } else if (kind === "pending") {
          const removed = pendingPhotos[index];
          if (!removed) return;
          removedThumbKey = `pending:${removed.pending_id}`;
          pendingPhotos.splice(index, 1);
          revokePendingPhotoPreview(removed.preview_url);
        }
        if (removedThumbKey && selectedThumbKey === removedThumbKey) {
          selectedThumbKey = "";
          resetPendingThumbSelection();
        } else if (removedThumbKey && pendingThumbSource?.key === removedThumbKey) {
          resetPendingThumbSelection();
        }
        ensureThumbSelection();
        if (photosMessage) photosMessage.textContent = "";
        renderAll();
        return;
      }

      const thumbBtn = ev.target.closest("[data-project-photo-thumb]");
      if (thumbBtn && isEditing) {
        const kind = String(thumbBtn.getAttribute("data-project-photo-kind") || "");
        const index = Number(thumbBtn.getAttribute("data-project-photo-index"));
        if (!Number.isInteger(index) || index < 0) return;
        try {
          let selected = false;
          if (kind === "stored") {
            selected = await selectStoredProjectThumbnail(index);
          } else if (kind === "pending") {
            selected = await selectPendingProjectThumbnail(index);
          }
          if (selected) {
            if (photosMessage) photosMessage.textContent = "Miniature prête à être enregistrée.";
            renderPhotos();
          }
        } catch (err) {
          if (photosMessage) photosMessage.textContent = err?.message || "Impossible de sélectionner la miniature.";
        }
        return;
      }

      const openBtn = ev.target.closest("[data-project-photo-open]");
      if (!openBtn) return;
      const kind = String(openBtn.getAttribute("data-project-photo-kind") || "");
      const index = Number(openBtn.getAttribute("data-project-photo-index"));
      if (!Number.isInteger(index) || index < 0) return;
      try {
        if (kind === "stored" || kind === "pending") {
          openProjectPhotoGallery(kind, index);
          return;
        }
      } catch (err) {
        if (photosMessage) photosMessage.textContent = err?.message || "Impossible d'ouvrir la photo.";
      }
    });
  }
  if (filesList) {
    filesList.addEventListener("click", async (ev) => {
      if (saving || !(ev.target instanceof Element)) return;
      const downloadBtn = ev.target.closest("[data-project-file-download]");
      if (downloadBtn) {
        const index = Number(downloadBtn.getAttribute("data-project-file-download"));
        if (!Number.isInteger(index) || index < 0) return;
        try {
          if (filesMessage) filesMessage.textContent = "";
          await downloadProjectFile(index);
        } catch (err) {
          if (filesMessage) filesMessage.textContent = err?.message || "Impossible de télécharger le fichier.";
        }
        return;
      }
      const openBtn = ev.target.closest("[data-project-file-open]");
      if (!openBtn) return;
      const index = Number(openBtn.getAttribute("data-project-file-open"));
      if (!Number.isInteger(index) || index < 0) return;
      try {
        if (filesMessage) filesMessage.textContent = "";
        await openProjectFile(index);
      } catch (err) {
        if (filesMessage) filesMessage.textContent = err?.message || "Impossible d'ouvrir le fichier.";
      }
    });
  }
  if (addCostLineBtn) {
    addCostLineBtn.addEventListener("click", () => {
      if (saving) return;
      if (isEditing) {
        costLines.push({
          description: "",
          quantity: 1,
          unit_price: 0,
          amount: 0,
          currency: PROJECT_COST_DEFAULT_CURRENCY,
          exchange_rate_to_cad: 1,
          expense_date: formatDateYmd(new Date()),
          labor_rate_key: null,
          source_kind: "custom",
        });
        renderAll();
        return;
      }
      if (!hasPersistedProject || quickCostDraft) return;
      quickCostDraft = buildEmptyQuickCostDraft();
      renderAll();
      requestAnimationFrame(() => {
        const input = body.querySelector(".project-cost-draft-desc-input");
        if (input instanceof HTMLElement) input.focus();
      });
    });
  }
  if (costLinesTbody) {
    costLinesTbody.addEventListener("input", (ev) => {
      if (saving || !(ev.target instanceof Element)) return;
      if (quickCostDraft) {
        const draftRow = ev.target.closest("tr[data-cost-draft-row]");
        if (draftRow) {
          const target = ev.target;
          if (!(target instanceof HTMLInputElement)) return;
          if (target.classList.contains("project-cost-draft-desc-input")) quickCostDraft.description = target.value;
          if (target.classList.contains("project-cost-draft-unit-input")) quickCostDraft.unit_price = Math.max(0, parseNumberOrZero(target.value));
          if (target.classList.contains("project-cost-draft-qty-input")) quickCostDraft.quantity = Math.max(0, parseNumberOrZero(target.value));
          if (target.classList.contains("project-cost-draft-fx-input")) {
            quickCostDraft.exchange_rate_to_cad = normalizeProjectCostExchangeRate(target.value, quickCostDraft.currency);
          }
          syncProjectCostLineDerivedValues(quickCostDraft);
          const totalEl = draftRow.querySelector(".facture-line-total");
          if (totalEl) {
            totalEl.textContent = formatMoney(getProjectCostLineCadAmount(quickCostDraft));
          }
          renderProjectCostSubtotal();
          return;
        }
      }
      if (!isEditing) return;
      const rowEl = ev.target.closest("tr[data-cost-line-id]");
      if (!rowEl) return;
      const idx = Number(rowEl.getAttribute("data-cost-line-id"));
      if (!Number.isInteger(idx) || idx < 0 || idx >= costLines.length) return;
      const line = costLines[idx];
      const target = ev.target;
      if (!(target instanceof HTMLInputElement)) return;
      if (target.classList.contains("project-cost-desc-input")) line.description = target.value;
      if (target.classList.contains("project-cost-unit-input")) {
        line.unit_price = Math.max(0, parseNumberOrZero(target.value));
        if (isProjectLaborLine(line)) line.labor_rate_key = customLaborRateKey;
      }
      if (target.classList.contains("project-cost-qty-input")) line.quantity = Math.max(0, parseNumberOrZero(target.value));
      if (target.classList.contains("project-cost-fx-input")) {
        line.exchange_rate_to_cad = normalizeProjectCostExchangeRate(target.value, line.currency);
      }
      syncProjectCostLineDerivedValues(line);
      const totalEl = rowEl.querySelector(".facture-line-total");
      if (totalEl) totalEl.textContent = formatMoney(getProjectCostLineCadAmount(line));
      renderProjectCostSubtotal();
    });
    costLinesTbody.addEventListener("change", async (ev) => {
      if (saving || !(ev.target instanceof Element)) return;
      if (quickCostDraft) {
        const draftRow = ev.target.closest("tr[data-cost-draft-row]");
        if (draftRow) {
          const target = ev.target;
          if (target instanceof HTMLSelectElement && target.classList.contains("project-cost-draft-currency-select")) {
            quickCostDraft.currency = normalizeProjectCostCurrency(target.value);
            if (quickCostDraft.currency === PROJECT_COST_DEFAULT_CURRENCY) {
              quickCostDraft.exchange_rate_to_cad = 1;
            } else {
              await fetchAndApplyProjectCostExchangeRate(quickCostDraft);
            }
            renderCostLines();
            return;
          }
          if (target instanceof HTMLInputElement && target.classList.contains("project-cost-draft-date-input")) {
            quickCostDraft.expense_date = normalizeProjectCostExpenseDate(target.value) || formatDateYmd(new Date());
            if (normalizeProjectCostCurrency(quickCostDraft.currency) !== PROJECT_COST_DEFAULT_CURRENCY) {
              await fetchAndApplyProjectCostExchangeRate(quickCostDraft);
            } else {
              syncProjectCostLineDerivedValues(quickCostDraft);
            }
            renderCostLines();
            return;
          }
          if (target instanceof HTMLInputElement && target.classList.contains("project-cost-draft-paypal-checkbox")) {
            quickCostDraft.paypal_purchase = target.checked;
            syncProjectCostLineDerivedValues(quickCostDraft);
            renderCostLines();
            return;
          }
        }
      }
      if (!isEditing) return;
      const rowEl = ev.target.closest("tr[data-cost-line-id]");
      if (!rowEl) return;
      const idx = Number(rowEl.getAttribute("data-cost-line-id"));
      if (!Number.isInteger(idx) || idx < 0 || idx >= costLines.length) return;
      const line = costLines[idx];
      const target = ev.target;
      if (target instanceof HTMLSelectElement && target.classList.contains("project-cost-rate-select")) {
        if (!isProjectLaborLine(line)) return;
        const nextRateKey = String(target.value || "").trim();
        if (!(laborRateOptionByKey[nextRateKey] || nextRateKey === customLaborRateKey)) return;
        line.labor_rate_key = nextRateKey;
        if (nextRateKey !== customLaborRateKey) {
          line.unit_price = getLaborRateValueByKey(nextRateKey);
        }
        syncProjectCostLineDerivedValues(line);
        renderCostLines();
        return;
      }
      if (target instanceof HTMLSelectElement && target.classList.contains("project-cost-currency-select")) {
        line.currency = normalizeProjectCostCurrency(target.value);
        if (line.currency === PROJECT_COST_DEFAULT_CURRENCY) {
          line.exchange_rate_to_cad = 1;
          syncProjectCostLineDerivedValues(line);
        } else {
          await fetchAndApplyProjectCostExchangeRate(line);
        }
        renderCostLines();
        return;
      }
      if (target instanceof HTMLInputElement && target.classList.contains("project-cost-date-input")) {
        line.expense_date = normalizeProjectCostExpenseDate(target.value) || formatDateYmd(new Date());
        if (normalizeProjectCostCurrency(line.currency) !== PROJECT_COST_DEFAULT_CURRENCY) {
          await fetchAndApplyProjectCostExchangeRate(line);
        } else {
          syncProjectCostLineDerivedValues(line);
        }
        renderCostLines();
        return;
      }
      if (target instanceof HTMLInputElement && target.classList.contains("project-cost-paypal-checkbox")) {
        line.paypal_purchase = target.checked;
        syncProjectCostLineDerivedValues(line);
        renderCostLines();
      }
    });
    costLinesTbody.addEventListener("click", async (ev) => {
      if (saving || !(ev.target instanceof Element)) return;
      const dateTriggerBtn = ev.target.closest("[data-cost-date-trigger]");
      if (dateTriggerBtn) {
        const targetKind = String(dateTriggerBtn.getAttribute("data-cost-date-trigger") || "");
        const rowEl = dateTriggerBtn.closest("tr");
        const input = rowEl?.querySelector(targetKind === "draft" ? ".project-cost-draft-date-input" : ".project-cost-date-input");
        if (input instanceof HTMLInputElement) {
          if (typeof input.showPicker === "function") input.showPicker();
          else input.click();
        }
        return;
      }
      const draftActionBtn = ev.target.closest("[data-cost-draft-action]");
      if (draftActionBtn) {
        const action = String(draftActionBtn.getAttribute("data-cost-draft-action") || "");
        if (action === "cancel") {
          quickCostDraft = null;
          if (messageEl) messageEl.textContent = "";
          renderAll();
          return;
        }
        if (action === "save") {
          await persistInlineProjectCostDraft();
        }
        return;
      }
      if (!isEditing) return;
      const removeBtn = ev.target.closest("[data-cost-line-remove-id]");
      if (!removeBtn) return;
      const idx = Number(removeBtn.getAttribute("data-cost-line-remove-id"));
      if (!Number.isInteger(idx) || idx < 0 || idx >= costLines.length) return;
      if (isProjectLaborLine(costLines[idx])) return;
      costLines.splice(idx, 1);
      if (!costLines.length && !isPersonalProjectValue) costLines.push(buildDefaultLaborLine(1));
      renderAll();
    });
  }

  const saveProjectChanges = async () => {
    if (saving) return null;
    if (quickCostDraft && !isEditing) {
      const savedInline = await persistInlineProjectCostDraft();
      return savedInline ? (row || true) : null;
    }
    if (quickCostDraft) {
      if (messageEl) messageEl.textContent = "Enregistrez ou annulez la ligne ajoutée avant d'enregistrer le projet.";
      renderAll();
      return null;
    }
    if (!isEditing && readDraftSnapshot() === initialSnapshot) {
      if (messageEl) messageEl.textContent = "Aucune modification à enregistrer.";
      return row || true;
    }
    projectTitle = cleanNullableText(titleInput?.value) || "";
    projectDescription = cleanNullableText(descriptionInput?.value) || "";
    projectNotes = cleanNullableText(notesInput?.value) || "";
    if (!isPersonalProjectValue && !selectedClientId) {
      if (messageEl) messageEl.textContent = "Sélectionnez un client avant d'enregistrer.";
      return null;
    }
    if (!projectTitle) {
      if (messageEl) messageEl.textContent = "Le titre du projet est requis.";
      return null;
    }
    const invalidFxLine = costLines.find((line) => {
      const currency = normalizeProjectCostCurrency(line.currency);
      if (currency === PROJECT_COST_DEFAULT_CURRENCY) return false;
      return !(normalizeProjectCostExchangeRate(line.exchange_rate_to_cad, currency) > 0);
    });
    if (invalidFxLine) {
      if (messageEl) messageEl.textContent = "Definissez un taux de change CAD valide pour chaque depense en devise etrangere.";
      return null;
    }
    const progress = progressSummary();
    const { normalizedCosts, billableHours } = normalizeProjectCostsForSave(costLines);
    const developerMode = isDeveloperModeEnabled();
    const client = state.authClient;
    if (!developerMode && !client) {
      if (messageEl) messageEl.textContent = "Connexion Supabase indisponible.";
      return null;
    }
    let saveSucceeded = false;
    let savedProject = null;
    try {
      saving = true;
      if (messageEl) messageEl.textContent = "Enregistrement du projet...";
      renderAll();
      const uploadedByPendingId = new Map();
      const uploadedPhotos = [];
      let photoSeq = existingPhotos.length + 1;
      for (const pending of pendingPhotos) {
        let uploaded = null;
        if (developerMode) {
          uploaded = {
            filename: buildProjectPhotoObjectPath(
              projectCode,
              photoSeq++,
              pending?.file?.name || pending?.original_name || "photo.jpg",
              pending?.file?.type || "",
            ),
            bucket: REPAIRS_PHOTOS_BUCKET,
            original_name: cleanNullableText(pending?.file?.name || pending?.original_name),
            content_type: cleanNullableText(pending?.file?.type),
            size: Number.isFinite(Number(pending?.file?.size)) ? Number(pending.file.size) : null,
            created_at: new Date().toISOString(),
          };
        } else {
          uploaded = await uploadProjectPhotoFile(client, projectCode, pending.file, photoSeq++);
        }
        uploadedByPendingId.set(String(pending.pending_id), uploaded);
        uploadedPhotos.push(uploaded);
      }
      const finalPhotos = normalizeProjectPhotos([...existingPhotos, ...uploadedPhotos]);
      let finalThumb = normalizeProjectPhotoThumbnail(row?.photo_thumbnail);
      let finalThumbSourcePhoto = null;
      if (selectedThumbKey.startsWith("existing:")) {
        const wanted = selectedThumbKey.replace(/^existing:/, "");
        finalThumbSourcePhoto = finalPhotos.find((photo) => photoKey(photo) === wanted) || null;
      } else if (selectedThumbKey.startsWith("pending:")) {
        const pendingId = selectedThumbKey.replace(/^pending:/, "");
        const uploaded = uploadedByPendingId.get(pendingId);
        if (uploaded) {
          const wanted = `${uploaded.bucket || REPAIRS_PHOTOS_BUCKET}/${uploaded.filename}`;
          finalThumbSourcePhoto = finalPhotos.find((photo) => photoKey(photo) === wanted) || null;
        }
      }
      if (pendingThumbBlob instanceof Blob) {
        if (developerMode) {
          finalThumb = {
            filename: buildProjectThumbnailObjectPath(projectCode),
            bucket: REPAIRS_PHOTOS_BUCKET,
            original_name: cleanNullableText(finalThumbSourcePhoto?.original_name) || "Miniature",
            content_type: "image/jpeg",
            source_filename: cleanNullableText(finalThumbSourcePhoto?.filename),
            size: Number.isFinite(Number(pendingThumbBlob.size)) ? Number(pendingThumbBlob.size) : null,
            created_at: new Date().toISOString(),
          };
        } else {
          finalThumb = await uploadProjectThumbnailBlob(client, projectCode, pendingThumbBlob, finalThumbSourcePhoto);
        }
      } else if (finalThumbSourcePhoto) {
        finalThumb = finalThumbSourcePhoto;
      } else if (finalThumb) {
        const thumbKey = `${finalThumb.bucket || REPAIRS_PHOTOS_BUCKET}/${finalThumb.filename}`;
        const thumbStillExists = finalPhotos.some((photo) => photoKey(photo) === thumbKey);
        const thumbIsDedicatedFile = /_thumb\.[a-z0-9]+$/i.test(String(finalThumb.filename || ""));
        if (!thumbStillExists && !thumbIsDedicatedFile) {
          finalThumb = null;
        }
      }
      if (!finalThumb && finalPhotos.length) finalThumb = finalPhotos[0];

      const payload = {
        numero: projectCode,
        client_item_id: selectedClientId ? Number(selectedClientId) : null,
        client_title: selectedClientTitle || null,
        is_personal: isPersonalProjectValue,
        title: projectTitle,
        description: projectDescription,
        notes: projectNotes,
        completion_percent: Math.max(0, Math.min(100, Number(progress.percent || 0))),
        current_stage_label: cleanNullableText(progress.phaseLabel) || "Ouverture",
        stages: progress.stages.map((stage) => ({
          id: stage.id,
          label: stage.label,
          description: cleanNullableText(stage.description) || "",
          done: stage.done === true,
        })),
        files: normalizeProjectFiles(existingFiles),
        photos: finalPhotos,
        photo_thumbnail: finalThumb || null,
        cost_lines: normalizedCosts,
        billable_hours: isPersonalProjectValue ? 0 : round2(billableHours),
        updated_at: new Date().toISOString(),
      };

      const previousRow = row ? { ...row } : null;
      if (developerMode) {
        const nowIso = new Date().toISOString();
        savedProject = {
          ...(row && typeof row === "object" ? row : {}),
          ...payload,
          id: Number.isInteger(Number(row?.id)) ? Number(row.id) : nextLocalEntityId(),
          created_at: cleanNullableText(row?.created_at) || nowIso,
          updated_at: nowIso,
        };
      } else if (Number.isInteger(Number(row?.id))) {
        const { data, error } = await client.from("projects").update(payload).eq("id", Number(row.id)).select("*").single();
        if (error) throw new Error(error.message || "Impossible de mettre à jour le projet.");
        savedProject = data;
      } else {
        const { data, error } = await client.from("projects").insert({
          ...payload,
          created_at: new Date().toISOString(),
        }).select("*").single();
        if (error) throw new Error(error.message || "Impossible de créer le projet.");
        savedProject = data;
      }
      if (!savedProject) throw new Error("Projet introuvable après enregistrement.");
      for (const pending of pendingPhotos) {
        if (pending?.preview_url) revokePendingPhotoPreview(pending.preview_url);
      }
      pendingPhotos = [];
      applySavedProjectState(savedProject, { successMessage: "Projet enregistré.", nextEditing: false, previousRow });
      saveSucceeded = true;
      return savedProject;
    } catch (err) {
      if (messageEl) messageEl.textContent = err?.message || "Impossible d'enregistrer le projet.";
      return null;
    } finally {
      saving = false;
      renderAll();
      if (saveSucceeded) initialSnapshot = readDraftSnapshot();
      if (hasPersistedProject) void hydrateStoredProjectPhotoPreviews();
    }
  };

  if (form) {
    form.addEventListener("submit", async (ev) => {
      ev.preventDefault();
      await saveProjectChanges();
    });
  }

  if (messageEl) messageEl.textContent = "";
  renderAll();
  void hydrateStoredProjectPhotoPreviews();
  initialSnapshot = readDraftSnapshot();

  modal.__projectCleanup = () => {
    for (const photo of pendingPhotos) {
      if (!photo?.preview_url || !String(photo.preview_url).startsWith("blob:")) continue;
      revokePendingPhotoPreview(photo.preview_url);
    }
    pendingPhotoPreviewUrls.clear();
  };
  modal.__projectIsDirty = () => readDraftSnapshot() !== initialSnapshot;
  modal.__projectIsSaving = () => saving;
  modal.__projectSave = saveProjectChanges;
  modal.hidden = false;
  projectModalState.open = true;
  document.body.style.overflow = "hidden";
}

function openCreateFactureModal(options = {}) {
  const modal = $("factureModal");
  const body = $("factureModalBody");
  const title = $("factureModalTitle");
  if (!modal || !body || !title) return;

  let sortedClients = [];
  const clientById = {};
  const rebuildFactureClientChoices = () => {
    sortedClients = [...(state.clients || [])]
      .filter((client) => Number.isInteger(getClientReferenceId(client)))
      .sort((a, b) => normalizeText(formatClientPrimaryName(a)).localeCompare(normalizeText(formatClientPrimaryName(b)), "fr-CA"));
    Object.keys(clientById).forEach((key) => delete clientById[key]);
    for (const client of sortedClients) {
      clientById[String(getClientReferenceId(client))] = client;
    }
  };
  rebuildFactureClientChoices();

  const mapFactureRepairRow = (repair) => {
    const linkedClient = resolveClientForRepair(repair);
    const clientItemId = Number(repair?.client_item_id ?? getClientReferenceId(linkedClient));
    const manufacturer = cleanNullableText(repair?.manufacturer) || "";
    const model = cleanNullableText(repair?.model) || "";
    const device = [manufacturer, model].filter(Boolean).join(" ").trim();
    return {
      ...repair,
      podioItemId: Number(repair.podio_item_id),
      repCode: formatRepairCode(repair),
      clientItemId: Number.isInteger(clientItemId) ? clientItemId : null,
      clientDisplay: linkedClient ? formatClientPrimaryName(linkedClient) : (cleanNullableText(repair?.client_title) || "Sans client"),
      clientCompany: cleanNullableText(linkedClient?.company),
      deviceLabel: device || "Appareil",
    };
  };
  let repairsAll = [...(state.repairs || [])]
    .filter((repair) => Number.isInteger(Number(repair?.podio_item_id)))
    .filter((repair) => !isPersonalRepair(repair))
    .map((repair) => mapFactureRepairRow(repair))
    .sort((a, b) => {
      const ta = Date.parse(a?.updated_at || a?.created_at || "");
      const tb = Date.parse(b?.updated_at || b?.created_at || "");
      const sa = Number.isFinite(ta) ? ta : 0;
      const sb = Number.isFinite(tb) ? tb : 0;
      return sb - sa;
    });

  const projectsAll = [...(state.projects || [])]
    .filter((project) => Number.isInteger(Number(project?.id)))
    .filter((project) => !isPersonalProject(project))
    .map((project) => {
      const linkedClient = state.clientsByPodioItemId?.[String(project?.client_item_id ?? "")] || null;
      const clientItemId = Number(project?.client_item_id ?? getClientReferenceId(linkedClient));
      return {
        ...project,
        projectId: Number(project.id),
        projectCode: (cleanNullableText(project?.numero) || "").toUpperCase(),
        projectTitle: cleanNullableText(project?.title) || "Projet",
        projectDescription: cleanNullableText(project?.description) || "",
        clientItemId: Number.isInteger(clientItemId) ? clientItemId : null,
        clientDisplay: linkedClient ? formatClientPrimaryName(linkedClient) : (cleanNullableText(project?.client_title) || "Sans client"),
        clientCompany: cleanNullableText(linkedClient?.company),
        costLines: normalizeProjectCostLines(project?.cost_lines),
        billableHours: Number.isFinite(Number(project?.billable_hours))
          ? Math.max(0, Number(project.billable_hours))
          : 0,
      };
    })
    .sort((a, b) => {
      const ta = Date.parse(a?.updated_at || a?.created_at || "");
      const tb = Date.parse(b?.updated_at || b?.created_at || "");
      const sa = Number.isFinite(ta) ? ta : 0;
      const sb = Number.isFinite(tb) ? tb : 0;
      return sb - sa;
    });

  const requestedPrefillRepairItemId = (() => {
    const n = Number(options?.prefillRepairItemId);
    return Number.isInteger(n) && n > 0 ? String(n) : "";
  })();
  const requestedPrefillProjectId = (() => {
    const n = Number(options?.prefillProjectId);
    return Number.isInteger(n) && n > 0 ? String(n) : "";
  })();
  const requestedContext = (() => {
    const text = normalizeText(options?.prefillContext);
    if (text === "projet" || text === "project") return "projet";
    if (text === "reparation" || text === "repair") return "reparation";
    return "";
  })();
  const requestedFactureId = String(options?.factureId ?? "").trim();
  const editingFacture = requestedFactureId
    ? ((state.factures || []).find((facture) => String(facture?.id) === requestedFactureId) || null)
    : null;
  const inferFactureContext = (facture) => {
    const directContext = normalizeText(facture?.contexte);
    if (directContext === "projet" || directContext === "project") return "projet";
    const refs = Array.isArray(facture?.reparations) ? facture.reparations : [];
    return refs.some((ref) => normalizeText(ref?.kind) === "project") ? "projet" : "reparation";
  };
  const editingFactureNumero = cleanNullableText(editingFacture?.numero) || cleanNullableText(editingFacture?.raw_title) || "";
  let factureContext = requestedContext || inferFactureContext(editingFacture) || (requestedPrefillProjectId ? "projet" : "reparation");
  const normalizeFactureStatusChoice = (value) => {
    const text = cleanNullableText(value) || "";
    const normalized = normalizeText(text);
    if (normalized === "payee" || normalized === "paye" || normalized === "paid") return "Payée";
    if (normalized === "en retard" || normalized === "retard") return "En retard";
    if (normalized === "emise" || normalized === "emisse" || normalized === "issued") return "Émise";
    return text || "Émise";
  };
  const buildFactureStatusOptionsHtml = (selectedValue) => {
    const currentValue = normalizeFactureStatusChoice(selectedValue);
    const baseOptions = ["Émise", "Payée", "En retard"];
    if (!baseOptions.includes(currentValue)) baseOptions.push(currentValue);
    return baseOptions.map((option) => {
      const selected = option === currentValue ? "selected" : "";
      return `<option value="${escapeHtml(option)}" ${selected}>${escapeHtml(option)}</option>`;
    }).join("");
  };

  let selectedClientId = "";
  let selectedClientTitle = "";
  let selectedClientCompany = "";
  let statusLabel = normalizeFactureStatusChoice(editingFacture?.etat_label);
  let clientPickerOpen = false;
  let selectedRepairs = new Map();
  let selectedProjects = new Map();
  let lineItems = [];
  let issueDate = formatDateYmd(new Date());
  let dueDate = addDaysToYmd(issueDate, FACTURE_DUE_DAYS);
  let laborRate = getConfiguredFactureLaborRate();
  const taxRates = getConfiguredFactureTaxRates();
  let applyTaxes = false;
  let problemDetails = "";
  let workDoneDetails = "";
  let lastAutoRepairProblemDetails = "";
  let lastAutoRepairWorkDoneDetails = "";
  let saving = false;
  let lineSeq = 0;
  const initialFactureContext = factureContext;

  const nextLineId = () => `ln-${Date.now()}-${lineSeq++}`;
  const round2 = (value) => {
    const n = Number(value);
    if (!Number.isFinite(n)) return 0;
    return Math.round(n * 100) / 100;
  };
  const parseNumberOrZero = (value) => {
    const raw = String(value ?? "").trim().replace(",", ".");
    const n = Number(raw);
    return Number.isFinite(n) ? n : 0;
  };
  const laborRateOptions = (() => {
    const cfg = state.appConfig || DEFAULT_APP_CONFIG;
    return [
      {
        key: "labor_rate_diagnostic",
        label: "Taux réduit",
        rate: normalizeNonNegativeNumber(cfg.labor_rate_diagnostic, DEFAULT_APP_CONFIG.labor_rate_diagnostic),
      },
      {
        key: "labor_rate_standard",
        label: "Taux standard",
        rate: normalizeNonNegativeNumber(cfg.labor_rate_standard, DEFAULT_APP_CONFIG.labor_rate_standard),
      },
      {
        key: "labor_rate_secondary",
        label: "Taux de service accéléré",
        rate: normalizeNonNegativeNumber(cfg.labor_rate_secondary, DEFAULT_APP_CONFIG.labor_rate_secondary),
      },
      {
        key: "labor_rate_estimation",
        label: "Taux d'estimation",
        rate: normalizeNonNegativeNumber(cfg.labor_rate_estimation, DEFAULT_APP_CONFIG.labor_rate_estimation),
      },
    ];
  })();
  const laborRateOptionByKey = Object.fromEntries(laborRateOptions.map((opt) => [opt.key, opt]));
  const defaultLaborRateKey = "labor_rate_standard";
  const getLaborRateValueByKey = (key) => {
    const option = laborRateOptionByKey[String(key ?? "").trim()];
    if (!option) return laborRate;
    return round2(option.rate);
  };
  const findLaborRateKeyByValue = (value) => {
    const target = round2(parseNumberOrZero(value));
    for (const option of laborRateOptions) {
      if (Math.abs(target - round2(option.rate)) < 0.005) return option.key;
    }
    return null;
  };
  const isLaborDescription = (value) => {
    const normalized = normalizeText(value)
      .replace(/[’'`]/g, "")
      .replace(/\s+/g, " ")
      .trim();
    return normalized.includes("main doeuvre") || normalized.includes("main d oeuvre");
  };
  const isLaborLine = (line) => {
    if (!line) return false;
    if (String(line.source_kind ?? "").trim() === "repair_labor") return true;
    return isLaborDescription(line.description);
  };
  const buildLaborRateOptionsHtml = (selectedKey) => {
    const currentKey = laborRateOptionByKey[selectedKey] ? selectedKey : defaultLaborRateKey;
    return laborRateOptions.map((option) => {
      const label = `${option.label} (${formatMoneyCompact(option.rate)}/h)`;
      const selected = option.key === currentKey ? "selected" : "";
      return `<option value="${escapeHtml(option.key)}" ${selected}>${escapeHtml(label)}</option>`;
    }).join("");
  };
  laborRate = getLaborRateValueByKey(defaultLaborRateKey);
  const lineAmount = (line) => {
    const qty = Math.max(0, parseNumberOrZero(line?.qty));
    const unit = Math.max(0, parseNumberOrZero(line?.unit_price));
    return round2(qty * unit);
  };
  const getTotals = () => {
    const subtotal = round2(lineItems.reduce((sum, line) => sum + lineAmount(line), 0));
    const tps = applyTaxes ? round2(subtotal * taxRates.tps) : 0;
    const tvq = applyTaxes ? round2(subtotal * taxRates.tvq) : 0;
    const total = round2(subtotal + tps + tvq);
    return { subtotal, tps, tvq, total };
  };
  const normalizeFactureModalDate = (value, fallback = formatDateYmd(new Date())) => {
    const text = String(value ?? "").trim();
    return /^\d{4}-\d{2}-\d{2}$/.test(text) ? text : fallback;
  };
  const readDraftSnapshot = () => JSON.stringify({
    context: factureContext,
    client_id: String(selectedClientId || ""),
    client_title: cleanNullableText(selectedClientTitle) || "",
    repairs: Array.from(selectedRepairs.values()).map((repair) => String(repair?.podioItemId || "")).sort(),
    projects: Array.from(selectedProjects.values()).map((project) => String(project?.projectId || "")).sort(),
    issue_date: cleanNullableText(issueDate) || "",
    due_date: cleanNullableText(dueDate) || "",
    status_label: normalizeFactureStatusChoice(statusLabel),
    taxes: applyTaxes === true,
    problem: cleanNullableText(problemDetails) || "",
    work_done: cleanNullableText(workDoneDetails) || "",
    lines: lineItems.map((line) => ({
      description: cleanNullableText(line?.description) || "",
      qty: round2(parseNumberOrZero(line?.qty)),
      unit_price: round2(parseNumberOrZero(line?.unit_price)),
      source_repair_item_id: cleanNullableText(line?.source_repair_item_id) || "",
      source_project_id: cleanNullableText(line?.source_project_id) || "",
    })),
  });
  let initialSnapshot = "";
  const hasDirtyData = () => readDraftSnapshot() !== initialSnapshot;

  const setClientFromId = (clientId) => {
    const key = String(clientId ?? "").trim();
    if (!key || !clientById[key]) {
      selectedClientId = "";
      selectedClientTitle = "";
      selectedClientCompany = "";
      return;
    }
    const row = clientById[key];
    selectedClientId = key;
    selectedClientTitle = formatClientPrimaryName(row);
    selectedClientCompany = cleanNullableText(row?.company) || "";
  };

  const clearAutoLinesForRepair = (repairItemId) => {
    const key = String(repairItemId ?? "").trim();
    lineItems = lineItems.filter((line) => String(line?.source_repair_item_id ?? "") !== key);
  };

  const clearAutoLinesForProject = (projectId) => {
    const key = String(projectId ?? "").trim();
    lineItems = lineItems.filter((line) => String(line?.source_project_id ?? "") !== key);
  };

  const clearSelectedRepairsAndAutoLines = () => {
    selectedRepairs.clear();
    lineItems = lineItems.filter((line) => !line?.source_repair_item_id);
  };

  const clearSelectedProjectsAndAutoLines = () => {
    selectedProjects.clear();
    lineItems = lineItems.filter((line) => !line?.source_project_id);
  };

  const addAutoLinesForRepair = (repairRow) => {
    const sourceId = String(repairRow.podioItemId);
    const repairLines = buildFactureLinesFromRepair(repairRow, { decorate: true });
    for (const line of repairLines) {
      lineItems.push({
        id: nextLineId(),
        description: cleanNullableText(line.description) || "",
        qty: round2(parseNumberOrZero(line.quantity)),
        unit_price: round2(parseNumberOrZero(line.unit_price)),
        labor_rate_key: isLaborLine(line)
          ? (cleanNullableText(line.labor_rate_key) || findLaborRateKeyByValue(line.unit_price) || defaultLaborRateKey)
          : null,
        source_kind: cleanNullableText(line.source_kind) || (isLaborLine(line) ? "repair_labor" : "repair_parts"),
        source_repair_item_id: sourceId,
      });
    }

  };
  const buildRepairNarrativeFromRows = (rows) => {
    const repairRows = Array.isArray(rows) ? rows.filter(Boolean) : [];
    const multiple = repairRows.length > 1;
    const problemParts = [];
    const workDoneParts = [];
    for (const repairRow of repairRows) {
      const repairProblem = cleanNullableText(htmlToText(repairRow?.desc_problem));
      const repairDone = cleanNullableText(htmlToText(repairRow?.desc_done));
      const header = [cleanNullableText(repairRow?.repCode), cleanNullableText(repairRow?.deviceLabel)]
        .filter(Boolean)
        .join(" | ");
      if (repairProblem) {
        problemParts.push(multiple && header ? `${header}\n${repairProblem}` : repairProblem);
      }
      if (repairDone) {
        workDoneParts.push(multiple && header ? `${header}\n${repairDone}` : repairDone);
      }
    }
    return {
      problemDetails: problemParts.join("\n\n").trim(),
      workDoneDetails: workDoneParts.join("\n\n").trim(),
    };
  };
  const buildRepairNarrativeSections = (rows) => {
    const repairRows = Array.isArray(rows) ? rows.filter(Boolean) : [];
    return repairRows
      .map((repairRow) => {
        const repCode = cleanNullableText(repairRow?.repCode) || "";
        const deviceLabel = cleanNullableText(repairRow?.deviceLabel) || "";
        const problem = cleanNullableText(htmlToText(repairRow?.desc_problem)) || "";
        const workDone = cleanNullableText(htmlToText(repairRow?.desc_done)) || "";
        if (!repCode && !deviceLabel && !problem && !workDone) return null;
        return {
          title: [repCode, deviceLabel].filter(Boolean).join(" | ").trim(),
          problem,
          work_done: workDone,
        };
      })
      .filter(Boolean);
  };
  const syncRepairNarrativeInputsFromSelection = ({ force = false } = {}) => {
    const rows = Array.from(selectedRepairs.values());
    const nextNarrative = buildRepairNarrativeFromRows(rows);
    const problemInput = $("factureProblemDetailsInput");
    const workInput = $("factureWorkDetailsInput");
    const currentProblemValue = cleanNullableText(problemInput?.value ?? problemDetails) || "";
    const currentWorkValue = cleanNullableText(workInput?.value ?? workDoneDetails) || "";
    const canReplaceProblem = force || !currentProblemValue || currentProblemValue === lastAutoRepairProblemDetails;
    const canReplaceWork = force || !currentWorkValue || currentWorkValue === lastAutoRepairWorkDoneDetails;

    lastAutoRepairProblemDetails = nextNarrative.problemDetails;
    lastAutoRepairWorkDoneDetails = nextNarrative.workDoneDetails;

    if (canReplaceProblem) {
      problemDetails = nextNarrative.problemDetails;
      if (problemInput) problemInput.value = problemDetails;
    }
    if (canReplaceWork) {
      workDoneDetails = nextNarrative.workDoneDetails;
      if (workInput) workInput.value = workDoneDetails;
    }
  };

  const addAutoLinesForProject = (projectRow) => {
    const sourceId = String(projectRow.projectId);
  const sourceLines = ensureProjectLaborCostLines(
      projectRow?.costLines,
      projectRow?.billableHours,
      getLaborRateValueByKey(defaultLaborRateKey),
      { includeLabor: !isPersonalProject(projectRow) },
    );
    if (sourceLines.length) {
      for (const line of sourceLines) {
        const quantity = Math.max(0, parseNumberOrZero(line?.quantity));
        const unitPrice = Math.max(0, getProjectCostLineCadUnitPrice(line));
        if (!cleanNullableText(line?.description) && quantity <= 0 && unitPrice <= 0) continue;
        lineItems.push({
          id: nextLineId(),
          description: cleanNullableText(line?.description) || "Ligne de projet",
          qty: quantity > 0 ? quantity : 1,
          unit_price: unitPrice,
          labor_rate_key: cleanNullableText(line?.labor_rate_key),
          source_kind: cleanNullableText(line?.source_kind) || "project_cost",
          source_project_id: sourceId,
        });
      }
    } else {
      lineItems.push({
        id: nextLineId(),
        description: projectRow.projectTitle || projectRow.projectCode || "Projet",
        qty: 1,
        unit_price: 0,
        labor_rate_key: null,
        source_kind: "project_cost",
        source_project_id: sourceId,
      });
    }

    const projectDescription = cleanNullableText(projectRow?.projectDescription);
    if (!cleanNullableText(problemDetails) && projectDescription) problemDetails = projectDescription;
    const workInputValue = $("factureWorkDetailsInput")?.value;
    const problemInputValue = $("factureProblemDetailsInput")?.value;
    if (!cleanNullableText(problemInputValue) && cleanNullableText(problemDetails)) {
      const problemEl = $("factureProblemDetailsInput");
      if (problemEl) problemEl.value = problemDetails;
    }
    if (!cleanNullableText(workDoneDetails) && !cleanNullableText(workInputValue)) {
      workDoneDetails = "";
    }
  };

  const addRepairSelection = (repairRow, messageEl) => {
    if (!repairRow) return;
    const repairKey = String(repairRow.podioItemId);
    if (selectedRepairs.has(repairKey)) {
      if (messageEl) messageEl.textContent = "Cette réparation est déjà sélectionnée.";
      return;
    }

    const repairClientId = repairRow.clientItemId != null ? String(repairRow.clientItemId) : "";
    if (repairClientId && selectedClientId && selectedClientId !== repairClientId) {
      const ok = window.confirm("Cette réparation appartient à un autre client. Changer le client et remplacer la sélection actuelle ?");
      if (!ok) return;
      clearSelectedRepairsAndAutoLines();
      setClientFromId(repairClientId);
    } else if (!selectedClientId && repairClientId) {
      setClientFromId(repairClientId);
    }

    selectedRepairs.set(repairKey, repairRow);
    addAutoLinesForRepair(repairRow);
    syncRepairNarrativeInputsFromSelection();
    if (messageEl) messageEl.textContent = "";
  };

  const addProjectSelection = (projectRow, messageEl) => {
    if (!projectRow) return;
    const projectKey = String(projectRow.projectId);
    if (selectedProjects.has(projectKey)) {
      if (messageEl) messageEl.textContent = "Ce projet est déjà sélectionné.";
      return;
    }

    const projectClientId = projectRow.clientItemId != null ? String(projectRow.clientItemId) : "";
    if (projectClientId && selectedClientId && selectedClientId !== projectClientId) {
      const ok = window.confirm("Ce projet appartient à un autre client. Changer le client et remplacer la sélection actuelle ?");
      if (!ok) return;
      clearSelectedProjectsAndAutoLines();
      setClientFromId(projectClientId);
    } else if (!selectedClientId && projectClientId) {
      setClientFromId(projectClientId);
    }

    clearSelectedProjectsAndAutoLines();
    selectedProjects.set(projectKey, projectRow);
    addAutoLinesForProject(projectRow);
    if (messageEl) messageEl.textContent = "";
  };

  const extractRepairCodeFromTitle = (value) => {
    const match = String(value ?? "").match(/REP\d{1,}/i);
    return match ? match[0].toUpperCase() : "";
  };
  const extractTitleSuffix = (value, fallback = "") => {
    const text = cleanNullableText(value);
    if (!text) return fallback;
    const parts = text.split("|").map((part) => String(part ?? "").trim()).filter(Boolean);
    if (parts.length >= 2) return parts.slice(1).join(" | ");
    return text;
  };

  const hydrateFromExistingFacture = () => {
    if (!editingFacture) return;

    const existingClientId = Number(editingFacture?.client_item_id);
    if (Number.isInteger(existingClientId) && clientById[String(existingClientId)]) {
      setClientFromId(String(existingClientId));
    } else {
      const fallbackClient = sortedClients.find((client) => normalizeText(formatClientPrimaryName(client)) === normalizeText(editingFacture?.client_title));
      if (getClientReferenceId(fallbackClient) != null) {
        setClientFromId(String(getClientReferenceId(fallbackClient)));
      } else {
        selectedClientId = Number.isInteger(existingClientId) ? String(existingClientId) : "";
        selectedClientTitle = cleanNullableText(editingFacture?.client_title) || "";
        selectedClientCompany = cleanNullableText(editingFacture?.client_snapshot?.company) || "";
      }
    }

    issueDate = normalizeFactureModalDate(editingFacture?.date_emission, formatDateYmd(new Date()));
    dueDate = normalizeFactureModalDate(editingFacture?.date_echeance, addDaysToYmd(issueDate, FACTURE_DUE_DAYS));
    applyTaxes = round2(editingFacture?.tps) > 0 || round2(editingFacture?.tvq) > 0;
    problemDetails = cleanNullableText(editingFacture?.desc_problem) || "";
    workDoneDetails = cleanNullableText(editingFacture?.desc_done) || "";

    const refs = Array.isArray(editingFacture?.reparations) ? editingFacture.reparations : [];
    const matchedRepairRows = [];
    const matchedProjectRows = [];
    if (factureContext === "projet") {
      for (const ref of refs) {
        const refKind = normalizeText(ref?.kind);
        if (refKind && refKind !== "project" && refKind !== "projet") continue;
        const projectId = Number(ref?.item_id);
        const projectCode = cleanNullableText(ref?.code)?.toUpperCase() || extractProjectCode(ref?.title);
        const matchedProject = projectsAll.find((project) => (
          (Number.isInteger(projectId) && project.projectId === projectId)
          || (projectCode && project.projectCode === projectCode)
        ));
        const projectRow = matchedProject || {
          projectId,
          projectCode: projectCode || `P${String(projectId || "").padStart(4, "0")}`,
          projectTitle: extractTitleSuffix(ref?.title, "Projet"),
          projectDescription: "",
          clientItemId: Number.isInteger(existingClientId) ? existingClientId : null,
          clientDisplay: cleanNullableText(editingFacture?.client_title) || "Sans client",
          clientCompany: cleanNullableText(editingFacture?.client_snapshot?.company) || "",
          costLines: [],
          billableHours: 0,
        };
        matchedProjectRows.push(projectRow);
        selectedProjects.set(String(projectRow.projectId), projectRow);
      }
    } else {
      for (const ref of refs) {
        const refKind = normalizeText(ref?.kind);
        if (refKind === "project" || refKind === "projet") continue;
        const repairItemId = Number(ref?.item_id);
        const repairCode = extractRepairCodeFromTitle(ref?.title);
        const matchedRepair = repairsAll.find((repair) => {
          const podioItemId = Number(repair?.podioItemId ?? repair?.podio_item_id);
          const localId = Number(repair?.id);
          const rowCode = cleanNullableText(repair?.repCode)?.toUpperCase() || "";
          if (Number.isInteger(repairItemId) && repairItemId > 0) {
            if (podioItemId === repairItemId || localId === repairItemId) return true;
          }
          return Boolean(repairCode && rowCode === repairCode);
        });
        const fallbackRepairItemId = Number(
          matchedRepair?.podioItemId
          ?? matchedRepair?.podio_item_id
          ?? repairItemId
        );
        const repairRow = matchedRepair || {
          podioItemId: Number.isInteger(fallbackRepairItemId) && fallbackRepairItemId > 0 ? fallbackRepairItemId : null,
          podio_item_id: Number.isInteger(fallbackRepairItemId) && fallbackRepairItemId > 0 ? fallbackRepairItemId : null,
          repCode: repairCode || `REP${String(repairItemId || "").slice(-4).padStart(4, "0")}`,
          clientItemId: Number.isInteger(existingClientId) ? existingClientId : null,
          clientDisplay: cleanNullableText(editingFacture?.client_title) || "Sans client",
          clientCompany: cleanNullableText(editingFacture?.client_snapshot?.company) || "",
          deviceLabel: extractTitleSuffix(ref?.title, "Appareil"),
          status_label: "",
          desc_problem: cleanNullableText(editingFacture?.desc_problem),
          desc_done: cleanNullableText(editingFacture?.desc_done),
        };
        matchedRepairRows.push(repairRow);
        selectedRepairs.set(String(repairRow.podioItemId || repairRow.podio_item_id || repairCode || repairItemId), repairRow);
      }
    }

    const storedLines = (Array.isArray(editingFacture?.lines) ? editingFacture.lines : [])
      .map((line) => {
        const description = cleanNullableText(line?.description) || "";
        const unitPrice = round2(parseNumberOrZero(line?.unit_price));
        return {
          id: nextLineId(),
          description,
          qty: Math.max(0, round2(parseNumberOrZero(line?.quantity))),
          unit_price: unitPrice,
          labor_rate_key: isLaborDescription(description) ? (findLaborRateKeyByValue(unitPrice) || defaultLaborRateKey) : null,
          source_kind: isLaborDescription(description) ? "repair_labor" : "custom",
          source_repair_item_id: null,
          source_project_id: null,
        };
      })
      .filter((line) => line.description || lineAmount(line) > 0);
    lineItems = storedLines;

    if (!lineItems.length) {
      if (factureContext === "projet") {
        for (const projectRow of matchedProjectRows) {
          addAutoLinesForProject(projectRow);
        }
      } else {
        for (const repairRow of matchedRepairRows) {
          addAutoLinesForRepair(repairRow);
        }
      }
    }
  };

  hydrateFromExistingFacture();

  body.innerHTML = `
    <form id="factureForm" class="facture-form">
      <section class="facture-section">
        <h4 class="facture-section-title">Contexte</h4>
        <div class="facture-grid-2">
          <div>
            <label class="facture-input-label" for="factureContextSelect">Type de facture</label>
            <select id="factureContextSelect" class="facture-input">
              <option value="reparation">Réparations</option>
              <option value="projet">Projet</option>
            </select>
          </div>
          <div>
            <label class="facture-input-label" for="factureStatusSelect">État</label>
            <select id="factureStatusSelect" class="facture-input">
              ${buildFactureStatusOptionsHtml(statusLabel)}
            </select>
          </div>
        </div>
      </section>

      <section class="facture-section">
        <h4 class="facture-section-title">Client</h4>
        <div class="facture-client-controls">
          <button id="factureClientAddBtn" type="button" class="repair-client-add-btn">Ajouter un client</button>
          <div id="factureClientBadge" class="facture-client-badge" hidden>
            <span id="factureClientBadgeName" class="facture-client-badge-name"></span>
            <span id="factureClientBadgeCompany" class="facture-client-badge-company" hidden></span>
            <button id="factureClientBadgeRemoveBtn" type="button" class="facture-client-badge-remove" aria-label="Retirer le client" title="Retirer le client" hidden>×</button>
          </div>
        </div>
        <div id="factureClientPicker" class="facture-picker" hidden>
          <input id="factureClientSearchInput" class="facture-input" type="search" placeholder="Rechercher un client">
          <div id="factureClientList" class="facture-picker-list"></div>
        </div>
      </section>

      <section id="factureRepairSection" class="facture-section">
        <h4 class="facture-section-title">Ajouter une réparation</h4>
        <div class="facture-grid-2">
          <div>
            <label class="facture-input-label" for="factureRepairSearchInput">Recherche globale (réparation)</label>
            <input id="factureRepairSearchInput" class="facture-input" type="search" placeholder="REP, client, marque, modèle...">
            <div id="factureRepairSearchList" class="facture-picker-list"></div>
          </div>
          <div>
            <label class="facture-input-label">Réparations du client sélectionné</label>
            <div id="factureClientRepairsList" class="facture-picker-list"></div>
          </div>
        </div>
        <div>
          <div id="factureSelectedRepairs" class="facture-repairs-selected"></div>
        </div>
      </section>

      <section id="factureProjectSection" class="facture-section" hidden>
        <h4 class="facture-section-title">Ajouter un projet</h4>
        <div class="facture-grid-2">
          <div>
            <label class="facture-input-label" for="factureProjectSearchInput">Recherche globale (projet)</label>
            <input id="factureProjectSearchInput" class="facture-input" type="search" placeholder="P0001, client, titre...">
            <div id="factureProjectSearchList" class="facture-picker-list"></div>
          </div>
          <div>
            <label class="facture-input-label">Projets du client sélectionné</label>
            <div id="factureClientProjectsList" class="facture-picker-list"></div>
          </div>
        </div>
        <div>
          <div id="factureSelectedProjects" class="facture-repairs-selected"></div>
        </div>
      </section>

      <section class="facture-section">
        <h4 class="facture-section-title">Lignes de facture</h4>
        <div class="facture-grid-2">
          <div>
            <label class="facture-input-label" for="factureIssueDateInput">Date d'émission</label>
            <input id="factureIssueDateInput" class="facture-input" type="date" value="${escapeHtml(issueDate)}">
          </div>
          <div>
            <label class="facture-input-label" for="factureDueDateInput">Date d'échéance</label>
            <input id="factureDueDateInput" class="facture-input" type="date" value="${escapeHtml(dueDate)}">
          </div>
          <div class="facture-tax-row">
            <input id="factureApplyTaxesInput" type="checkbox">
            <label for="factureApplyTaxesInput">Appliquer les taxes</label>
          </div>
        </div>
        <div class="facture-lines-table-wrap">
          <table class="facture-lines-table">
            <thead>
              <tr>
                <th>Description</th>
                <th>Qt</th>
                <th>Coût</th>
                <th>Montant</th>
                <th></th>
              </tr>
            </thead>
            <tbody id="factureLinesTbody"></tbody>
          </table>
        </div>
        <div class="facture-lines-actions">
          <button id="factureAddLineBtn" type="button" class="facture-add-line-btn">Ajouter une ligne</button>
        </div>
        <div class="facture-totals">
          <div class="facture-total-row"><span>Sous-total</span><span id="factureSubtotalValue">0,00 $</span></div>
          <div class="facture-total-row"><span>TPS</span><span id="factureTpsValue">0,00 $</span></div>
          <div class="facture-total-row"><span>TVQ</span><span id="factureTvqValue">0,00 $</span></div>
          <div class="facture-total-row"><strong>Total</strong><strong id="factureTotalValue">0,00 $</strong></div>
        </div>
      </section>

      <section class="facture-section">
        <h4 class="facture-section-title">Détails à afficher sur la facture</h4>
        <div class="facture-grid-2">
          <div>
            <label id="facturePrimaryDetailsLabel" class="facture-input-label" for="factureProblemDetailsInput">Description du problème</label>
            <textarea id="factureProblemDetailsInput" class="facture-textarea"></textarea>
          </div>
          <div id="factureWorkDetailsWrap">
            <label class="facture-input-label" for="factureWorkDetailsInput">Travail effectué</label>
            <textarea id="factureWorkDetailsInput" class="facture-textarea"></textarea>
          </div>
        </div>
      </section>

      <div class="facture-actions">
        <p id="factureMessage" class="facture-message"></p>
        <div class="facture-actions-right">
          ${editingFacture ? '<button id="factureRefreshRepairsBtn" type="button" class="repair-btn">Rafraîchir les bons</button>' : ""}
          ${editingFacture ? '<button id="factureSyncRepairBtn" type="button" class="repair-btn">Générer le PDF du bon</button>' : ""}
          <button id="factureCancelBtn" type="button" class="repair-btn">Annuler</button>
          <button id="factureSaveBtn" type="submit" class="repair-btn repair-btn-primary">Enregistrer</button>
        </div>
      </div>
    </form>
  `;

  const clientAddBtn = $("factureClientAddBtn");
  const contextSelect = $("factureContextSelect");
  const statusSelect = $("factureStatusSelect");
  const repairSection = $("factureRepairSection");
  const projectSection = $("factureProjectSection");
  const clientBadge = $("factureClientBadge");
  const clientBadgeName = $("factureClientBadgeName");
  const clientBadgeCompany = $("factureClientBadgeCompany");
  const clientBadgeRemoveBtn = $("factureClientBadgeRemoveBtn");
  const clientPicker = $("factureClientPicker");
  const clientSearchInput = $("factureClientSearchInput");
  const clientList = $("factureClientList");
  const repairSearchInput = $("factureRepairSearchInput");
  const repairSearchList = $("factureRepairSearchList");
  const clientRepairsList = $("factureClientRepairsList");
  const selectedRepairsEl = $("factureSelectedRepairs");
  const projectSearchInput = $("factureProjectSearchInput");
  const projectSearchList = $("factureProjectSearchList");
  const clientProjectsList = $("factureClientProjectsList");
  const selectedProjectsEl = $("factureSelectedProjects");
  const linesTbody = $("factureLinesTbody");
  const addLineBtn = $("factureAddLineBtn");
  const taxesInput = $("factureApplyTaxesInput");
  const issueDateInput = $("factureIssueDateInput");
  const dueDateInput = $("factureDueDateInput");
  const primaryDetailsLabel = $("facturePrimaryDetailsLabel");
  const problemInput = $("factureProblemDetailsInput");
  const workDetailsWrap = $("factureWorkDetailsWrap");
  const workInput = $("factureWorkDetailsInput");
  const subtotalValueEl = $("factureSubtotalValue");
  const tpsValueEl = $("factureTpsValue");
  const tvqValueEl = $("factureTvqValue");
  const totalValueEl = $("factureTotalValue");
  const form = $("factureForm");
  const saveBtn = $("factureSaveBtn");
  const refreshRepairsBtn = $("factureRefreshRepairsBtn");
  const syncRepairBtn = $("factureSyncRepairBtn");
  const cancelBtn = $("factureCancelBtn");
  const messageEl = $("factureMessage");

  const renderClientBadge = () => {
    const hasClient = Boolean((selectedClientId && clientById[selectedClientId]) || cleanNullableText(selectedClientTitle));
    if (clientAddBtn) {
      clientAddBtn.hidden = hasClient;
      clientAddBtn.disabled = saving || hasClient;
    }
    if (clientPicker) {
      const canOpenPicker = clientPickerOpen && !saving && !hasClient;
      clientPicker.hidden = !canOpenPicker;
    }
    if (clientSearchInput) clientSearchInput.disabled = saving || !clientPickerOpen || hasClient;
    if (clientBadgeRemoveBtn) {
      clientBadgeRemoveBtn.hidden = !hasClient;
      clientBadgeRemoveBtn.disabled = saving || !hasClient;
    }
    if (!clientBadge || !clientBadgeName || !clientBadgeCompany) return;
    clientBadge.hidden = !hasClient;
    if (!hasClient) {
      clientBadgeName.textContent = "";
      clientBadgeCompany.hidden = true;
      clientBadgeCompany.textContent = "";
      return;
    }
    clientBadgeName.textContent = selectedClientTitle || "Sans client";
    const company = cleanNullableText(selectedClientCompany);
    clientBadgeCompany.hidden = !company;
    clientBadgeCompany.textContent = company ? `(${company})` : "";
  };

  const renderClientPicker = (query = "") => {
    if (!clientList) return;
    const q = normalizeText(query);
    const filtered = sortedClients.filter((client) => {
      if (!q) return true;
      const hay = `${formatClientPrimaryName(client)} ${client.company || ""} ${client.email || ""} ${client.phone || ""}`;
      return normalizeText(hay).includes(q);
    });
    const createClientRow = `
      <button type="button" class="facture-picker-item facture-picker-item-create" data-client-action="create">
        <span class="facture-picker-item-name">+ Créer un nouveau client</span>
      </button>
    `;
    if (!filtered.length) {
      clientList.innerHTML = `${createClientRow}<p class="repair-client-picker-empty">Aucun client trouvé.</p>`;
      return;
    }
    clientList.innerHTML = `${createClientRow}${filtered.map((client) => {
      const id = String(getClientReferenceId(client) || "");
      const name = formatClientPrimaryName(client);
      const company = cleanNullableText(client.company);
      return `
        <button type="button" class="facture-picker-item" data-client-id="${escapeHtml(id)}">
          <span class="facture-picker-item-name">${escapeHtml(name)}</span>${company ? `<span class="facture-picker-item-sub"> (${escapeHtml(company)})</span>` : ""}
        </button>
      `;
    }).join("")}`;
  };

  const renderRepairSearchList = (query = "") => {
    if (!repairSearchList) return;
    const q = normalizeText(query);
    if (!q) {
      repairSearchList.innerHTML = "";
      repairSearchList.hidden = true;
      return;
    }

    repairSearchList.hidden = false;
    const filtered = repairsAll.filter((repair) => {
      const hay = `${repair.repCode} ${repair.deviceLabel} ${repair.clientDisplay} ${repair.clientCompany || ""} ${repair.status_label || ""}`;
      return normalizeText(hay).includes(q);
    }).slice(0, 30);

    if (!filtered.length) {
      repairSearchList.innerHTML = `<p class="repair-client-picker-empty">Aucune réparation trouvée.</p>`;
      return;
    }
    repairSearchList.innerHTML = filtered.map((repair) => `
      <button type="button" class="facture-picker-item" data-repair-item-id="${escapeHtml(repair.podioItemId)}">
        <span class="facture-picker-item-name">${escapeHtml(`${repair.repCode} | ${repair.deviceLabel}`)}</span>
        <span class="facture-picker-item-sub">${escapeHtml(`${repair.clientDisplay} • ${repair.status_label || "-"}`)}</span>
      </button>
    `).join("");
  };

  const renderClientRepairsList = () => {
    if (!clientRepairsList) return;
    if (!selectedClientId) {
      clientRepairsList.innerHTML = `<p class="repair-client-picker-empty">Sélectionnez d'abord un client.</p>`;
      return;
    }
    const rows = repairsAll.filter((repair) => String(repair.clientItemId || "") === String(selectedClientId)).slice(0, 80);
    if (!rows.length) {
      clientRepairsList.innerHTML = `<p class="repair-client-picker-empty">Aucune réparation associée à ce client.</p>`;
      return;
    }
    clientRepairsList.innerHTML = rows.map((repair) => {
      const selected = selectedRepairs.has(String(repair.podioItemId));
      return `
        <button type="button" class="facture-picker-item${selected ? " is-current" : ""}" data-repair-item-id="${escapeHtml(repair.podioItemId)}" ${selected || saving ? "disabled" : ""}>
          <span class="facture-picker-item-name">${escapeHtml(`${repair.repCode} | ${repair.deviceLabel}`)}</span>
          <span class="facture-picker-item-sub">${escapeHtml(repair.status_label || "-")}</span>
        </button>
      `;
    }).join("");
  };

  const renderSelectedRepairs = () => {
    if (!selectedRepairsEl) return;
    const rows = Array.from(selectedRepairs.values());
    if (!rows.length) {
      selectedRepairsEl.innerHTML = "";
      return;
    }
    selectedRepairsEl.innerHTML = rows.map((repair) => `
      <div class="facture-repair-chip">
        <span>${escapeHtml(`${repair.repCode} | ${repair.deviceLabel}`)}</span>
        <button type="button" class="facture-repair-chip-remove" data-remove-repair-item-id="${escapeHtml(repair.podioItemId)}" aria-label="Retirer">×</button>
      </div>
    `).join("");
  };

  const renderProjectSearchList = (query = "") => {
    if (!projectSearchList) return;
    const q = normalizeText(query);
    if (!q) {
      projectSearchList.innerHTML = "";
      projectSearchList.hidden = true;
      return;
    }
    projectSearchList.hidden = false;
    const filtered = projectsAll.filter((project) => {
      const hay = `${project.projectCode} ${project.projectTitle} ${project.clientDisplay} ${project.clientCompany || ""} ${project.projectDescription || ""}`;
      return normalizeText(hay).includes(q);
    }).slice(0, 30);
    if (!filtered.length) {
      projectSearchList.innerHTML = `<p class="repair-client-picker-empty">Aucun projet trouvé.</p>`;
      return;
    }
    projectSearchList.innerHTML = filtered.map((project) => `
      <button type="button" class="facture-picker-item" data-project-id="${escapeHtml(project.projectId)}">
        <span class="facture-picker-item-name">${escapeHtml(`${project.projectCode} | ${project.projectTitle}`)}</span>
        <span class="facture-picker-item-sub">${escapeHtml(project.clientDisplay || "-")}</span>
      </button>
    `).join("");
  };

  const renderClientProjectsList = () => {
    if (!clientProjectsList) return;
    if (!selectedClientId) {
      clientProjectsList.innerHTML = `<p class="repair-client-picker-empty">Sélectionnez d'abord un client.</p>`;
      return;
    }
    const rows = projectsAll.filter((project) => String(project.clientItemId || "") === String(selectedClientId)).slice(0, 80);
    if (!rows.length) {
      clientProjectsList.innerHTML = `<p class="repair-client-picker-empty">Aucun projet associé à ce client.</p>`;
      return;
    }
    const selectedProjectId = Array.from(selectedProjects.keys())[0] || "";
    clientProjectsList.innerHTML = rows.map((project) => {
      const selected = String(project.projectId) === selectedProjectId;
      return `
        <button type="button" class="facture-picker-item${selected ? " is-current" : ""}" data-project-id="${escapeHtml(project.projectId)}" ${selected || saving ? "disabled" : ""}>
          <span class="facture-picker-item-name">${escapeHtml(`${project.projectCode} | ${project.projectTitle}`)}</span>
          <span class="facture-picker-item-sub">${escapeHtml(project.projectDescription || "-")}</span>
        </button>
      `;
    }).join("");
  };

  const renderSelectedProjects = () => {
    if (!selectedProjectsEl) return;
    const rows = Array.from(selectedProjects.values());
    if (!rows.length) {
      selectedProjectsEl.innerHTML = "";
      return;
    }
    selectedProjectsEl.innerHTML = rows.map((project) => `
      <div class="facture-repair-chip">
        <span>${escapeHtml(`${project.projectCode} | ${project.projectTitle}`)}</span>
        <button type="button" class="facture-repair-chip-remove" data-remove-project-id="${escapeHtml(project.projectId)}" aria-label="Retirer">×</button>
      </div>
    `).join("");
  };

  const renderLines = () => {
    if (!linesTbody) return;
    refreshRepairMismatchFlags();
    if (!lineItems.length) {
      linesTbody.innerHTML = `
        <tr>
          <td colspan="5" class="data-empty">${factureContext === "projet" ? "Aucune ligne. Ajoutez un projet ou une ligne personnalisée." : "Aucune ligne. Ajoutez une réparation ou une ligne personnalisée."}</td>
        </tr>
      `;
      return;
    }
    linesTbody.innerHTML = lineItems.map((line) => {
      const amount = lineAmount(line);
      return `
        <tr data-line-id="${escapeHtml(line.id)}" class="${line?.repair_source_mismatch ? "is-repair-mismatch" : ""}">
          <td><input type="text" class="facture-input facture-line-desc-input" value="${escapeHtml(line.description || "")}" ${saving ? "disabled" : ""}></td>
          <td><input type="number" min="0" step="0.01" class="facture-input facture-line-qty-input" value="${escapeHtml(round2(parseNumberOrZero(line.qty)))}" ${saving ? "disabled" : ""}></td>
          <td><input type="number" min="0" step="0.01" class="facture-input facture-line-unit-input" value="${escapeHtml(round2(parseNumberOrZero(line.unit_price)))}" ${saving ? "disabled" : ""}></td>
          <td><span class="facture-line-total">${escapeHtml(formatMoney(amount))}</span></td>
          <td><button type="button" class="facture-line-remove" data-line-remove-id="${escapeHtml(line.id)}" ${saving ? "disabled" : ""}>×</button></td>
        </tr>
      `;
    }).join("");
  };

  const renderTotals = () => {
    const totals = getTotals();
    if (subtotalValueEl) subtotalValueEl.textContent = formatMoney(totals.subtotal);
    if (tpsValueEl) tpsValueEl.textContent = formatMoney(totals.tps);
    if (tvqValueEl) tvqValueEl.textContent = formatMoney(totals.tvq);
    if (totalValueEl) totalValueEl.textContent = formatMoney(totals.total);
  };
  const renderLineTotalCell = (row, line) => {
    if (!row || !line) return;
    const totalEl = row.querySelector(".facture-line-total");
    if (!totalEl) return;
    totalEl.textContent = formatMoney(lineAmount(line));
  };
  const buildFactureComparisonKey = (line) => {
    const description = normalizeText(cleanNullableText(line?.description) || "")
      .replace(/\s+/g, " ")
      .trim();
    const qty = round2(parseNumberOrZero(line?.qty ?? line?.quantity));
    const unitPrice = round2(parseNumberOrZero(line?.unit_price));
    return JSON.stringify([description, qty, unitPrice]);
  };
  const refreshRepairMismatchFlags = () => {
    const shouldCompare = Boolean(editingFacture && factureContext === "reparation" && selectedRepairs.size > 0);
    if (!shouldCompare) {
      lineItems.forEach((line) => {
        line.repair_source_mismatch = false;
      });
      return;
    }

    const expectedCounts = new Map();
    for (const repairRow of selectedRepairs.values()) {
      const currentRepairLines = buildFactureLinesFromRepair(repairRow, { decorate: true });
      for (const line of currentRepairLines) {
        const key = buildFactureComparisonKey(line);
        expectedCounts.set(key, (expectedCounts.get(key) || 0) + 1);
      }
    }

    lineItems.forEach((line) => {
      const hasContent = Boolean(
        cleanNullableText(line?.description)
        || round2(parseNumberOrZero(line?.qty)) > 0
        || round2(parseNumberOrZero(line?.unit_price)) > 0
      );
      if (!hasContent) {
        line.repair_source_mismatch = false;
        return;
      }
      const key = buildFactureComparisonKey(line);
      const remaining = expectedCounts.get(key) || 0;
      if (remaining > 0) {
        expectedCounts.set(key, remaining - 1);
        line.repair_source_mismatch = false;
      } else {
        line.repair_source_mismatch = true;
      }
    });

    const hasMissingExpectedLine = Array.from(expectedCounts.values()).some((count) => Number(count) > 0);
    if (hasMissingExpectedLine) {
      lineItems.forEach((line) => {
        const hasContent = Boolean(
          cleanNullableText(line?.description)
          || round2(parseNumberOrZero(line?.qty)) > 0
          || round2(parseNumberOrZero(line?.unit_price)) > 0
        );
        if (hasContent) line.repair_source_mismatch = true;
      });
    }
  };
  const buildSyncedRepairFactureDraft = () => {
    const nextLines = [];

    for (const repairRow of selectedRepairs.values()) {
      const sourceId = String(repairRow?.podioItemId ?? repairRow?.podio_item_id ?? "");
      const currentRepairLines = buildFactureLinesFromRepair(repairRow, { decorate: true });
      for (const line of currentRepairLines) {
        nextLines.push({
          id: nextLineId(),
          description: cleanNullableText(line?.description) || "",
          qty: round2(parseNumberOrZero(line?.quantity)),
          unit_price: round2(parseNumberOrZero(line?.unit_price)),
          labor_rate_key: isLaborLine(line)
            ? (cleanNullableText(line?.labor_rate_key) || findLaborRateKeyByValue(line?.unit_price) || defaultLaborRateKey)
            : null,
          source_kind: cleanNullableText(line?.source_kind) || (isLaborLine(line) ? "repair_labor" : "repair_parts"),
          source_repair_item_id: sourceId,
          source_project_id: null,
          repair_source_mismatch: false,
        });
      }
    }

    const nextNarrative = buildRepairNarrativeFromRows(Array.from(selectedRepairs.values()));

    return {
      lines: nextLines.filter((line) => line.description || lineAmount(line) > 0),
      problemDetails: nextNarrative.problemDetails,
      workDoneDetails: nextNarrative.workDoneDetails,
      narrativeSections: buildRepairNarrativeSections(Array.from(selectedRepairs.values())),
    };
  };
  const renderLineMismatchStates = () => {
    if (!linesTbody) return;
    for (const line of lineItems) {
      const row = Array.from(linesTbody.querySelectorAll("tr[data-line-id]"))
        .find((candidate) => String(candidate.getAttribute("data-line-id") ?? "") === String(line.id));
      if (!row) continue;
      const isMismatch = Boolean(line?.repair_source_mismatch);
      row.classList.toggle("is-repair-mismatch", isMismatch);
      if (isMismatch) {
        row.setAttribute("title", "Cette ligne est differente des informations actuelles du bon de reparation.");
      } else {
        row.removeAttribute("title");
      }
    }
  };

  const renderAll = () => {
    const isProjectContext = factureContext === "projet";
    if (contextSelect && contextSelect.value !== factureContext) contextSelect.value = factureContext;
    if (statusSelect) {
      const normalizedStatus = normalizeFactureStatusChoice(statusLabel);
      if (statusSelect.innerHTML !== buildFactureStatusOptionsHtml(normalizedStatus)) {
        statusSelect.innerHTML = buildFactureStatusOptionsHtml(normalizedStatus);
      }
      if (statusSelect.value !== normalizedStatus) statusSelect.value = normalizedStatus;
    }
    if (repairSection) repairSection.hidden = isProjectContext;
    if (projectSection) projectSection.hidden = !isProjectContext;
    if (primaryDetailsLabel) {
      primaryDetailsLabel.textContent = isProjectContext ? "Description du projet" : "Description du problème";
    }
    if (workDetailsWrap) workDetailsWrap.hidden = isProjectContext;
    if (isProjectContext) {
      workDoneDetails = "";
      if (workInput && workInput.value) workInput.value = "";
    }
    renderClientBadge();
    if (clientPickerOpen && clientSearchInput && !selectedClientId) {
      renderClientPicker(clientSearchInput.value);
    }
    if (isProjectContext) {
      renderClientProjectsList();
      renderSelectedProjects();
    } else {
      renderClientRepairsList();
      renderSelectedRepairs();
    }
    renderLines();
    renderLineMismatchStates();
    renderTotals();
    if (saveBtn) saveBtn.disabled = saving;
    if (refreshRepairsBtn) {
      refreshRepairsBtn.disabled = saving;
      refreshRepairsBtn.hidden = !(editingFacture && factureContext === "reparation" && selectedRepairs.size > 0);
    }
    if (syncRepairBtn) {
      syncRepairBtn.disabled = saving;
      syncRepairBtn.hidden = !(editingFacture && factureContext === "reparation" && selectedRepairs.size > 0);
    }
    if (cancelBtn) cancelBtn.disabled = saving;
    if (addLineBtn) addLineBtn.disabled = saving;
    if (clientBadgeRemoveBtn) clientBadgeRemoveBtn.disabled = saving;
    if (taxesInput) taxesInput.disabled = saving;
    if (issueDateInput) issueDateInput.disabled = saving;
    if (dueDateInput) dueDateInput.disabled = saving;
    if (problemInput) problemInput.disabled = saving;
    if (workInput) workInput.disabled = saving;
    if (repairSearchInput) repairSearchInput.disabled = saving || isProjectContext;
    if (projectSearchInput) projectSearchInput.disabled = saving || !isProjectContext;
    if (contextSelect) contextSelect.disabled = saving;
    if (statusSelect) statusSelect.disabled = saving;
  };

  const removeRepairSelection = (repairItemId) => {
    const key = String(repairItemId ?? "").trim();
    if (!key) return;
    selectedRepairs.delete(key);
    clearAutoLinesForRepair(key);
    syncRepairNarrativeInputsFromSelection();
  };

  if (clientAddBtn) {
    clientAddBtn.addEventListener("click", () => {
      const hasClient = Boolean(selectedClientId && clientById[selectedClientId]);
      if (!clientPicker || saving || hasClient) return;
      clientPickerOpen = !clientPickerOpen;
      if (clientPickerOpen) {
        renderClientPicker(clientSearchInput?.value ?? "");
        clientSearchInput?.focus();
      }
      renderAll();
    });
  }

  const openCreateClientFromFacturePicker = () => {
    if (saving) return;
    openCreateClientModal({
      onCreated: (createdClient) => {
        rebuildFactureClientChoices();
        const createdItemId = getClientReferenceId(createdClient);
        const nextClientId = Number.isInteger(createdItemId) ? String(createdItemId) : "";
        if (nextClientId) {
          setClientFromId(nextClientId);
        }
        clientPickerOpen = false;
        if (messageEl) messageEl.textContent = "";
        renderAll();
      }
    });
  };

  if (clientSearchInput) {
    clientSearchInput.addEventListener("input", () => {
      renderClientPicker(clientSearchInput.value);
    });
  }

  if (clientList) {
    clientList.addEventListener("click", (ev) => {
      if (saving) return;
      const actionBtn = ev.target?.closest?.("[data-client-action]");
      if (actionBtn && actionBtn.getAttribute("data-client-action") === "create") {
        openCreateClientFromFacturePicker();
        return;
      }
      const btn = ev.target?.closest?.("[data-client-id]");
      if (!btn) return;
      const nextClientId = String(btn.getAttribute("data-client-id") ?? "").trim();
      if (!nextClientId) return;

      const hasRepairSelection = selectedRepairs.size > 0;
      const hasProjectSelection = selectedProjects.size > 0;
      if (selectedClientId && selectedClientId !== nextClientId && (hasRepairSelection || hasProjectSelection)) {
        const warningText = hasProjectSelection
          ? "Changer de client va retirer le projet sélectionné. Continuer ?"
          : "Changer de client va retirer les réparations déjà sélectionnées. Continuer ?";
        const ok = window.confirm(warningText);
        if (!ok) return;
        clearSelectedRepairsAndAutoLines();
        clearSelectedProjectsAndAutoLines();
      }

      setClientFromId(nextClientId);
      clientPickerOpen = false;
      if (messageEl) messageEl.textContent = "";
      renderAll();
    });
  }

  if (clientBadgeRemoveBtn) {
    clientBadgeRemoveBtn.addEventListener("click", () => {
      if (saving) return;
      setClientFromId("");
      clientPickerOpen = false;
      clearSelectedRepairsAndAutoLines();
      clearSelectedProjectsAndAutoLines();
      renderAll();
      if (messageEl) messageEl.textContent = "";
    });
  }

  if (repairSearchInput) {
    repairSearchInput.addEventListener("input", () => {
      renderRepairSearchList(repairSearchInput.value);
    });
    renderRepairSearchList("");
  }

  if (projectSearchInput) {
    projectSearchInput.addEventListener("input", () => {
      renderProjectSearchList(projectSearchInput.value);
    });
    renderProjectSearchList("");
  }

  const handlePickRepair = (repairItemId, { fromSearch = false } = {}) => {
    const repair = repairsAll.find((row) => String(row.podioItemId) === String(repairItemId));
    if (!repair) return;
    addRepairSelection(repair, messageEl);
    renderAll();
    if (repairSearchInput) {
      if (fromSearch) {
        repairSearchInput.value = "";
        renderRepairSearchList("");
        repairSearchInput.blur();
      } else {
        renderRepairSearchList(repairSearchInput.value);
      }
    }
  };

  const handlePickProject = (projectId, { fromSearch = false } = {}) => {
    const project = projectsAll.find((row) => String(row.projectId) === String(projectId));
    if (!project) return;
    addProjectSelection(project, messageEl);
    renderAll();
    if (projectSearchInput) {
      if (fromSearch) {
        projectSearchInput.value = "";
        renderProjectSearchList("");
        projectSearchInput.blur();
      } else {
        renderProjectSearchList(projectSearchInput.value);
      }
    }
  };

  if (repairSearchList) {
    repairSearchList.addEventListener("click", (ev) => {
      if (saving) return;
      const btn = ev.target?.closest?.("[data-repair-item-id]");
      if (!btn) return;
      handlePickRepair(btn.getAttribute("data-repair-item-id"), { fromSearch: true });
    });
  }

  if (clientRepairsList) {
    clientRepairsList.addEventListener("click", (ev) => {
      if (saving) return;
      const btn = ev.target?.closest?.("[data-repair-item-id]");
      if (!btn) return;
      handlePickRepair(btn.getAttribute("data-repair-item-id"));
    });
  }

  if (projectSearchList) {
    projectSearchList.addEventListener("click", (ev) => {
      if (saving) return;
      const btn = ev.target?.closest?.("[data-project-id]");
      if (!btn) return;
      handlePickProject(btn.getAttribute("data-project-id"), { fromSearch: true });
    });
  }

  if (clientProjectsList) {
    clientProjectsList.addEventListener("click", (ev) => {
      if (saving) return;
      const btn = ev.target?.closest?.("[data-project-id]");
      if (!btn) return;
      handlePickProject(btn.getAttribute("data-project-id"));
    });
  }

  if (selectedRepairsEl) {
    selectedRepairsEl.addEventListener("click", (ev) => {
      if (saving) return;
      const btn = ev.target?.closest?.("[data-remove-repair-item-id]");
      if (!btn) return;
      removeRepairSelection(btn.getAttribute("data-remove-repair-item-id"));
      renderAll();
      if (repairSearchInput) renderRepairSearchList(repairSearchInput.value);
    });
  }

  if (selectedProjectsEl) {
    selectedProjectsEl.addEventListener("click", (ev) => {
      if (saving) return;
      const btn = ev.target?.closest?.("[data-remove-project-id]");
      if (!btn) return;
      const projectKey = String(btn.getAttribute("data-remove-project-id") || "");
      if (!projectKey) return;
      selectedProjects.delete(projectKey);
      clearAutoLinesForProject(projectKey);
      renderAll();
      if (projectSearchInput) renderProjectSearchList(projectSearchInput.value);
    });
  }

  if (contextSelect) {
    contextSelect.addEventListener("change", () => {
      if (saving) return;
      const nextContext = normalizeText(contextSelect.value) === "projet" ? "projet" : "reparation";
      if (nextContext === factureContext) return;
      if (nextContext === "projet" && selectedRepairs.size > 0) {
        const ok = window.confirm("Passer en mode Projet va retirer les réparations sélectionnées. Continuer ?");
        if (!ok) {
          contextSelect.value = factureContext;
          return;
        }
      }
      if (nextContext === "reparation" && selectedProjects.size > 0) {
        const ok = window.confirm("Passer en mode Réparations va retirer le projet sélectionné. Continuer ?");
        if (!ok) {
          contextSelect.value = factureContext;
          return;
        }
      }
      factureContext = nextContext;
      if (factureContext === "projet") {
        clearSelectedRepairsAndAutoLines();
      } else {
        clearSelectedProjectsAndAutoLines();
      }
      if (messageEl) messageEl.textContent = "";
      renderAll();
    });
  }

  if (statusSelect) {
    statusSelect.addEventListener("change", () => {
      statusLabel = normalizeFactureStatusChoice(statusSelect.value);
      if (messageEl) messageEl.textContent = "";
    });
  }

  if (addLineBtn) {
    addLineBtn.addEventListener("click", () => {
      if (saving) return;
      lineItems.push({
        id: nextLineId(),
        description: "",
        qty: 1,
        unit_price: 0,
        labor_rate_key: null,
        source_kind: "custom",
        source_repair_item_id: null,
        source_project_id: null,
      });
      renderAll();
    });
  }

  if (linesTbody) {
    linesTbody.addEventListener("input", (ev) => {
      if (saving) return;
      const row = ev.target?.closest?.("tr[data-line-id]");
      if (!row) return;
      const lineId = row.getAttribute("data-line-id");
      const line = lineItems.find((item) => String(item.id) === String(lineId));
      if (!line) return;

      const target = ev.target;
      if (!(target instanceof HTMLInputElement)) return;
      if (target.classList.contains("facture-line-desc-input")) {
        line.description = target.value;
      } else if (target.classList.contains("facture-line-qty-input")) {
        line.qty = Math.max(0, parseNumberOrZero(target.value));
      } else if (target.classList.contains("facture-line-unit-input")) {
        line.unit_price = Math.max(0, parseNumberOrZero(target.value));
      }
      refreshRepairMismatchFlags();
      renderLineTotalCell(row, line);
      renderTotals();
      renderLineMismatchStates();
    });

    linesTbody.addEventListener("click", (ev) => {
      if (saving) return;
      const btn = ev.target?.closest?.("[data-line-remove-id]");
      if (!btn) return;
      const lineId = String(btn.getAttribute("data-line-remove-id") ?? "");
      lineItems = lineItems.filter((line) => String(line.id) !== lineId);
      renderAll();
    });
  }

  if (taxesInput) {
    taxesInput.addEventListener("change", () => {
      applyTaxes = Boolean(taxesInput.checked);
      renderTotals();
    });
  }

  if (issueDateInput) {
    issueDateInput.addEventListener("change", () => {
      issueDate = String(issueDateInput.value ?? "").trim() || formatDateYmd(new Date());
      if (!dueDateInput?.value) {
        dueDate = addDaysToYmd(issueDate, FACTURE_DUE_DAYS);
        if (dueDateInput) dueDateInput.value = dueDate;
      }
    });
  }

  if (dueDateInput) {
    dueDateInput.addEventListener("change", () => {
      dueDate = String(dueDateInput.value ?? "").trim();
    });
  }

  if (problemInput) {
    problemInput.addEventListener("input", () => {
      problemDetails = problemInput.value;
    });
  }

  if (workInput) {
    workInput.addEventListener("input", () => {
      workDoneDetails = workInput.value;
    });
  }

  if (cancelBtn) {
    cancelBtn.addEventListener("click", async () => {
      await closeFactureModal();
    });
  }

  const refreshSelectedRepairsFromSource = async () => {
    if (!(editingFacture && factureContext === "reparation" && selectedRepairs.size > 0)) return false;
    const client = state.authClient || ensureAuthClient();
    if (!client) throw new Error("Connexion Supabase indisponible.");
    const repairIds = Array.from(selectedRepairs.values())
      .map((repair) => Number(repair?.podioItemId ?? repair?.podio_item_id))
      .filter((id) => Number.isInteger(id) && id !== 0);
    if (!repairIds.length) return false;

    const { data, error } = await client
      .from("repairs")
      .select("*")
      .in("podio_item_id", repairIds);
    if (error) throw new Error(error.message || "Impossible de rafraîchir les bons.");

    const refreshedRows = Array.isArray(data) ? data.map((repair) => mapFactureRepairRow(repair)) : [];
    if (!refreshedRows.length) return false;

    const refreshedByKey = new Map(refreshedRows.map((repair) => [String(repair.podioItemId), repair]));
    selectedRepairs = new Map(
      Array.from(selectedRepairs.entries()).map(([key, value]) => [key, refreshedByKey.get(String(value?.podioItemId ?? key)) || value]),
    );

    repairsAll = repairsAll.map((repair) => refreshedByKey.get(String(repair.podioItemId)) || repair);
    const knownRepairIds = new Set(repairsAll.map((repair) => String(repair.podioItemId)));
    for (const repair of refreshedRows) {
      const key = String(repair.podioItemId);
      if (!knownRepairIds.has(key)) repairsAll.push(repair);
    }
    state.repairs = (state.repairs || []).map((repair) => {
      const refreshed = refreshedByKey.get(String(repair?.podio_item_id ?? ""));
      return refreshed ? { ...repair, ...refreshed, podio_item_id: refreshed.podioItemId } : repair;
    });
    const stateRepairIds = new Set((state.repairs || []).map((repair) => String(repair?.podio_item_id ?? "")));
    for (const repair of refreshedRows) {
      if (!stateRepairIds.has(String(repair.podioItemId))) {
        state.repairs.push({ ...repair, podio_item_id: repair.podioItemId });
      }
    }

    const syncedDraft = buildSyncedRepairFactureDraft();
    if (syncedDraft.lines.length) {
      lineItems = syncedDraft.lines;
    }
    problemDetails = syncedDraft.problemDetails || "";
    workDoneDetails = syncedDraft.workDoneDetails || "";
    if (problemInput) problemInput.value = problemDetails;
    if (workInput) workInput.value = workDoneDetails;
    lastAutoRepairProblemDetails = problemDetails;
    lastAutoRepairWorkDoneDetails = workDoneDetails;
    refreshRepairMismatchFlags();
    renderAll();
    renderBonsBoard();
    return true;
  };

  if (refreshRepairsBtn) {
    refreshRepairsBtn.addEventListener("click", async () => {
      if (saving) return;
      try {
        saving = true;
        if (messageEl) messageEl.textContent = "Rafraîchissement des bons...";
        renderAll();
        const refreshed = await refreshSelectedRepairsFromSource();
        if (messageEl) messageEl.textContent = refreshed ? "Bons mis à jour." : "Aucune mise à jour trouvée.";
      } catch (err) {
        if (messageEl) messageEl.textContent = err?.message || "Impossible de rafraîchir les bons.";
      } finally {
        saving = false;
        renderAll();
      }
    });
  }

  const saveFactureChanges = async ({ syncRepairSource = false } = {}) => {
    if (saving) return;
    if (!selectedClientId) {
      if (messageEl) messageEl.textContent = "Sélectionnez un client avant d'enregistrer la facture.";
      return;
    }

    let validLines = lineItems
      .map((line) => ({
        ...line,
        description: String(line.description ?? "").trim(),
        qty: Math.max(0, parseNumberOrZero(line.qty)),
        unit_price: Math.max(0, parseNumberOrZero(line.unit_price)),
      }))
      .filter((line) => line.description || lineAmount(line) > 0);

    let repairNarrativeSections = null;
    if (syncRepairSource) {
      if (!(editingFacture && factureContext === "reparation" && selectedRepairs.size > 0)) {
        if (messageEl) messageEl.textContent = "Associez d'abord un bon de réparation avant de régénérer le PDF.";
        return;
      }
      const syncedDraft = buildSyncedRepairFactureDraft();
      if (syncedDraft.lines.length) {
        validLines = syncedDraft.lines;
        lineItems = syncedDraft.lines;
        problemDetails = syncedDraft.problemDetails || "";
        workDoneDetails = syncedDraft.workDoneDetails || "";
        repairNarrativeSections = Array.isArray(syncedDraft.narrativeSections) ? syncedDraft.narrativeSections : null;
        if (problemInput) problemInput.value = problemDetails;
        if (workInput) workInput.value = workDoneDetails;
      }
    }

    if (!validLines.length) {
      if (messageEl) messageEl.textContent = "Ajoutez au moins une ligne de facture.";
      return;
    }

    lineItems = validLines;
    const totals = getTotals();
    const repairsPayload = Array.from(selectedRepairs.values()).map((repair) => ({
      kind: "repair",
      item_id: Number(repair.podioItemId),
      title: `${repair.repCode} | ${repair.deviceLabel}`,
    }));
    const projectsPayload = Array.from(selectedProjects.values()).map((project) => ({
      kind: "project",
      item_id: Number(project.projectId),
      code: cleanNullableText(project.projectCode),
      title: `${project.projectCode} | ${project.projectTitle}`,
    }));
    const selectedClientRow = clientById[String(selectedClientId)] || null;
    const rawAddressLines = extractAddressLinesFromRaw(selectedClientRow?.raw?.address ?? null);
    const textAddressLines = extractAddressLinesFromText(selectedClientRow?.address ?? null);
    const clientStreet = rawAddressLines.street || textAddressLines.street || cleanNullableText(selectedClientRow?.address);
    const clientCityProvince = rawAddressLines.cityProvince || textAddressLines.cityProvince || null;
    const clientPostalCode = rawAddressLines.postal || textAddressLines.postal || null;
    const firstRepair = Array.from(selectedRepairs.values())[0] || null;
    const firstProject = Array.from(selectedProjects.values())[0] || null;
    const deviceLabel = factureContext === "projet"
      ? (cleanNullableText(firstProject?.projectTitle) || cleanNullableText(firstProject?.projectCode) || null)
      : (selectedRepairs.size > 1
        ? "Plusieurs réparations"
        : (cleanNullableText(firstRepair?.deviceLabel) || null));
    const clientSnapshot = selectedClientRow
      ? {
          name: formatClientPrimaryName(selectedClientRow),
          type: normalizeClientType(selectedClientRow?.type) || cleanNullableText(selectedClientRow?.type),
          company: cleanNullableText(selectedClientRow?.company),
          phone: cleanNullableText(selectedClientRow?.phone),
          email: cleanNullableText(selectedClientRow?.email),
          address_line1: cleanNullableText(clientStreet),
          city_province: cleanNullableText(clientCityProvince),
          postal_code: cleanNullableText(clientPostalCode),
        }
      : null;

    if (factureContext === "projet" && !projectsPayload.length) {
      if (messageEl) messageEl.textContent = "Sélectionnez un projet avant d'enregistrer la facture.";
      return;
    }

    const payload = {
      id: Number.isInteger(Number(editingFacture?.id)) ? Number(editingFacture.id) : undefined,
      podio_item_id: Number.isInteger(Number(editingFacture?.podio_item_id)) ? Number(editingFacture.podio_item_id) : undefined,
      podio_app_item_id: Number.isInteger(Number(editingFacture?.podio_app_item_id)) ? Number(editingFacture.podio_app_item_id) : undefined,
      numero: editingFactureNumero || undefined,
      client_item_id: Number(selectedClientId),
      client_title: selectedClientTitle || null,
      client_snapshot: clientSnapshot,
      date_emission: issueDate,
      date_echeance: dueDate || null,
      subtotal: totals.subtotal,
      tps: totals.tps,
      tvq: totals.tvq,
      total: totals.total,
      status_label: normalizeFactureStatusChoice(statusLabel),
      tax_year_label: String(new Date(`${issueDate}T00:00:00`).getFullYear()),
      reparations: factureContext === "projet" ? projectsPayload : repairsPayload,
      device_label: deviceLabel,
      lines: validLines.map((line) => ({
        description: line.description,
        quantity: round2(line.qty),
        unit_price: round2(line.unit_price),
        amount: lineAmount(line),
      })),
      desc_problem: cleanNullableText(problemDetails),
      desc_done: factureContext === "projet" ? null : cleanNullableText(workDoneDetails),
      repair_narrative_sections: syncRepairSource ? repairNarrativeSections : null,
      company_snapshot: getCompanySnapshotFromConfig(),
      contexte: factureContext,
    };

    try {
      saving = true;
      if (messageEl) {
        if (syncRepairSource) {
          messageEl.textContent = "Synchronisation du bon et régénération du PDF...";
        } else {
          messageEl.textContent = editingFacture ? "Mise à jour de la facture..." : "Création de la facture...";
        }
      }
      renderAll();
      await createFactureRecord(payload);
      renderFacturesList();
      renderBonsBoard();
      renderProjectsList();
      renderAccueilCards();
      finalizeCloseFactureModal();
    } catch (err) {
      if (messageEl) messageEl.textContent = err?.message || "Impossible d'enregistrer la facture.";
    } finally {
      saving = false;
      renderAll();
    }
  };

  if (form && saveBtn) {
    form.addEventListener("submit", async (ev) => {
      ev.preventDefault();
      await saveFactureChanges();
    });
  }

  if (syncRepairBtn) {
    syncRepairBtn.addEventListener("click", async () => {
      await saveFactureChanges({ syncRepairSource: true });
    });
  }

  if (messageEl) messageEl.textContent = "";
  if (problemInput) problemInput.value = problemDetails;
  if (workInput) workInput.value = workDoneDetails;
  if (taxesInput) taxesInput.checked = applyTaxes;
  if (contextSelect) contextSelect.value = factureContext;
  if (requestedPrefillRepairItemId) {
    const prefillRepair = repairsAll.find((row) => String(row.podioItemId) === requestedPrefillRepairItemId);
    if (prefillRepair) {
      factureContext = "reparation";
      addRepairSelection(prefillRepair, messageEl);
    }
  }
  if (requestedPrefillProjectId) {
    const prefillProject = projectsAll.find((row) => String(row.projectId) === requestedPrefillProjectId);
    if (prefillProject) {
      factureContext = "projet";
      addProjectSelection(prefillProject, messageEl);
    }
  }

  initialSnapshot = readDraftSnapshot();
  renderAll();

  title.textContent = editingFacture
    ? `Modifier la facture ${editingFactureNumero}`.trim()
    : "Nouvelle facture";
  modal.hidden = false;
  modal.__factureIsDirty = hasDirtyData;
  modal.__factureIsSaving = () => saving;
  modal.__factureFinalizeClose = finalizeCloseFactureModal;
  factureModalState.open = true;
  document.body.style.overflow = "hidden";
}

function openCreateDevisModal(typeDocument, options = {}) {
  const modal = $("factureModal");
  const body = $("factureModalBody");
  const title = $("factureModalTitle");
  if (!modal || !body || !title) return;

  const normalizedType = normalizeText(typeDocument) === "soumission" ? "soumission" : "estimation";
  const typeLabel = normalizedType === "soumission" ? "Soumission" : "Estimation";
  let row = options?.row && typeof options.row === "object" ? options.row : null;

  let sortedClients = [];
  const clientById = {};
  const rebuildDevisClientChoices = () => {
    sortedClients = [...(state.clients || [])]
      .filter((client) => Number.isInteger(getClientReferenceId(client)))
      .sort((a, b) => normalizeText(formatClientPrimaryName(a)).localeCompare(normalizeText(formatClientPrimaryName(b)), "fr-CA"));
    Object.keys(clientById).forEach((key) => delete clientById[key]);
    for (const client of sortedClients) {
      clientById[String(getClientReferenceId(client))] = client;
    }
  };
  rebuildDevisClientChoices();

  const normalizeDateInput = (value, fallback = formatDateYmd(new Date())) => {
    const raw = String(value ?? "").trim();
    if (/^\d{4}-\d{2}-\d{2}$/.test(raw)) return raw;
    if (/^\d{4}-\d{2}-\d{2}/.test(raw)) return raw.slice(0, 10);
    return fallback;
  };
  const normalizeStatusChoice = (value) => {
    const key = normalizeDevisStatusKey(value);
    if (key === "emise") return "Émise";
    if (key === "accepte") return "Accepté";
    if (key === "refuse") return "Refusé";
    if (key === "annule") return "Annulé";
    if (key === "echoue") return "Échoué";
    return "En rédaction";
  };

  let numero = cleanNullableText(row?.numero) || "";
  let emissionDate = normalizeDateInput(row?.date_emission);
  let validUntil = normalizeDateInput(row?.date_valid_until, addDaysToYmd(emissionDate, 30));
  let statusLabel = normalizeStatusChoice(row?.etat_label);
  let contexte = normalizeText(row?.contexte) === "reparation" ? "reparation" : "projet";
  let description = cleanNullableText(row?.description) || "";
  let notes = cleanNullableText(row?.notes) || "";
  let selectedClientId = "";
  let selectedClientTitle = cleanNullableText(row?.client_title) || "";
  let selectedClientCompany = "";
  const initialClientItemId = Number(row?.client_item_id);
  if (Number.isInteger(initialClientItemId) && initialClientItemId > 0) {
    const key = String(initialClientItemId);
    selectedClientId = key;
    const selectedClient = clientById[key] || null;
    if (selectedClient) {
      selectedClientTitle = formatClientPrimaryName(selectedClient);
      selectedClientCompany = cleanNullableText(selectedClient?.company) || "";
    }
  }

  let clientPickerOpen = false;
  let saving = false;
  let lineSeq = 0;
  let applyTaxes = Number(row?.tps || 0) > 0 || Number(row?.tvq || 0) > 0;
  const taxRates = getConfiguredFactureTaxRates();

  const nextLineId = () => `devis-ln-${Date.now()}-${lineSeq++}`;
  const round2 = (value) => {
    const n = Number(value);
    if (!Number.isFinite(n)) return 0;
    return Math.round(n * 100) / 100;
  };
  const parseNumberOrZero = (value) => {
    const raw = String(value ?? "").trim().replace(",", ".");
    const n = Number(raw);
    return Number.isFinite(n) ? n : 0;
  };
  let lineItems = Array.isArray(row?.lines)
    ? row.lines
      .map((line) => {
        if (!line || typeof line !== "object") return null;
        return {
          id: nextLineId(),
          description: cleanNullableText(line.description) || "",
          qty: Math.max(0, parseNumberOrZero(line.quantity ?? line.qty)),
          unit_price: Math.max(0, parseNumberOrZero(line.unit_price)),
        };
      })
      .filter(Boolean)
    : [];
  if (!lineItems.length) {
    const fallbackTotal = Math.max(0, parseNumberOrZero(row?.montant_sans_taxes ?? row?.total));
    if (fallbackTotal > 0) {
      lineItems = [{
        id: nextLineId(),
        description: normalizedType === "soumission" ? "Soumission" : "Estimation",
        qty: 1,
        unit_price: fallbackTotal,
      }];
    } else {
      lineItems = [{
        id: nextLineId(),
        description: "",
        qty: 1,
        unit_price: 0,
      }];
    }
  }

  const lineAmount = (line) => round2(
    Math.max(0, parseNumberOrZero(line?.qty)) * Math.max(0, parseNumberOrZero(line?.unit_price)),
  );
  const getTotals = () => {
    const subtotal = round2(lineItems.reduce((sum, line) => sum + lineAmount(line), 0));
    const tps = applyTaxes ? round2(subtotal * taxRates.tps) : 0;
    const tvq = applyTaxes ? round2(subtotal * taxRates.tvq) : 0;
    const total = round2(subtotal + tps + tvq);
    return { subtotal, tps, tvq, total };
  };
  const serializeLinesForDraft = () => lineItems.map((line) => ({
    description: String(line?.description ?? "").trim(),
    qty: round2(parseNumberOrZero(line?.qty)),
    unit_price: round2(parseNumberOrZero(line?.unit_price)),
  }));

  body.innerHTML = `
    <form id="devisForm" class="facture-form">
      <section class="facture-section">
        <h4 class="facture-section-title">Informations</h4>
        <div class="facture-grid-2">
          <div>
            <label class="facture-input-label" for="devisNumeroInput">Numero</label>
            <input id="devisNumeroInput" class="facture-input" type="text" value="${escapeHtml(numero)}" placeholder="G\u00e9n\u00e9r\u00e9 \u00e0 l'enregistrement" ${row ? "" : ""}>
          </div>
          <div>
            <label class="facture-input-label" for="devisIssueDateInput">Date d'emission</label>
            <input id="devisIssueDateInput" class="facture-input" type="date" value="${escapeHtml(emissionDate)}">
          </div>
          <div>
            <label class="facture-input-label" for="devisValidUntilInput">Valide jusqu'au</label>
            <input id="devisValidUntilInput" class="facture-input" type="date" value="${escapeHtml(validUntil)}">
          </div>
          <div>
            <label class="facture-input-label" for="devisStatusSelect">Statut</label>
            <select id="devisStatusSelect" class="facture-input">
              <option value="En rédaction">En rédaction</option>
              <option value="Émise">Émise</option>
              <option value="Accepté">Accepté</option>
              <option value="Refusé">Refusé</option>
              <option value="Annulé">Annulé</option>
              <option value="Échoué">Échoué</option>
            </select>
          </div>
          <div>
            <label class="facture-input-label" for="devisContextSelect">Type</label>
            <select id="devisContextSelect" class="facture-input">
              <option value="projet">Projet</option>
              <option value="reparation">Reparation</option>
            </select>
          </div>
          <div class="facture-tax-row">
            <input id="devisApplyTaxesInput" type="checkbox" ${applyTaxes ? "checked" : ""}>
            <label for="devisApplyTaxesInput">Appliquer les taxes</label>
          </div>
        </div>
      </section>

      <section class="facture-section">
        <h4 class="facture-section-title">Client</h4>
        <div class="facture-client-controls">
          <button id="devisClientAddBtn" type="button" class="repair-client-add-btn">Ajouter un client</button>
          <div id="devisClientBadge" class="facture-client-badge" hidden>
            <span id="devisClientBadgeName" class="facture-client-badge-name"></span>
            <span id="devisClientBadgeCompany" class="facture-client-badge-company" hidden></span>
            <button id="devisClientBadgeRemoveBtn" type="button" class="facture-client-badge-remove" aria-label="Retirer le client" title="Retirer le client" hidden>×</button>
          </div>
        </div>
        <div id="devisClientPicker" class="facture-picker" hidden>
          <input id="devisClientSearchInput" class="facture-input" type="search" placeholder="Rechercher un client">
          <div id="devisClientList" class="facture-picker-list"></div>
        </div>
      </section>

      <section class="facture-section">
        <h4 class="facture-section-title">Lignes</h4>
        <div class="facture-lines-table-wrap">
          <table class="facture-lines-table">
            <thead>
              <tr>
                <th>Description</th>
                <th>Qt</th>
                <th>Coût</th>
                <th>Montant</th>
                <th></th>
              </tr>
            </thead>
            <tbody id="devisLinesTbody"></tbody>
          </table>
        </div>
        <div class="facture-lines-actions">
          <button id="devisAddLineBtn" type="button" class="facture-add-line-btn">Ajouter une ligne</button>
        </div>
        <div class="facture-totals">
          <div class="facture-total-row"><span>Sous-total</span><span id="devisSubtotalValue">0,00 $</span></div>
          <div class="facture-total-row"><span>TPS</span><span id="devisTpsValue">0,00 $</span></div>
          <div class="facture-total-row"><span>TVQ</span><span id="devisTvqValue">0,00 $</span></div>
          <div class="facture-total-row"><strong>Total</strong><strong id="devisTotalValue">0,00 $</strong></div>
        </div>
      </section>

      <section class="facture-section">
        <h4 class="facture-section-title">Description</h4>
        <div class="facture-grid-2">
          <div>
            <label class="facture-input-label" for="devisDescriptionInput">Description</label>
            <textarea id="devisDescriptionInput" class="facture-textarea">${escapeHtml(description)}</textarea>
          </div>
          <div>
            <label class="facture-input-label" for="devisNotesInput">Notes internes</label>
            <textarea id="devisNotesInput" class="facture-textarea">${escapeHtml(notes)}</textarea>
          </div>
        </div>
      </section>
      <div class="facture-actions">
        <p id="devisMessage" class="facture-message"></p>
        <div class="facture-actions-right">
          <button id="devisPreviewBtn" type="button" class="repair-btn" ${Number.isInteger(Number(row?.id)) ? "" : "disabled"}>Apercu PDF</button>
          <button id="devisCancelBtn" type="button" class="repair-btn">Annuler</button>
          <button id="devisSaveBtn" type="submit" class="repair-btn repair-btn-primary">Enregistrer</button>
        </div>
      </div>
    </form>
  `;

  const numeroInput = $("devisNumeroInput");
  const clientAddBtn = $("devisClientAddBtn");
  const clientBadge = $("devisClientBadge");
  const clientBadgeName = $("devisClientBadgeName");
  const clientBadgeCompany = $("devisClientBadgeCompany");
  const clientBadgeRemoveBtn = $("devisClientBadgeRemoveBtn");
  const clientPicker = $("devisClientPicker");
  const clientSearchInput = $("devisClientSearchInput");
  const clientList = $("devisClientList");
  const issueInput = $("devisIssueDateInput");
  const validInput = $("devisValidUntilInput");
  const statusSelect = $("devisStatusSelect");
  const contextSelect = $("devisContextSelect");
  const taxesInput = $("devisApplyTaxesInput");
  const linesTbody = $("devisLinesTbody");
  const addLineBtn = $("devisAddLineBtn");
  const subtotalValueEl = $("devisSubtotalValue");
  const tpsValueEl = $("devisTpsValue");
  const tvqValueEl = $("devisTvqValue");
  const totalValueEl = $("devisTotalValue");
  const descriptionInput = $("devisDescriptionInput");
  const notesInput = $("devisNotesInput");
  const form = $("devisForm");
  const saveBtn = $("devisSaveBtn");
  const cancelBtn = $("devisCancelBtn");
  const previewBtn = $("devisPreviewBtn");
  const messageEl = $("devisMessage");

  if (statusSelect) statusSelect.value = statusLabel;
  if (contextSelect) contextSelect.value = contexte;

  const setClientFromId = (clientId) => {
    const key = String(clientId ?? "").trim();
    if (!key || !clientById[key]) {
      selectedClientId = "";
      selectedClientTitle = "";
      selectedClientCompany = "";
      return;
    }
    const selectedClient = clientById[key];
    selectedClientId = key;
    selectedClientTitle = formatClientPrimaryName(selectedClient);
    selectedClientCompany = cleanNullableText(selectedClient?.company) || "";
  };
  if (selectedClientId) setClientFromId(selectedClientId);

  const renderClientBadge = () => {
    const hasClient = Boolean(selectedClientId && clientById[selectedClientId]);
    if (clientAddBtn) {
      clientAddBtn.hidden = hasClient;
      clientAddBtn.disabled = saving || hasClient;
    }
    if (clientPicker) {
      clientPicker.hidden = !(clientPickerOpen && !saving && !hasClient);
    }
    if (clientSearchInput) clientSearchInput.disabled = saving || !clientPickerOpen || hasClient;
    if (clientBadgeRemoveBtn) {
      clientBadgeRemoveBtn.hidden = !hasClient;
      clientBadgeRemoveBtn.disabled = saving || !hasClient;
    }
    if (!clientBadge || !clientBadgeName || !clientBadgeCompany) return;
    clientBadge.hidden = !hasClient;
    if (!hasClient) {
      clientBadgeName.textContent = "";
      clientBadgeCompany.hidden = true;
      clientBadgeCompany.textContent = "";
      return;
    }
    clientBadgeName.textContent = selectedClientTitle || "Sans client";
    const company = cleanNullableText(selectedClientCompany);
    clientBadgeCompany.hidden = !company;
    clientBadgeCompany.textContent = company ? `(${company})` : "";
  };

  const renderClientPicker = (query = "") => {
    if (!clientList) return;
    const q = normalizeText(query);
    const filtered = sortedClients.filter((client) => {
      if (!q) return true;
      const hay = `${formatClientPrimaryName(client)} ${client.company || ""} ${client.email || ""} ${client.phone || ""}`;
      return normalizeText(hay).includes(q);
    });
    const createClientRow = `
      <button type="button" class="facture-picker-item facture-picker-item-create" data-client-action="create">
        <span class="facture-picker-item-name">+ Créer un nouveau client</span>
      </button>
    `;
    if (!filtered.length) {
      clientList.innerHTML = `${createClientRow}<p class="repair-client-picker-empty">Aucun client trouvé.</p>`;
      return;
    }
    clientList.innerHTML = `${createClientRow}${filtered.map((client) => {
      const id = String(getClientReferenceId(client) || "");
      const name = formatClientPrimaryName(client);
      const company = cleanNullableText(client.company);
      return `
        <button type="button" class="facture-picker-item" data-client-id="${escapeHtml(id)}">
          <span class="facture-picker-item-name">${escapeHtml(name)}</span>${company ? `<span class="facture-picker-item-sub"> (${escapeHtml(company)})</span>` : ""}
        </button>
      `;
    }).join("")}`;
  };

  const openCreateClientFromDevisPicker = () => {
    if (saving) return;
    openCreateClientModal({
      onCreated: (createdClient) => {
        rebuildDevisClientChoices();
        const createdItemId = getClientReferenceId(createdClient);
        const nextClientId = Number.isInteger(createdItemId) ? String(createdItemId) : "";
        if (nextClientId) setClientFromId(nextClientId);
        clientPickerOpen = false;
        if (messageEl) messageEl.textContent = "";
        renderAll();
      }
    });
  };

  const renderLines = () => {
    if (!linesTbody) return;
    if (!lineItems.length) {
      linesTbody.innerHTML = `
        <tr>
          <td colspan="5" class="data-empty">Aucune ligne. Ajoutez une ligne.</td>
        </tr>
      `;
      return;
    }
    linesTbody.innerHTML = lineItems.map((line) => `
      <tr data-line-id="${escapeHtml(line.id)}">
        <td><input type="text" class="facture-input devis-line-desc-input" value="${escapeHtml(line.description || "")}" ${saving ? "disabled" : ""}></td>
        <td><input type="number" min="0" step="0.01" class="facture-input devis-line-qty-input" value="${escapeHtml(round2(parseNumberOrZero(line.qty)))}" ${saving ? "disabled" : ""}></td>
        <td><input type="number" min="0" step="0.01" class="facture-input devis-line-unit-input" value="${escapeHtml(round2(parseNumberOrZero(line.unit_price)))}" ${saving ? "disabled" : ""}></td>
        <td><span class="facture-line-total">${escapeHtml(formatMoney(lineAmount(line)))}</span></td>
        <td><button type="button" class="facture-line-remove" data-line-remove-id="${escapeHtml(line.id)}" ${saving ? "disabled" : ""}>×</button></td>
      </tr>
    `).join("");
  };

  const renderTotals = () => {
    const totals = getTotals();
    if (subtotalValueEl) subtotalValueEl.textContent = formatMoney(totals.subtotal);
    if (tpsValueEl) tpsValueEl.textContent = formatMoney(totals.tps);
    if (tvqValueEl) tvqValueEl.textContent = formatMoney(totals.tvq);
    if (totalValueEl) totalValueEl.textContent = formatMoney(totals.total);
  };

  const readDraft = () => {
    return {
      numero: cleanNullableText(numeroInput?.value),
      client_item_id: selectedClientId ? Number(selectedClientId) : null,
      client_title: cleanNullableText(selectedClientTitle),
      date_emission: normalizeDateInput(issueInput?.value, formatDateYmd(new Date())),
      date_valid_until: normalizeDateInput(validInput?.value, addDaysToYmd(normalizeDateInput(issueInput?.value), 30)),
      etat_label: cleanNullableText(statusSelect?.value) || "En rédaction",
      contexte: cleanNullableText(contextSelect?.value) || "projet",
      apply_taxes: Boolean(taxesInput?.checked),
      lines: serializeLinesForDraft(),
      description: cleanNullableText(descriptionInput?.value),
      notes: cleanNullableText(notesInput?.value),
    };
  };

  const initialSerialized = JSON.stringify(readDraft());
  const hasDirtyData = () => JSON.stringify(readDraft()) !== initialSerialized;

  const renderUi = () => {
    const controls = [
      numeroInput,
      issueInput,
      validInput,
      statusSelect,
      contextSelect,
      taxesInput,
      descriptionInput,
      notesInput,
      saveBtn,
      cancelBtn,
      addLineBtn,
      clientAddBtn,
      clientBadgeRemoveBtn,
      clientSearchInput,
    ];
    controls.forEach((el) => {
      if (el) el.disabled = saving;
    });
    if (previewBtn) previewBtn.disabled = saving || !Number.isInteger(Number(row?.id));
  };

  if (previewBtn) {
    previewBtn.addEventListener("click", async () => {
      if (!row || !Number.isInteger(Number(row?.id))) {
        if (messageEl) messageEl.textContent = "Enregistrez d'abord le document pour générer le PDF.";
        return;
      }
      try {
        await openDevisPdfPreview(row);
      } catch (err) {
        if (messageEl) messageEl.textContent = err?.message || "Impossible d'ouvrir le PDF.";
      }
    });
  }

  if (cancelBtn) {
    cancelBtn.addEventListener("click", async () => {
      await closeFactureModal();
    });
  }

  if (issueInput && validInput) {
    issueInput.addEventListener("change", () => {
      const nextIssue = normalizeDateInput(issueInput.value, formatDateYmd(new Date()));
      if (!String(validInput.value ?? "").trim()) {
        validInput.value = addDaysToYmd(nextIssue, 30);
      }
    });
  }

  if (taxesInput) {
    taxesInput.addEventListener("change", () => {
      applyTaxes = Boolean(taxesInput.checked);
      renderTotals();
    });
  }

  if (clientAddBtn) {
    clientAddBtn.addEventListener("click", () => {
      const hasClient = Boolean(selectedClientId && clientById[selectedClientId]);
      if (!clientPicker || saving || hasClient) return;
      clientPickerOpen = !clientPickerOpen;
      if (clientPickerOpen) {
        renderClientPicker(clientSearchInput?.value ?? "");
        clientSearchInput?.focus();
      }
      renderAll();
    });
  }

  if (clientSearchInput) {
    clientSearchInput.addEventListener("input", () => {
      renderClientPicker(clientSearchInput.value);
    });
  }

  if (clientList) {
    clientList.addEventListener("click", (ev) => {
      if (saving) return;
      const actionBtn = ev.target?.closest?.("[data-client-action]");
      if (actionBtn && actionBtn.getAttribute("data-client-action") === "create") {
        openCreateClientFromDevisPicker();
        return;
      }
      const btn = ev.target?.closest?.("[data-client-id]");
      if (!btn) return;
      const nextClientId = String(btn.getAttribute("data-client-id") ?? "").trim();
      if (!nextClientId) return;
      setClientFromId(nextClientId);
      clientPickerOpen = false;
      if (messageEl) messageEl.textContent = "";
      renderAll();
    });
  }

  if (clientBadgeRemoveBtn) {
    clientBadgeRemoveBtn.addEventListener("click", () => {
      if (saving) return;
      setClientFromId("");
      clientPickerOpen = false;
      renderAll();
      if (messageEl) messageEl.textContent = "";
    });
  }

  if (addLineBtn) {
    addLineBtn.addEventListener("click", () => {
      if (saving) return;
      lineItems.push({
        id: nextLineId(),
        description: "",
        qty: 1,
        unit_price: 0,
      });
      renderAll();
    });
  }

  if (linesTbody) {
    linesTbody.addEventListener("input", (ev) => {
      if (saving) return;
      const rowEl = ev.target?.closest?.("tr[data-line-id]");
      if (!rowEl) return;
      const lineId = String(rowEl.getAttribute("data-line-id") ?? "");
      const line = lineItems.find((item) => String(item.id) === lineId);
      if (!line) return;
      const target = ev.target;
      if (!(target instanceof HTMLInputElement)) return;

      if (target.classList.contains("devis-line-desc-input")) {
        line.description = target.value;
      } else if (target.classList.contains("devis-line-qty-input")) {
        line.qty = Math.max(0, parseNumberOrZero(target.value));
      } else if (target.classList.contains("devis-line-unit-input")) {
        line.unit_price = Math.max(0, parseNumberOrZero(target.value));
      }

      const totalEl = rowEl.querySelector(".facture-line-total");
      if (totalEl) totalEl.textContent = formatMoney(lineAmount(line));
      renderTotals();
    });

    linesTbody.addEventListener("click", (ev) => {
      if (saving) return;
      const btn = ev.target?.closest?.("[data-line-remove-id]");
      if (!btn) return;
      const lineId = String(btn.getAttribute("data-line-remove-id") ?? "");
      lineItems = lineItems.filter((line) => String(line.id) !== lineId);
      renderAll();
    });
  }

  if (form && saveBtn) {
    form.addEventListener("submit", async (ev) => {
      ev.preventDefault();
      if (saving) return;
      const draft = readDraft();
      if (!draft.client_item_id) {
        if (messageEl) messageEl.textContent = "Sélectionnez un client avant d'enregistrer.";
        return;
      }
      if (!draft.date_emission) {
        if (messageEl) messageEl.textContent = "La date d'\u00e9mission est requise.";
        return;
      }
      if (!draft.date_valid_until) {
        if (messageEl) messageEl.textContent = "La date de validite est requise.";
        return;
      }
      const validLines = lineItems
        .map((line) => ({
          ...line,
          description: String(line.description ?? "").trim(),
          qty: Math.max(0, parseNumberOrZero(line.qty)),
          unit_price: Math.max(0, parseNumberOrZero(line.unit_price)),
        }))
        .filter((line) => line.description || lineAmount(line) > 0);
      if (!validLines.length) {
        if (messageEl) messageEl.textContent = "Ajoutez au moins une ligne.";
        return;
      }
      lineItems = validLines;
      const totals = getTotals();
      const selectedClientRow = clientById[String(selectedClientId)] || null;
      const rawAddressLines = extractAddressLinesFromRaw(selectedClientRow?.raw?.address ?? null);
      const textAddressLines = extractAddressLinesFromText(selectedClientRow?.address ?? null);
      const clientStreet = rawAddressLines.street || textAddressLines.street || cleanNullableText(selectedClientRow?.address);
      const clientCityProvince = rawAddressLines.cityProvince || textAddressLines.cityProvince || null;
      const clientPostalCode = rawAddressLines.postal || textAddressLines.postal || null;
      const clientSnapshot = selectedClientRow
        ? {
            name: formatClientPrimaryName(selectedClientRow),
            type: normalizeClientType(selectedClientRow?.type) || cleanNullableText(selectedClientRow?.type),
            company: cleanNullableText(selectedClientRow?.company),
            phone: cleanNullableText(selectedClientRow?.phone),
            email: cleanNullableText(selectedClientRow?.email),
            address_line1: cleanNullableText(clientStreet),
            city_province: cleanNullableText(clientCityProvince),
            postal_code: cleanNullableText(clientPostalCode),
          }
        : null;

      try {
        saving = true;
        if (messageEl) messageEl.textContent = `Enregistrement de la ${normalizedType}...`;
        renderUi();
        const payloadLines = validLines.map((line) => ({
          description: line.description,
          quantity: round2(line.qty),
          unit_price: round2(line.unit_price),
          amount: lineAmount(line),
        }));

        let updated = null;
        if (isDeveloperModeEnabled()) {
          const nowIso = new Date().toISOString();
          const generatedNumero = cleanNullableText(draft.numero)
            || cleanNullableText(row?.numero)
            || nextLocalDevisNumero(normalizedType, draft.date_emission);
          updated = {
            ...(row && typeof row === "object" ? row : {}),
            id: Number.isInteger(Number(row?.id)) ? Number(row.id) : nextLocalEntityId(),
            numero: generatedNumero.toUpperCase(),
            type_document: normalizedType,
            contexte: draft.contexte,
            client_item_id: Number(selectedClientId),
            client_title: selectedClientTitle || draft.client_title || null,
            client_snapshot: clientSnapshot,
            company_snapshot: getCompanySnapshotFromConfig(),
            date_emission: draft.date_emission,
            date_valid_until: draft.date_valid_until,
            etat_key: normalizeDevisStatusKey(draft.etat_label) || "en_redaction",
            etat_label: draft.etat_label,
            montant_sans_taxes: totals.subtotal,
            tps: totals.tps,
            tvq: totals.tvq,
            total: totals.total,
            lines: payloadLines,
            description: draft.description,
            notes: draft.notes,
            raw_title: cleanNullableText(draft.numero) || generatedNumero.toUpperCase(),
            updated_at: nowIso,
            created_at: cleanNullableText(row?.created_at) || nowIso,
          };
        } else {
          const client = state.authClient;
          if (!client) throw new Error("Connexion Supabase indisponible.");
          const { data, error } = await client.functions.invoke("estimations-submit", {
            body: {
              id: row?.id ?? undefined,
              type_document: normalizedType,
              numero: draft.numero,
              client_item_id: Number(selectedClientId),
              client_title: selectedClientTitle || draft.client_title,
              client_snapshot: clientSnapshot,
              company_snapshot: getCompanySnapshotFromConfig(),
              date_emission: draft.date_emission,
              date_valid_until: draft.date_valid_until,
              etat_label: draft.etat_label,
              contexte: draft.contexte,
              montant_sans_taxes: totals.subtotal,
              subtotal: totals.subtotal,
              tps: totals.tps,
              tvq: totals.tvq,
              total: totals.total,
              lines: payloadLines,
              description: draft.description,
              notes: draft.notes,
              raw_title: draft.numero,
            }
          });
          if (error) throw new Error(resolveEdgeFunctionInvokeError(error, "Erreur estimations-submit."));
          if (!data?.ok) {
            throw new Error(data?.detail || data?.error || "Impossible d'enregistrer ce document.");
          }
          updated = data?.item && typeof data.item === "object" ? data.item : null;
        }

        if (updated) {
          const hasExisting = (state.estimations || []).some((x) => String(x?.id) === String(updated.id));
          const nextRows = (state.estimations || []).map((x) => (String(x?.id) === String(updated.id) ? updated : x));
          state.estimations = hasExisting ? nextRows : [updated, ...nextRows];
          row = updated;
        }
        renderFacturesList();
        finalizeCloseFactureModal();
      } catch (err) {
        if (messageEl) messageEl.textContent = err?.message || "Impossible d'enregistrer ce document.";
      } finally {
        saving = false;
        renderUi();
      }
    });
  }

  const renderAll = () => {
    renderClientBadge();
    if (clientPickerOpen && clientSearchInput && !selectedClientId) {
      renderClientPicker(clientSearchInput.value);
    }
    renderLines();
    renderTotals();
    renderUi();
  };

  if (messageEl) messageEl.textContent = "";
  title.textContent = row
    ? `${typeLabel} ${cleanNullableText(row?.numero) || ""}`.trim()
    : `Nouvelle ${typeLabel.toLowerCase()}`;
  modal.hidden = false;
  modal.__factureIsDirty = hasDirtyData;
  modal.__factureIsSaving = () => saving;
  modal.__factureFinalizeClose = finalizeCloseFactureModal;
  factureModalState.open = true;
  document.body.style.overflow = "hidden";
  renderAll();
}

function updateThemeControls() {
  const switchBtn = $("themeSwitchBtn");
  const lightBtn = $("themeLightBtn");
  const darkBtn = $("themeDarkBtn");
  const isDark = state.theme === "dark";

  if (switchBtn) {
    switchBtn.classList.toggle("is-dark", isDark);
    switchBtn.setAttribute("aria-checked", String(isDark));
  }
  if (lightBtn) lightBtn.classList.toggle("is-active", !isDark);
  if (darkBtn) darkBtn.classList.toggle("is-active", isDark);
}

function renderDeveloperModeBanner() {
  const banner = $("developerModeBanner");
  const text = $("developerModeBannerText");
  if (!banner || !text) return;
  const active = state.appConfigLoaded && isDeveloperModeEnabled();
  banner.hidden = !active;
  if (active) {
    text.textContent = "Site en mode développeur, aucune modification ne sera enregistrée.";
  }
}

function applyTheme(theme) {
  state.theme = theme === "dark" ? "dark" : "light";
  document.documentElement.setAttribute("data-theme", state.theme);
  localStorage.setItem(THEME_KEY, state.theme);
  updateThemeControls();
  updateMenuCompanyLogo();
}

function initTheme() {
  const saved = localStorage.getItem(THEME_KEY);
  applyTheme(saved === "dark" ? "dark" : "light");
}

function toggleTheme() {
  applyTheme(state.theme === "dark" ? "light" : "dark");
}

function getViewFromHash() {
  const hash = (window.location.hash || "#accueil").replace("#", "").trim().toLowerCase();
  return VIEW_IDS.includes(hash) ? hash : "accueil";
}

function setActiveView(viewId) {
  const requested = VIEW_IDS.includes(viewId) ? viewId : "accueil";
  const target = requested === "projets" || requested === "bons-anciens" ? "bons" : requested;
  const activeMenuView = requested === "projets" || requested === "bons-anciens" ? "bons" : target;
  document.querySelectorAll(".view-panel").forEach((panel) => {
    panel.classList.toggle("is-active", panel.id === target);
  });
  document.querySelectorAll("[data-view-link]").forEach((link) => {
    link.classList.toggle("is-active", link.getAttribute("data-view-link") === activeMenuView);
  });

  if (target === "bons") {
    if (requested === "projets") {
      setBonsProjectsTab("projets");
    } else {
      setBonsProjectsTab("bons", { render: false });
      setBonsRepairsTab(requested === "bons-anciens" ? "archives" : "board");
    }
  }
  if (target === "clients") renderClientsList();
  if (target === "factures") renderFacturesList();
  if (target === "depenses") renderExpensesList();
  if (target === "configuration") renderConfigurationView();
}

async function loadDashboard() {
  if (!isAppPage()) return;
  const user = state.session?.user;
  const client = state.authClient;
  if (!user || !client) return;

  if (state.dashboardLoadedFor === user.id) {
    renderAccueilCards();
    renderBonsBoard();
    renderBonsOtherList();
    renderClientsList();
    renderProjectsList();
    renderFacturesList();
    renderExpensesList();
    renderConfigurationView();
    return;
  }

  state.dashboardLoadedFor = user.id;

  const accueilLoadingIds = [
    "accueilAnnivTitle",
    "accueilAnnivText",
    "accueilAnnivCounter",
    "accueilTotalRepairs",
    "accueilCompletedRepairs",
    "accueilAvgCycle",
    "accueilAvgCycleSub",
    "accueilAvgOperation",
    "accueilAvgOperationSub",
    "accueilAvgPartsCost",
    "accueilAvgPartsCostSub",
    "accueilBrandsMeta",
    "accueilYearMeta",
    "accueilStatusMeta",
  ];
  accueilLoadingIds.forEach((id) => {
    const el = $(id);
    if (el) el.textContent = "...";
  });
  if ($("accueilBrandLegend")) $("accueilBrandLegend").innerHTML = `<li class="data-empty">Chargement...</li>`;
  if ($("accueilYearTableBody")) $("accueilYearTableBody").innerHTML = `<tr><td colspan="5" class="data-empty">Chargement...</td></tr>`;
  if ($("accueilStatusTableBody")) $("accueilStatusTableBody").innerHTML = `<tr><td colspan="3" class="data-empty">Chargement...</td></tr>`;
  const accueilChart = $("accueilBrandChart");
  if (accueilChart) {
    accueilChart.classList.add("is-empty");
    accueilChart.style.background = "";
  }
  const accueilTrafficChart = $("accueilTrafficChart");
  if (accueilTrafficChart) {
    accueilTrafficChart.classList.add("is-empty");
    accueilTrafficChart.innerHTML = `<p class="data-empty">Chargement...</p>`;
  }
  if ($("bonsBoard")) $("bonsBoard").innerHTML = `<div class="data-empty">Chargement...</div>`;
  if ($("bonsOtherTableBody")) $("bonsOtherTableBody").innerHTML = `<tr><td colspan="5" class="data-empty">Chargement...</td></tr>`;
  if ($("clientsTableBody")) $("clientsTableBody").innerHTML = `<tr><td colspan="6" class="data-empty">Chargement...</td></tr>`;
  if ($("projectsList")) $("projectsList").innerHTML = `<div class="data-empty">Chargement...</div>`;
  if ($("facturesTableBody")) $("facturesTableBody").innerHTML = `<tr><td colspan="8" class="data-empty">Chargement...</td></tr>`;
  if ($("expensesTableBody")) $("expensesTableBody").innerHTML = `<tr><td colspan="5" class="data-empty">Chargement...</td></tr>`;
  if ($("bonsMeta")) $("bonsMeta").textContent = "...";
  if ($("bonsOtherMeta")) $("bonsOtherMeta").textContent = "";
  if ($("clientsMeta")) $("clientsMeta").textContent = "...";
  if ($("projectsMeta")) $("projectsMeta").textContent = "...";
  if ($("facturesMeta")) $("facturesMeta").textContent = "...";
  if ($("expensesMeta")) $("expensesMeta").textContent = "...";
  state.taxReportsLoading = true;
  state.taxReportsLoadError = null;

  const repairsQuery = client
    .from("repairs")
    .select("*")
    .order("updated_at", { ascending: false })
    .limit(250);

  const facturesQuery = client
    .from("factures")
    .select("id,podio_item_id,podio_app_item_id,numero,client_item_id,client_title,reparations,lines,company_snapshot,pdf_filename,raw_title,etat_label,montant_sans_taxes,tps,tvq,total,date_emission,date_echeance,desc_problem,desc_done,updated_at,created_at")
    .order("updated_at", { ascending: false });

  const estimationsQuery = client
    .from("estimations")
    .select("id,numero,type_document,contexte,client_item_id,client_title,client_snapshot,company_snapshot,raw_title,date_emission,date_valid_until,etat_key,etat_label,montant_sans_taxes,tps,tvq,total,lines,description,notes,pdf_filename,updated_at,created_at")
    .order("updated_at", { ascending: false });

  const clientsQuery = client
    .from("clients")
    .select("id,podio_item_id,podio_app_item_id,name,title,company,phone,email,address,type,raw,updated_at,created_at")
    .order("name", { ascending: true })
    .limit(1000);

  const projectsQuery = client
    .from("projects")
    .select("id,numero,client_item_id,client_title,is_personal,title,description,notes,completion_percent,current_stage_label,stages,files,photos,photo_thumbnail,cost_lines,billable_hours,updated_at,created_at")
    .order("updated_at", { ascending: false });

  const configQuery = client
    .from("app_config")
    .select("*")
    .eq("id", 1)
    .maybeSingle();

  const expensesQuery = client
    .from("expenses")
    .select("*")
    .order("purchase_date", { ascending: false })
    .order("id", { ascending: false });

  const taxReportsQuery = client
    .from("tax_reports")
    .select("*")
    .order("report_year", { ascending: false });

  let [repairsRes, facturesRes, estimationsRes, clientsRes, projectsRes, configRes, taxReportsRes, expensesRes] = await Promise.all([
    repairsQuery,
    facturesQuery,
    estimationsQuery,
    clientsQuery,
    projectsQuery,
    configQuery,
    taxReportsQuery,
    expensesQuery,
  ]);

  if (isMissingFacturesNumeroColumnError(facturesRes?.error)) {
    facturesRes = await client
      .from("factures")
      .select("id,podio_item_id,podio_app_item_id,reparations,lines,company_snapshot,pdf_filename,raw_title,client_title,etat_label,total,date_emission,date_echeance,desc_problem,desc_done,updated_at,created_at")
      .order("updated_at", { ascending: false });
  }

  state.repairs = repairsRes?.data || [];
  state.factures = facturesRes?.data || [];
  state.estimations = estimationsRes?.data || [];
  state.clients = clientsRes?.data || [];
  state.projects = projectsRes?.data || [];
  state.expenses = Array.isArray(expensesRes?.data)
    ? expensesRes.data.map((row) => normalizeExpenseRow(row)).filter(Boolean)
    : [];
  state.taxReports = Array.isArray(taxReportsRes?.data)
    ? taxReportsRes.data.map((row) => normalizeTaxReportRow(row)).filter(Boolean)
    : [];
  state.taxReportsLoading = false;
  state.clientsByPodioItemId = buildClientIndex(state.clients);
  state.factureByRepairItemId = buildFactureIndex(state.factures);
  state.estimationsLoadError = estimationsRes?.error
    ? `Estimations/Soumissions: ${estimationsRes.error.message || "Impossible de charger les devis."}`
    : null;
  if (configRes?.error) {
    state.appConfig = null;
    state.appConfigLoaded = false;
  } else {
    state.appConfig = configRes?.data ? normalizeAppConfigRow(configRes.data) : null;
    state.appConfigLoaded = true;
  }
  void updateMenuCompanyLogo();
  state.configEditing.honoraires = false;
  state.configEditing.entreprise = false;
  state.configEditing.achats = false;
  state.configEditing.developpement = false;
  state.configDeveloperDraft = null;
  state.configPodioSyncDraft = null;
  state.configPurchaseLinksDraft = null;
  renderDeveloperModeBanner();
  state.clientsLoadError = clientsRes?.error
    ? `Clients: ${clientsRes.error.message || "Impossible de charger les clients."}`
    : null;
  state.projectsLoadError = projectsRes?.error
    ? `Projets: ${projectsRes.error.message || "Impossible de charger les projets."}`
    : null;
  state.expensesLoadError = expensesRes?.error
    ? `Dépenses: ${expensesRes.error.message || "Impossible de charger les dépenses."}`
    : null;
  state.taxReportsLoadError = taxReportsRes?.error
    ? `Documents fiscaux: ${taxReportsRes.error.message || "Impossible de charger les rapports."}`
    : null;

  if (repairsRes?.error && facturesRes?.error && clientsRes?.error) {
    state.dashboardLoadedFor = null;
    const msg = "Impossible de charger le dashboard (permissions RLS ou tables manquantes).";
    if ($("accueilAnnivTitle")) $("accueilAnnivTitle").textContent = "Fondation: 4 mars 2024";
    if ($("accueilAnnivText")) $("accueilAnnivText").textContent = msg;
    if ($("accueilAnnivCounter")) $("accueilAnnivCounter").textContent = "";
    if ($("accueilTotalRepairs")) $("accueilTotalRepairs").textContent = "-";
    if ($("accueilCompletedRepairs")) $("accueilCompletedRepairs").textContent = msg;
    if ($("accueilAvgCycle")) $("accueilAvgCycle").textContent = "-";
    if ($("accueilAvgOperation")) $("accueilAvgOperation").textContent = "-";
    if ($("accueilAvgPartsCost")) $("accueilAvgPartsCost").textContent = "-";
    if ($("accueilBrandLegend")) $("accueilBrandLegend").innerHTML = `<li class="data-empty">${escapeHtml(msg)}</li>`;
    if ($("accueilTrafficChart")) $("accueilTrafficChart").innerHTML = `<p class="data-empty">${escapeHtml(msg)}</p>`;
    if ($("accueilYearTableBody")) $("accueilYearTableBody").innerHTML = `<tr><td colspan="5" class="data-empty">${escapeHtml(msg)}</td></tr>`;
    if ($("accueilStatusTableBody")) $("accueilStatusTableBody").innerHTML = `<tr><td colspan="3" class="data-empty">${escapeHtml(msg)}</td></tr>`;
    if ($("bonsBoard")) $("bonsBoard").innerHTML = `<div class="data-empty">${escapeHtml(msg)}</div>`;
    if ($("bonsOtherTableBody")) $("bonsOtherTableBody").innerHTML = `<tr><td colspan="5" class="data-empty">${escapeHtml(msg)}</td></tr>`;
    if ($("clientsTableBody")) $("clientsTableBody").innerHTML = `<tr><td colspan="6" class="data-empty">${escapeHtml(msg)}</td></tr>`;
    if ($("projectsList")) $("projectsList").innerHTML = `<div class="data-empty">${escapeHtml(msg)}</div>`;
    if ($("facturesTableBody")) $("facturesTableBody").innerHTML = `<tr><td colspan="8" class="data-empty">${escapeHtml(msg)}</td></tr>`;
    if ($("expensesTableBody")) $("expensesTableBody").innerHTML = `<tr><td colspan="5" class="data-empty">${escapeHtml(msg)}</td></tr>`;
    return;
  }

  renderAccueilCards();
  renderBonsBoard();
  renderBonsOtherList();
  renderClientsList();
  renderProjectsList();
  renderFacturesList();
  renderExpensesList();
  renderConfigurationView();
}

async function loadProfile() {
  const user = state.session?.user;
  const client = state.authClient;
  if (!user || !client) return;
  if (state.profileUserId === user.id) return;

  state.profileUserId = user.id;
  state.profile = null;
  updateMenuUser();
  void updateMenuCompanyLogo();

  try {
    const { data, error } = await client
      .from("profiles")
      .select("display_name, role")
      .eq("id", user.id)
      .maybeSingle();
    if (error) throw error;
    state.profile = data || null;
    updateMenuUser();
    void updateMenuCompanyLogo();
  } catch (_err) {
    // optional
  }
}

async function handlePasswordLogin(ev) {
  ev.preventDefault();
  const client = ensureAuthClient();
  if (!client) return;

  const email = $("loginEmail")?.value.trim();
  const password = $("loginPassword")?.value;
  const loginBtn = $("loginBtn");
  if (!email || !password) {
    setLoginStatus("Courriel et mot de passe obligatoires.", "warn");
    return;
  }

  try {
    if (loginBtn) loginBtn.disabled = true;
    setLoginStatus("Connexion en cours...", "warn");
    const { data, error } = await client.auth.signInWithPassword({ email, password });
    if (error) throw error;
    state.session = data.session ?? null;
    setLoginStatus("Connexion réussie.", "ok");
    goToApp();
  } catch (err) {
    setLoginStatus(err.message || "Erreur de connexion.", "bad");
  } finally {
    if (loginBtn) loginBtn.disabled = false;
  }
}

async function handleLogout() {
  const client = ensureAuthClient();
  if (!client) return;
  await client.auth.signOut();
  state.dashboardLoadedFor = null;
  state.profile = null;
  state.profileUserId = null;
  state.repairs = [];
  state.factures = [];
  state.expenses = [];
  state.estimations = [];
  state.clients = [];
  state.projects = [];
  state.taxReports = [];
  state.expensesLoadError = null;
  state.appConfig = null;
  state.appConfigLoaded = false;
  state.configEditing.honoraires = false;
  state.configEditing.entreprise = false;
  state.configEditing.achats = false;
  state.configEditing.developpement = false;
  state.configDeveloperDraft = null;
  state.configPodioSyncDraft = null;
  state.configPurchaseLinksDraft = null;
  renderDeveloperModeBanner();
  state.clientsByPodioItemId = {};
  state.factureByRepairItemId = {};
  state.estimationsLoadError = null;
  state.clientsLoadError = null;
  state.projectsLoadError = null;
  state.taxReportsLoadError = null;
  state.taxReportsLoading = false;
  state.taxReportsGeneratingYear = null;
  state.bonsProjectsActiveTab = "bons";
  state.bonsRepairsActiveTab = "board";
  state.projectsActiveTab = "board";
  state.facturationActiveTab = "factures";
  state.clientsSearchQuery = "";
  state.projectsSearchQuery = "";
  state.facturesSearchQuery = "";
  state.devisSearchQuery = "";
  state.facturesYearFilter = "all";
  state.facturesStatusFilter = "all";
  state.devisStatusFilter = "all";
  state.bonsOtherStatusFilter = "other";
  state.bonsOtherSearchQuery = "";
  state.bonsDrag.activeRepairId = null;
  state.bonsDrag.sourceColumnKey = null;
  state.bonsDrag.overColumnKey = null;
  state.bonsDrag.dropTargetKey = null;
  state.bonsDrag.isSaving = false;
  state.bonsDragIgnoreClickUntil = 0;
  clientModalState.clientId = null;
  clientModalState.original = null;
  clientModalState.draft = null;
  clientModalState.editing = false;
  clientModalState.saving = false;
  clientModalState.message = "";
  clientModalState.isCreating = false;
  clientModalState.onCreated = null;
  resetClientAddressSuggestState();
  projectThumbUrlCache.clear();
  projectThumbLoadingKeys.clear();
  repairThumbUrlCache.clear();
  repairThumbLoadingKeys.clear();
  state.session = null;
  goToLogin();
}

function bindLoginUi() {
  const form = $("loginForm");
  const pwd = $("loginPassword");
  const toggle = $("togglePasswordBtn");
  if (!form || !pwd || !toggle) return;

  let visible = false;
  setPasswordVisibility(visible);

  toggle.addEventListener("click", () => {
    visible = !visible;
    setPasswordVisibility(visible);
  });

  pwd.addEventListener("keydown", (ev) => {
    if (ev.key === "Enter") {
      ev.preventDefault();
      form.requestSubmit();
    }
  });

  form.addEventListener("submit", handlePasswordLogin);
}

function closeMenu() {
  const btn = $("burgerBtn");
  const menu = $("burgerMenu");
  if (!btn || !menu) return;
  menu.classList.remove("open");
  btn.setAttribute("aria-expanded", "false");
}

function openMenu() {
  const btn = $("burgerBtn");
  const menu = $("burgerMenu");
  if (!btn || !menu) return;
  menu.classList.add("open");
  btn.setAttribute("aria-expanded", "true");
}

function bindAppUi() {
  const btn = $("burgerBtn");
  const menu = $("burgerMenu");
  const logoutBtn = $("logoutBtn");
  const themeSwitchBtn = $("themeSwitchBtn");
  const themeLightBtn = $("themeLightBtn");
  const themeDarkBtn = $("themeDarkBtn");
  const bonsBoard = $("bonsBoard");
  const bonsOtherStatusFilter = $("bonsOtherStatusFilter");
  const bonsOtherSearchInput = $("bonsOtherSearchInput");
  const bonsOtherTableBody = $("bonsOtherTableBody");
  const bonsCreateBtn = $("bonsCreateBtn");
  const clientsTableBody = $("clientsTableBody");
  const clientsSearchInput = $("clientsSearchInput");
  const clientsRefreshBtn = $("clientsRefreshBtn");
  const clientsAddBtn = $("clientsAddBtn");
  const projectsList = $("projectsList");
  const projectsArchivedList = $("projectsArchivedList");
  const projectsCreateBtn = $("projectsCreateBtn");
  const facturesTableBody = $("facturesTableBody");
  const facturesCreateBtn = $("facturesCreateBtn");
  const expensesTableBody = $("expensesTableBody");
  const expensesCreateBtn = $("expensesCreateBtn");
  const facturesSearchInput = $("facturesSearchInput");
  const facturesYearFilterSelect = $("facturesYearFilterSelect");
  const facturesStatusFilterSelect = $("facturesStatusFilterSelect");
  const devisStatusFilterSelect = $("devisStatusFilterSelect");
  const bonsProjectsTabButtons = Array.from(document.querySelectorAll("[data-bons-projects-tab]"));
  const bonsRepairsTabButtons = Array.from(document.querySelectorAll("[data-bons-repairs-tab]"));
  const projectsTabButtons = Array.from(document.querySelectorAll("[data-projects-tab]"));
  const facturationTabButtons = Array.from(document.querySelectorAll("[data-facturation-tab]"));
  const configLaborRateStandardInput = $("configLaborRateStandardInput");
  const configLaborRateSecondaryInput = $("configLaborRateSecondaryInput");
  const configLaborRateDiagnosticInput = $("configLaborRateDiagnosticInput");
  const configLaborRateEstimationInput = $("configLaborRateEstimationInput");
  const configDeveloperModeOffBtn = $("configDeveloperModeOffBtn");
  const configDeveloperModeSwitchBtn = $("configDeveloperModeSwitchBtn");
  const configDeveloperModeOnBtn = $("configDeveloperModeOnBtn");
  const configCompanyNameInput = $("configCompanyNameInput");
  const configCompanyPhoneInput = $("configCompanyPhoneInput");
  const configCompanyEmailInput = $("configCompanyEmailInput");
  const configCompanyAddress1Input = $("configCompanyAddress1Input");
  const configCompanyAddress2Input = $("configCompanyAddress2Input");
  const configCompanyCityInput = $("configCompanyCityInput");
  const configCompanyProvinceInput = $("configCompanyProvinceInput");
  const configCompanyPostalCodeInput = $("configCompanyPostalCodeInput");
  const configCompanyCountryInput = $("configCompanyCountryInput");
  const configCompanyLogoPathInput = $("configCompanyLogoPathInput");
  const configPurchaseLinksList = $("configPurchaseLinksList");
  const configTaxReportsList = $("configTaxReportsList");
  const configTpsNumberInput = $("configTpsNumberInput");
  const configTpsRateInput = $("configTpsRateInput");
  const configTvqNumberInput = $("configTvqNumberInput");
  const configTvqRateInput = $("configTvqRateInput");
  const configHonorairesEditBtn = $("configHonorairesEditBtn");
  const configHonorairesCancelBtn = $("configHonorairesCancelBtn");
  const configHonorairesSaveBtn = $("configHonorairesSaveBtn");
  const configEntrepriseEditBtn = $("configEntrepriseEditBtn");
  const configEntrepriseCancelBtn = $("configEntrepriseCancelBtn");
  const configEntrepriseSaveBtn = $("configEntrepriseSaveBtn");
  const configAchatsEditBtn = $("configAchatsEditBtn");
  const configAchatsCancelBtn = $("configAchatsCancelBtn");
  const configAchatsSaveBtn = $("configAchatsSaveBtn");
  const configDeveloppementCancelBtn = $("configDeveloppementCancelBtn");
  const configDeveloppementSaveBtn = $("configDeveloppementSaveBtn");
  const configMessage = $("configMessage");
  const configTabButtons = Array.from(document.querySelectorAll("[data-config-tab]"));
  const repairModal = $("repairModal");
  const repairModalCloseBtn = $("repairModalCloseBtn");
  const clientModal = $("clientModal");
  const clientModalCloseBtn = $("clientModalCloseBtn");
  const clientModalEditBtn = $("clientModalEditBtn");
  const factureModal = $("factureModal");
  const factureModalCloseBtn = $("factureModalCloseBtn");
  const projectModal = $("projectModal");
  const projectModalCloseBtn = $("projectModalCloseBtn");
  const expenseModal = $("expenseModal");
  const expenseModalCloseBtn = $("expenseModalCloseBtn");
  const facturePdfModal = $("facturePdfModal");
  const facturePdfModalCloseBtn = $("facturePdfModalCloseBtn");
  const facturePdfPrevBtn = $("facturePdfPrevBtn");
  const facturePdfNextBtn = $("facturePdfNextBtn");
  const facturePdfZoomOutBtn = $("facturePdfZoomOutBtn");
  const facturePdfZoomInBtn = $("facturePdfZoomInBtn");
  if (!btn || !menu) return;

  // Evite l'effet de "flash" de valeurs restaurees par le navigateur
  // avant la fin du chargement de la configuration distante.
  renderConfigurationView();
  setBonsProjectsTab(state.bonsProjectsActiveTab, { render: false });
  setBonsRepairsTab(state.bonsRepairsActiveTab, { render: false });
  setProjectsTab(state.projectsActiveTab, { render: false });
  setFacturationTab(state.facturationActiveTab, { render: false });

  btn.addEventListener("click", () => {
    if (menu.classList.contains("open")) closeMenu();
    else openMenu();
  });

  document.addEventListener("click", (ev) => {
    if (menu.classList.contains("open") && !menu.contains(ev.target) && !btn.contains(ev.target)) closeMenu();
  });

  document.addEventListener("keydown", (ev) => {
    if (ev.key === "Escape") {
      closeMenu();
      closeRepairModal();
      closeClientModal();
      closeFactureModal();
      closeProjectModal();
      closeExpenseModal();
      closeFacturePdfModal();
    }
  });

  menu.querySelectorAll("a").forEach((link) => {
    link.addEventListener("click", () => closeMenu());
  });

  window.addEventListener("hashchange", () => {
    setActiveView(getViewFromHash());
  });

  if (themeSwitchBtn) themeSwitchBtn.addEventListener("click", () => toggleTheme());
  if (themeLightBtn) themeLightBtn.addEventListener("click", () => applyTheme("light"));
  if (themeDarkBtn) themeDarkBtn.addEventListener("click", () => applyTheme("dark"));
  if (logoutBtn) logoutBtn.addEventListener("click", handleLogout);
  if (bonsCreateBtn) bonsCreateBtn.addEventListener("click", () => openCreateRepairModal());
  if (bonsOtherStatusFilter) {
    bonsOtherStatusFilter.addEventListener("change", () => {
      state.bonsOtherStatusFilter = String(bonsOtherStatusFilter.value || "other");
      renderBonsOtherList();
    });
  }
  if (bonsOtherSearchInput) {
    bonsOtherSearchInput.addEventListener("input", () => {
      state.bonsOtherSearchQuery = bonsOtherSearchInput.value || "";
      renderBonsOtherList();
    });
  }
  if (bonsOtherTableBody) {
    bonsOtherTableBody.addEventListener("click", (ev) => {
      if (!(ev.target instanceof Element)) return;
      const trigger = ev.target.closest("[data-repair-id]");
      if (!trigger) return;
      const repairId = String(trigger.getAttribute("data-repair-id") || "");
      if (!repairId) return;
      const repair = getRepairById(repairId);
      if (repair) openRepairModal(repair);
    });
  }

  if (bonsBoard) {
    const resetBonsDragState = () => {
      state.bonsDrag.activeRepairId = null;
      state.bonsDrag.sourceColumnKey = null;
      state.bonsDrag.overColumnKey = null;
      state.bonsDrag.dropTargetKey = null;
    };

    const setBonsDragOverColumn = (columnKey) => {
      const normalized = isBonsColumnKey(columnKey) ? columnKey : null;
      if (state.bonsDrag.overColumnKey === normalized) return;
      state.bonsDrag.overColumnKey = normalized;
      bonsBoard.querySelectorAll(".bons-column.is-drop-target").forEach((el) => el.classList.remove("is-drop-target"));
      if (normalized) {
        const next = bonsBoard.querySelector(`[data-column-key="${normalized}"]`);
        if (next) next.classList.add("is-drop-target");
      }
    };

    const finalizeBonsMove = async (activeRepairId, sourceColumnKey, targetColumnKey) => {
      if (!isBonsColumnKey(targetColumnKey) || targetColumnKey === sourceColumnKey) {
        state.bonsDragIgnoreClickUntil = Date.now() + 350;
        return;
      }

      const statusLabel = getRepairStatusLabelForColumnKey(targetColumnKey);
      if (!statusLabel) return;

      const repair = getRepairById(activeRepairId);
      if (!repair) return;

      state.bonsDrag.isSaving = true;
      renderBonsBoard();
      try {
        await updateRepairRecord(repair.id, { status_label: statusLabel });
        renderAccueilCards();
        renderBonsBoard();
        renderFacturesList();
      } catch (err) {
        const metaEl = $("bonsMeta");
        if (metaEl) metaEl.textContent = err?.message || "Impossible de déplacer ce bon.";
        renderBonsBoard();
      } finally {
        state.bonsDrag.isSaving = false;
        state.bonsDragIgnoreClickUntil = Date.now() + 350;
        renderBonsBoard();
      }
    };

    const bonsPointerDrag = {
      pendingRepairId: null,
      pendingSourceKey: null,
      pendingElement: null,
      startX: 0,
      startY: 0,
      lastX: 0,
      lastY: 0,
      started: false,
      holdTimer: null,
      overlayEl: null
    };
    const HOLD_TO_DRAG_MS = 1000;
    const HOLD_MOVE_TOLERANCE_PX = 14;

    const updateBonsDragOverlayPosition = (x, y) => {
      const overlay = bonsPointerDrag.overlayEl;
      if (!overlay) return;
      overlay.style.left = `${Math.round(x)}px`;
      overlay.style.top = `${Math.round(y)}px`;
    };

    const destroyBonsDragOverlay = () => {
      const overlay = bonsPointerDrag.overlayEl;
      if (overlay?.parentNode) overlay.parentNode.removeChild(overlay);
      bonsPointerDrag.overlayEl = null;
    };

    const createBonsDragOverlay = () => {
      destroyBonsDragOverlay();
      const sourceEl = bonsPointerDrag.pendingElement;
      if (!sourceEl) return;
      const clone = sourceEl.cloneNode(true);
      if (!(clone instanceof HTMLElement)) return;
      clone.classList.add("bons-drag-overlay-card");
      clone.setAttribute("aria-hidden", "true");
      clone.removeAttribute("tabindex");
      clone.querySelectorAll("[id]").forEach((el) => el.removeAttribute("id"));
      clone.style.width = "140px";
      document.body.appendChild(clone);
      bonsPointerDrag.overlayEl = clone;
      updateBonsDragOverlayPosition(bonsPointerDrag.lastX || bonsPointerDrag.startX, bonsPointerDrag.lastY || bonsPointerDrag.startY);
    };

    const clearBonsPointerDrag = () => {
      if (bonsPointerDrag.holdTimer) {
        clearTimeout(bonsPointerDrag.holdTimer);
      }
      bonsPointerDrag.holdTimer = null;
      bonsPointerDrag.pendingRepairId = null;
      bonsPointerDrag.pendingSourceKey = null;
      bonsPointerDrag.pendingElement = null;
      bonsPointerDrag.startX = 0;
      bonsPointerDrag.startY = 0;
      bonsPointerDrag.lastX = 0;
      bonsPointerDrag.lastY = 0;
      bonsPointerDrag.started = false;
      document.body.classList.remove("bons-drag-noselect");
      destroyBonsDragOverlay();
    };

    const startBonsHeldDrag = () => {
      if (!bonsPointerDrag.pendingRepairId || bonsPointerDrag.started || state.bonsDrag.isSaving) return;

      bonsPointerDrag.started = true;
      state.bonsDrag.activeRepairId = bonsPointerDrag.pendingRepairId;
      state.bonsDrag.sourceColumnKey = bonsPointerDrag.pendingSourceKey;
      state.bonsDrag.overColumnKey = bonsPointerDrag.pendingSourceKey;
      state.bonsDrag.dropTargetKey = null;
      state.bonsDragIgnoreClickUntil = Date.now() + 450;
      renderBonsBoard();
      setBonsDragOverColumn(bonsPointerDrag.pendingSourceKey);
      createBonsDragOverlay();
    };

    const onBonsPointerMove = (ev) => {
      if (!bonsPointerDrag.pendingRepairId) return;
      bonsPointerDrag.lastX = ev.clientX;
      bonsPointerDrag.lastY = ev.clientY;

      if (!bonsPointerDrag.started) {
        const dx = Math.abs(ev.clientX - bonsPointerDrag.startX);
        const dy = Math.abs(ev.clientY - bonsPointerDrag.startY);
        if (Math.max(dx, dy) > HOLD_MOVE_TOLERANCE_PX) {
          clearBonsPointerDrag();
          window.removeEventListener("mousemove", onBonsPointerMove);
          window.removeEventListener("mouseup", onBonsPointerUp);
          window.removeEventListener("blur", onBonsPointerUp);
        }
        return;
      }

      ev.preventDefault();
      updateBonsDragOverlayPosition(ev.clientX, ev.clientY);
      const columnKey = getBonsColumnKeyFromDragEvent(ev, bonsBoard);
      setBonsDragOverColumn(columnKey);
    };

    const onBonsPointerUp = async () => {
      if (!bonsPointerDrag.pendingRepairId) return;

      const activeRepairId = bonsPointerDrag.pendingRepairId;
      const sourceColumnKey = bonsPointerDrag.pendingSourceKey;
      const hasStarted = bonsPointerDrag.started;
      const targetColumnKey = state.bonsDrag.overColumnKey;

      clearBonsPointerDrag();
      window.removeEventListener("mousemove", onBonsPointerMove);
      window.removeEventListener("mouseup", onBonsPointerUp);
      window.removeEventListener("blur", onBonsPointerUp);

      if (!hasStarted) return;

      resetBonsDragState();
      renderBonsBoard();
      await finalizeBonsMove(activeRepairId, sourceColumnKey, targetColumnKey);
    };

    bonsBoard.addEventListener("mousedown", (ev) => {
      if (ev.button !== 0) return;
      if (state.bonsDrag.isSaving) return;
      if (bonsPointerDrag.pendingRepairId) return;
      if (ev.target?.closest?.(".bons-pill-create-facture-btn")) return;
      const btnRepair = ev.target?.closest?.(".bons-repair-btn[data-repair-id]");
      if (!btnRepair) return;

      const repair = getRepairById(btnRepair.getAttribute("data-repair-id"));
      if (!repair) return;

      const sourceKey = canonicalRepairStatus(repair.status_label);
      if (!isBonsColumnKey(sourceKey)) return;

      bonsPointerDrag.pendingRepairId = String(repair.id);
      bonsPointerDrag.pendingSourceKey = sourceKey;
      bonsPointerDrag.pendingElement = btnRepair;
      bonsPointerDrag.startX = ev.clientX;
      bonsPointerDrag.startY = ev.clientY;
      bonsPointerDrag.lastX = ev.clientX;
      bonsPointerDrag.lastY = ev.clientY;
      bonsPointerDrag.started = false;
      document.body.classList.add("bons-drag-noselect");

      bonsPointerDrag.holdTimer = setTimeout(startBonsHeldDrag, HOLD_TO_DRAG_MS);
      window.addEventListener("mousemove", onBonsPointerMove);
      window.addEventListener("mouseup", onBonsPointerUp);
      window.addEventListener("blur", onBonsPointerUp);
    });

    bonsBoard.addEventListener("click", async (ev) => {
      if (Date.now() < (state.bonsDragIgnoreClickUntil || 0)) return;
      if (state.bonsDrag.isSaving) return;
      if (!(ev.target instanceof Element)) return;

      const createFactureBtn = ev.target.closest(".bons-pill-create-facture-btn");
      if (createFactureBtn) {
        ev.preventDefault();
        ev.stopPropagation();
        const repairBtn = createFactureBtn.closest(".bons-repair-btn[data-repair-id]");
        const repair = repairBtn ? getRepairById(repairBtn.getAttribute("data-repair-id")) : null;
        if (repair) {
          await promptRepairFactureCreation(repair);
        }
        return;
      }

      const btnRepair = ev.target.closest("[data-repair-id]");
      if (!btnRepair) return;
      const repair = getRepairById(btnRepair.getAttribute("data-repair-id"));
      if (repair) openRepairModal(repair);
    });
  }

  if (clientsSearchInput) {
    clientsSearchInput.addEventListener("input", () => {
      state.clientsSearchQuery = clientsSearchInput.value || "";
      renderClientsList();
    });
  }

  if (clientsRefreshBtn) {
    clientsRefreshBtn.addEventListener("click", async () => {
      clientsRefreshBtn.disabled = true;
      try {
        await refreshClientsData();
      } catch (err) {
        state.clientsLoadError = `Clients: ${err?.message || "Impossible de charger les clients."}`;
      } finally {
        renderClientsList();
        clientsRefreshBtn.disabled = false;
      }
    });
  }

  if (clientsAddBtn) {
    clientsAddBtn.addEventListener("click", () => {
      openCreateClientModal();
    });
  }

  if (projectsCreateBtn) {
    projectsCreateBtn.addEventListener("click", () => {
      openProjectModal();
    });
  }

  if (expensesCreateBtn) {
    expensesCreateBtn.addEventListener("click", () => {
      openExpenseModal();
    });
  }

  if (expensesTableBody) {
    expensesTableBody.addEventListener("click", async (ev) => {
      if (!(ev.target instanceof Element)) return;
      const invoiceBtn = ev.target.closest("[data-expense-open-invoice]");
      if (invoiceBtn) {
        const expense = getExpenseById(invoiceBtn.getAttribute("data-expense-open-invoice"));
        if (!expense) return;
        try {
          await openExpenseInvoice(expense);
        } catch (err) {
          window.alert(err?.message || "Impossible d'ouvrir la facture.");
        }
        return;
      }
      const openBtn = ev.target.closest("[data-expense-id]");
      if (!openBtn) return;
      const expense = getExpenseById(openBtn.getAttribute("data-expense-id"));
      if (expense) openExpenseModal({ row: expense });
    });
  }

  const bindProjectListClick = (targetList) => {
    if (!targetList) return;
    targetList.addEventListener("click", (ev) => {
      if (!(ev.target instanceof Element)) return;
      const card = ev.target.closest("[data-project-id]");
      if (!card) return;
      const projectId = String(card.getAttribute("data-project-id") || "").trim();
      if (!projectId) return;
      const project = getProjectById(projectId);
      if (project) openProjectModal({ row: project });
    });
  };
  bindProjectListClick(projectsList);
  bindProjectListClick(projectsArchivedList);

  if (expenseModalCloseBtn) {
    expenseModalCloseBtn.addEventListener("click", async () => {
      await closeExpenseModal();
    });
  }
  if (expenseModal) {
    expenseModal.addEventListener("click", async (ev) => {
      if (!(ev.target instanceof Element)) return;
      if (ev.target.closest("[data-expense-modal-close='true']")) {
        await closeExpenseModal();
      }
    });
  }

  if (bonsProjectsTabButtons.length) {
    bonsProjectsTabButtons.forEach((btn) => {
      btn.addEventListener("click", () => {
        const tab = btn.getAttribute("data-bons-projects-tab");
        setBonsProjectsTab(tab);
      });
    });
  }

  if (bonsRepairsTabButtons.length) {
    bonsRepairsTabButtons.forEach((btn) => {
      btn.addEventListener("click", () => {
        const tab = btn.getAttribute("data-bons-repairs-tab");
        setBonsRepairsTab(tab);
      });
    });
  }

  if (projectsTabButtons.length) {
    projectsTabButtons.forEach((btn) => {
      btn.addEventListener("click", () => {
        const tab = btn.getAttribute("data-projects-tab");
        setProjectsTab(tab);
      });
    });
  }

  if (facturesCreateBtn) {
    facturesCreateBtn.addEventListener("click", async () => {
      if (state.facturationActiveTab === "factures") {
        openCreateFactureModal();
        return;
      }

      const typeDocument = state.facturationActiveTab === "soumissions" ? "soumission" : "estimation";
      openCreateDevisModal(typeDocument);
    });
  }

  if (facturesSearchInput) {
    facturesSearchInput.value = state.facturationActiveTab === "factures"
      ? (state.facturesSearchQuery || "")
      : (state.devisSearchQuery || "");
    facturesSearchInput.addEventListener("input", () => {
      if (state.facturationActiveTab === "factures") {
        state.facturesSearchQuery = facturesSearchInput.value || "";
      } else {
        state.devisSearchQuery = facturesSearchInput.value || "";
      }
      renderFacturesList();
    });
  }

  if (facturesYearFilterSelect) {
    facturesYearFilterSelect.value = state.facturesYearFilter || "all";
    facturesYearFilterSelect.addEventListener("change", () => {
      state.facturesYearFilter = facturesYearFilterSelect.value || "all";
      renderFacturesList();
    });
  }

  if (facturesStatusFilterSelect) {
    facturesStatusFilterSelect.value = state.facturesStatusFilter || "all";
    facturesStatusFilterSelect.addEventListener("change", () => {
      state.facturesStatusFilter = facturesStatusFilterSelect.value || "all";
      renderFacturesList();
    });
  }

  if (devisStatusFilterSelect) {
    devisStatusFilterSelect.value = state.devisStatusFilter || "all";
    devisStatusFilterSelect.addEventListener("change", () => {
      state.devisStatusFilter = devisStatusFilterSelect.value || "all";
      renderFacturesList();
    });
  }

  if (facturationTabButtons.length) {
    facturationTabButtons.forEach((btn) => {
      btn.addEventListener("click", () => {
        const tab = btn.getAttribute("data-facturation-tab");
        setFacturationTab(tab);
      });
    });
  }

  if (facturesTableBody) {
    facturesTableBody.addEventListener("click", async (ev) => {
      if (!(ev.target instanceof Element)) return;
      const btnPreview = ev.target.closest("[data-facture-preview-id]");
      if (btnPreview) {
        const factureId = String(btnPreview.getAttribute("data-facture-preview-id") ?? "").trim();
        if (!factureId) return;

        const facture = (state.factures || []).find((f) => String(f?.id) === factureId);
        if (!facture) return;

        try {
          await openFacturePdfPreview(facture);
        } catch (err) {
          window.alert(err?.message || "Impossible d'ouvrir le PDF.");
        }
        return;
      }

      const btnEdit = ev.target.closest("[data-facture-edit-id]");
      if (btnEdit) {
        const factureId = String(btnEdit.getAttribute("data-facture-edit-id") ?? "").trim();
        if (!factureId) return;
        try {
          openCreateFactureModal({ factureId });
        } catch (err) {
          window.alert(err?.message || "Impossible d'ouvrir cette facture en modification.");
        }
        return;
      }

      const btnDevisPreview = ev.target.closest("[data-estimation-preview-id]");
      if (btnDevisPreview) {
        const docId = String(btnDevisPreview.getAttribute("data-estimation-preview-id") ?? "").trim();
        if (!docId) return;
        const doc = (state.estimations || []).find((row) => String(row?.id) === docId);
        if (!doc) return;
        try {
          await openDevisPdfPreview(doc);
        } catch (err) {
          window.alert(err?.message || "Impossible d'ouvrir le PDF.");
        }
        return;
      }

      if (state.facturationActiveTab !== "factures") {
        const rowEl = ev.target.closest("tr[data-estimation-edit-id]");
        if (!rowEl) return;
        const docId = String(rowEl.getAttribute("data-estimation-edit-id") ?? "").trim();
        if (!docId) return;
        const doc = (state.estimations || []).find((item) => String(item?.id) === docId);
        if (!doc) return;
        openCreateDevisModal(doc?.type_document, { row: doc });
      }
    });
  }

  const setConfigEditing = (section, editing) => {
    const wanted = String(section ?? "").trim().toLowerCase();
    const enabled = Boolean(editing);
    if (wanted === "entreprise" || wanted === "achats") {
      state.configEditing.entreprise = wanted === "entreprise" && enabled;
      state.configEditing.achats = wanted === "achats" && enabled;
      state.configEditing.honoraires = false;
    } else {
      state.configEditing.honoraires = enabled;
      state.configEditing.entreprise = false;
      state.configEditing.achats = false;
    }
    if (state.configEditing.achats) {
      state.configPurchaseLinksDraft = getConfigPurchaseLinksDraftValue();
    } else {
      state.configPurchaseLinksDraft = null;
    }
    renderConfigurationView(true);
  };

  const parseRate = (value) => {
    const raw = String(value ?? "").trim().replace(",", ".");
    const n = Number(raw);
    return Number.isFinite(n) && n >= 0 ? n : NaN;
  };

  const saveHonorairesConfiguration = async () => {
    if (
      !configLaborRateStandardInput
      || !configLaborRateSecondaryInput
      || !configLaborRateDiagnosticInput
      || !configLaborRateEstimationInput
    ) return;

    const standard = parseRate(configLaborRateStandardInput.value);
    const secondary = parseRate(configLaborRateSecondaryInput.value);
    const diagnostic = parseRate(configLaborRateDiagnosticInput.value);
    const estimation = parseRate(configLaborRateEstimationInput.value);
    if (
      !Number.isFinite(standard)
      || !Number.isFinite(secondary)
      || !Number.isFinite(diagnostic)
      || !Number.isFinite(estimation)
    ) {
      setConfigFeedback(configMessage, "Les taux doivent être des nombres >= 0.");
      return;
    }

    try {
      if (configHonorairesSaveBtn) configHonorairesSaveBtn.disabled = true;
      setConfigFeedback(configMessage, "Enregistrement...");
      await saveAppConfig({
        labor_rate_standard: standard,
        labor_rate_secondary: secondary,
        labor_rate_diagnostic: diagnostic,
        labor_rate_estimation: estimation,
      });
      state.configEditing.honoraires = false;
      renderConfigurationView(true);
      setConfigFeedback(configMessage, "Configuration enregistrée.", { clearAfterMs: 5000 });
    } catch (err) {
      setConfigFeedback(configMessage, err?.message || "Impossible d'enregistrer la configuration.");
    } finally {
      if (configHonorairesSaveBtn) configHonorairesSaveBtn.disabled = false;
    }
  };

  const saveEntrepriseConfiguration = async () => {
    if (
      !configCompanyNameInput
      || !configCompanyPhoneInput
      || !configCompanyEmailInput
      || !configCompanyAddress1Input
      || !configCompanyAddress2Input
      || !configCompanyCityInput
      || !configCompanyProvinceInput
      || !configCompanyPostalCodeInput
      || !configCompanyCountryInput
      || !configCompanyLogoPathInput
      || !configTpsNumberInput
      || !configTpsRateInput
      || !configTvqNumberInput
      || !configTvqRateInput
    ) return;
    try {
      if (configEntrepriseSaveBtn) configEntrepriseSaveBtn.disabled = true;
      setConfigFeedback(configMessage, "Enregistrement...");
      await saveAppConfig({
        company_name: cleanNullableText(configCompanyNameInput.value),
        company_phone: cleanNullableText(configCompanyPhoneInput.value),
        company_email: cleanNullableText(configCompanyEmailInput.value),
        company_address_line1: cleanNullableText(configCompanyAddress1Input.value),
        company_address_line2: cleanNullableText(configCompanyAddress2Input.value),
        company_city: cleanNullableText(configCompanyCityInput.value),
        company_province: cleanNullableText(configCompanyProvinceInput.value),
        company_postal_code: cleanNullableText(configCompanyPostalCodeInput.value),
        company_country: cleanNullableText(configCompanyCountryInput.value),
        company_logo_path: cleanNullableText(configCompanyLogoPathInput.value),
        tax_tps_number: cleanNullableText(configTpsNumberInput.value),
        tax_tps_rate: normalizeTaxRateNumber(configTpsRateInput.value, DEFAULT_APP_CONFIG.tax_tps_rate),
        tax_tvq_number: cleanNullableText(configTvqNumberInput.value),
        tax_tvq_rate: normalizeTaxRateNumber(configTvqRateInput.value, DEFAULT_APP_CONFIG.tax_tvq_rate),
      });
      state.configEditing.entreprise = false;
      renderConfigurationView(true);
      setConfigFeedback(configMessage, "Configuration enregistrée.", { clearAfterMs: 5000 });
    } catch (err) {
      setConfigFeedback(configMessage, err?.message || "Impossible d'enregistrer la configuration.");
    } finally {
      if (configEntrepriseSaveBtn) configEntrepriseSaveBtn.disabled = false;
    }
  };

  const saveAchatsConfiguration = async () => {
    if (!Array.isArray(state.configPurchaseLinksDraft)) return;
    const cleanedLinks = state.configPurchaseLinksDraft
      .map((link, index) => {
        const src = link && typeof link === "object" ? link : {};
        const title = cleanNullableText(src.title);
        const url = sanitizeExternalUrl(src.url);
        const note = cleanNullableText(src.note) || "";
        if (!title && !url) return null;
        const key = normalizePurchaseLinkKey(src.key || title || `purchase_link_${index + 1}`, `purchase_link_${index + 1}`);
        return {
          key,
          title: title || `Lien ${index + 1}`,
          url: url || "",
          note,
          domain: extractDomainFromUrl(url, src.domain || ""),
        };
      })
      .filter(Boolean);

    try {
      if (configAchatsSaveBtn) configAchatsSaveBtn.disabled = true;
      setConfigFeedback(configMessage, "Enregistrement...");
      await saveAppConfig({
        purchase_links: cleanedLinks,
      });
      state.configEditing.achats = false;
      state.configPurchaseLinksDraft = null;
      renderConfigurationView(true);
      setConfigFeedback(configMessage, "Liens d'achats enregistrés.", { clearAfterMs: 5000 });
    } catch (err) {
      setConfigFeedback(configMessage, err?.message || "Impossible d'enregistrer les liens d'achats.");
    } finally {
      if (configAchatsSaveBtn) configAchatsSaveBtn.disabled = false;
    }
  };

  const saveDeveloppementConfiguration = async () => {
    if (!state.appConfigLoaded) return;
    try {
      if (configDeveloppementSaveBtn) configDeveloppementSaveBtn.disabled = true;
      setConfigFeedback(configMessage, "Enregistrement...");
      await saveAppConfig({
        developer_mode: getDeveloperModeDraftValue(),
      });
      state.configDeveloperDraft = null;
      renderConfigurationView(true);
      const modeDevLabel = state.appConfig?.developer_mode ? "activé" : "désactivé";
      setConfigFeedback(configMessage, `Mode développeur ${modeDevLabel}.`, { clearAfterMs: 5000 });
    } catch (err) {
      setConfigFeedback(configMessage, err?.message || "Impossible d'enregistrer le mode développeur.");
    } finally {
      if (configDeveloppementSaveBtn) configDeveloppementSaveBtn.disabled = false;
    }
  };

  if (configHonorairesEditBtn) {
    configHonorairesEditBtn.addEventListener("click", () => {
      setConfigFeedback(configMessage, "");
      setConfigEditing("honoraires", !state.configEditing.honoraires);
    });
  }
  if (configHonorairesCancelBtn) {
    configHonorairesCancelBtn.addEventListener("click", () => {
      setConfigFeedback(configMessage, "");
      setConfigEditing("honoraires", false);
    });
  }
  if (configHonorairesSaveBtn) {
    configHonorairesSaveBtn.addEventListener("click", () => {
      saveHonorairesConfiguration();
    });
  }
  if (configEntrepriseEditBtn) {
    configEntrepriseEditBtn.addEventListener("click", () => {
      setConfigFeedback(configMessage, "");
      setConfigEditing("entreprise", !state.configEditing.entreprise);
    });
  }
  if (configEntrepriseCancelBtn) {
    configEntrepriseCancelBtn.addEventListener("click", () => {
      setConfigFeedback(configMessage, "");
      setConfigEditing("entreprise", false);
    });
  }
  if (configEntrepriseSaveBtn) {
    configEntrepriseSaveBtn.addEventListener("click", () => {
      saveEntrepriseConfiguration();
    });
  }
  if (configAchatsEditBtn) {
    configAchatsEditBtn.addEventListener("click", () => {
      setConfigFeedback(configMessage, "");
      setConfigEditing("achats", !state.configEditing.achats);
    });
  }
  if (configAchatsCancelBtn) {
    configAchatsCancelBtn.addEventListener("click", () => {
      setConfigFeedback(configMessage, "");
      state.configPurchaseLinksDraft = null;
      setConfigEditing("achats", false);
    });
  }
  if (configAchatsSaveBtn) {
    configAchatsSaveBtn.addEventListener("click", () => {
      saveAchatsConfiguration();
    });
  }
  if (configPurchaseLinksList) {
    configPurchaseLinksList.addEventListener("click", (ev) => {
      if (!(ev.target instanceof Element)) return;
      const addBtn = ev.target.closest("[data-config-purchase-link-add]");
      if (addBtn) {
        if (!state.configEditing.achats) return;
        const draft = getConfigPurchaseLinksDraftValue();
        draft.push({
          key: `purchase_link_${draft.length + 1}`,
          title: "",
          url: "",
          note: "",
          domain: "",
        });
        state.configPurchaseLinksDraft = draft;
        renderConfigurationView(true);
        const lastIndex = draft.length - 1;
        const titleInput = configPurchaseLinksList.querySelector(
          `[data-config-purchase-link-field="title"][data-config-purchase-link-index="${lastIndex}"]`
        );
        if (titleInput instanceof HTMLElement) titleInput.focus();
        return;
      }

      const removeBtn = ev.target.closest("[data-config-purchase-link-remove]");
      if (!removeBtn || !state.configEditing.achats) return;
      const index = Number(removeBtn.getAttribute("data-config-purchase-link-remove"));
      if (!Number.isInteger(index) || index < 0) return;
      const draft = getConfigPurchaseLinksDraftValue();
      if (index >= draft.length) return;
      draft.splice(index, 1);
      state.configPurchaseLinksDraft = draft;
      renderConfigurationView(true);
    });

    configPurchaseLinksList.addEventListener("input", (ev) => {
      if (!(ev.target instanceof HTMLInputElement)) return;
      if (!state.configEditing.achats) return;
      const field = String(ev.target.getAttribute("data-config-purchase-link-field") || "").trim();
      const index = Number(ev.target.getAttribute("data-config-purchase-link-index"));
      if (!Number.isInteger(index) || index < 0) return;
      if (field !== "title" && field !== "url" && field !== "note") return;
      const draft = getConfigPurchaseLinksDraftValue();
      if (!draft[index]) return;
      draft[index][field] = ev.target.value || "";
      if (field === "url") {
        draft[index].domain = extractDomainFromUrl(draft[index].url, draft[index].domain || "");
      }
      if (field === "title" && !cleanNullableText(draft[index].key)) {
        draft[index].key = normalizePurchaseLinkKey(draft[index].title, `purchase_link_${index + 1}`);
      }
      state.configPurchaseLinksDraft = draft;
    });

    configPurchaseLinksList.addEventListener("keydown", (ev) => {
      if (ev.key !== "Enter") return;
      if (!(ev.target instanceof HTMLInputElement)) return;
      if (!state.configEditing.achats) return;
      if (!ev.target.hasAttribute("data-config-purchase-link-field")) return;
      ev.preventDefault();
      saveAchatsConfiguration();
    });
  }
  if (configTaxReportsList) {
    configTaxReportsList.addEventListener("click", async (ev) => {
      if (!(ev.target instanceof Element)) return;

      const openBtn = ev.target.closest("[data-tax-report-open]");
      if (openBtn) {
        const year = Number(openBtn.getAttribute("data-tax-report-open"));
        if (!Number.isInteger(year)) return;
        setConfigFeedback(configMessage, "");
        try {
          await openTaxReportPdf(year);
        } catch (err) {
          setConfigFeedback(configMessage, err?.message || "Impossible d'ouvrir le rapport fiscal.");
        }
        return;
      }

      const downloadBtn = ev.target.closest("[data-tax-report-download]");
      if (downloadBtn) {
        const year = Number(downloadBtn.getAttribute("data-tax-report-download"));
        if (!Number.isInteger(year)) return;
        setConfigFeedback(configMessage, "");
        try {
          await downloadTaxReportPdf(year);
        } catch (err) {
          setConfigFeedback(configMessage, err?.message || "Impossible de télécharger le rapport fiscal.");
        }
        return;
      }

      const generateBtn = ev.target.closest("[data-tax-report-generate]");
      if (!generateBtn) return;
      const year = Number(generateBtn.getAttribute("data-tax-report-generate"));
      if (!Number.isInteger(year)) return;
      try {
        setConfigFeedback(configMessage, `G\u00e9n\u00e9ration du rapport fiscal ${year}...`);
        await generateFiscalReport(year);
        setConfigFeedback(configMessage, `Rapport fiscal ${year} g\u00e9n\u00e9r\u00e9.`, { clearAfterMs: 5000 });
      } catch (err) {
        setConfigFeedback(configMessage, err?.message || "Impossible de g\u00e9n\u00e9rer le rapport fiscal.");
      }
    });
  }
  if (configDeveloppementCancelBtn) {
    configDeveloppementCancelBtn.addEventListener("click", () => {
      setConfigFeedback(configMessage, "");
      state.configDeveloperDraft = null;
      renderConfigurationView(true);
    });
  }
  if (configDeveloppementSaveBtn) {
    configDeveloppementSaveBtn.addEventListener("click", () => {
      saveDeveloppementConfiguration();
    });
  }
  const setDeveloperModeDraft = (nextValue) => {
    if (!state.appConfigLoaded) return;
    const wanted = Boolean(nextValue);
    const current = getDeveloperModeDraftValue();
    if (wanted === current) return;
    state.configDeveloperDraft = wanted === Boolean(state.appConfig?.developer_mode) ? null : wanted;
    setConfigFeedback(configMessage, "");
    renderConfigurationView(true);
  };
  if (configDeveloperModeSwitchBtn) {
    configDeveloperModeSwitchBtn.addEventListener("click", () => {
      setDeveloperModeDraft(!getDeveloperModeDraftValue());
    });
  }
  if (configDeveloperModeOffBtn) {
    configDeveloperModeOffBtn.addEventListener("click", () => {
      setDeveloperModeDraft(false);
    });
  }
  if (configDeveloperModeOnBtn) {
    configDeveloperModeOnBtn.addEventListener("click", () => {
      setDeveloperModeDraft(true);
    });
  }

  configTabButtons.forEach((btn) => {
    btn.addEventListener("click", () => {
      if (!(btn instanceof HTMLElement)) return;
      const tab = btn.getAttribute("data-config-tab");
      setConfigurationTab(tab);
    });
  });
  [
    configLaborRateStandardInput,
    configLaborRateSecondaryInput,
    configLaborRateDiagnosticInput,
    configLaborRateEstimationInput,
    configCompanyNameInput,
    configCompanyPhoneInput,
    configCompanyEmailInput,
    configCompanyAddress1Input,
    configCompanyAddress2Input,
    configCompanyCityInput,
    configCompanyProvinceInput,
    configCompanyPostalCodeInput,
    configCompanyCountryInput,
    configCompanyLogoPathInput,
    configTpsNumberInput,
    configTpsRateInput,
    configTvqNumberInput,
    configTvqRateInput,
  ].forEach((field) => {
    if (!field) return;
    field.addEventListener("keydown", (ev) => {
      if (ev.key !== "Enter") return;
      ev.preventDefault();
      if (state.configEditing.honoraires && (
        field === configLaborRateStandardInput
        || field === configLaborRateSecondaryInput
        || field === configLaborRateDiagnosticInput
        || field === configLaborRateEstimationInput
      )) {
        saveHonorairesConfiguration();
        return;
      }
      if (state.configEditing.entreprise && (
        field === configCompanyNameInput
        || field === configCompanyPhoneInput
        || field === configCompanyEmailInput
        || field === configCompanyAddress1Input
        || field === configCompanyAddress2Input
        || field === configCompanyCityInput
        || field === configCompanyProvinceInput
        || field === configCompanyPostalCodeInput
        || field === configCompanyCountryInput
        || field === configCompanyLogoPathInput
        || field === configTpsNumberInput
        || field === configTpsRateInput
        || field === configTvqNumberInput
        || field === configTvqRateInput
      )) {
        saveEntrepriseConfiguration();
        return;
      }
    });
  });

  if (clientsTableBody) {
    clientsTableBody.addEventListener("click", (ev) => {
      if (!(ev.target instanceof Element)) return;
      const row = ev.target.closest("[data-client-id]");
      if (!row) return;
      const clientRow = getClientById(row.getAttribute("data-client-id"));
      if (clientRow) openClientModal(clientRow);
    });
  }

  if (repairModalCloseBtn) repairModalCloseBtn.addEventListener("click", () => closeRepairModal());
  if (repairModal) {
    repairModal.addEventListener("click", (ev) => {
      if (ev.target && ev.target.getAttribute("data-modal-close") === "true") closeRepairModal();
    });
  }

  if (clientModalCloseBtn) clientModalCloseBtn.addEventListener("click", () => closeClientModal());
  if (clientModal) {
    clientModal.addEventListener("click", (ev) => {
      if (ev.target && ev.target.getAttribute("data-client-modal-close") === "true") closeClientModal();
    });
  }
  if (clientModalEditBtn) {
    clientModalEditBtn.addEventListener("click", () => {
      toggleClientModalEditMode();
    });
  }

  if (factureModalCloseBtn) factureModalCloseBtn.addEventListener("click", () => closeFactureModal());
  if (factureModal) {
    factureModal.addEventListener("click", (ev) => {
      if (ev.target && ev.target.getAttribute("data-facture-modal-close") === "true") closeFactureModal();
    });
  }

  if (projectModalCloseBtn) projectModalCloseBtn.addEventListener("click", () => closeProjectModal());
  if (projectModal) {
    projectModal.addEventListener("click", (ev) => {
      if (ev.target && ev.target.getAttribute("data-project-modal-close") === "true") closeProjectModal();
    });
  }

  if (facturePdfModalCloseBtn) facturePdfModalCloseBtn.addEventListener("click", () => closeFacturePdfModal());
  if (facturePdfModal) {
    facturePdfModal.addEventListener("click", (ev) => {
      if (ev.target && ev.target.getAttribute("data-facture-pdf-modal-close") === "true") closeFacturePdfModal();
    });
  }
  if (facturePdfPrevBtn) {
    facturePdfPrevBtn.addEventListener("click", () => {
      if (facturePdfModalState.loading || !facturePdfModalState.doc) return;
      if (facturePdfModalState.page <= 1) return;
      facturePdfModalState.page -= 1;
      renderFacturePdfPage();
    });
  }
  if (facturePdfNextBtn) {
    facturePdfNextBtn.addEventListener("click", () => {
      if (facturePdfModalState.loading || !facturePdfModalState.doc) return;
      if (facturePdfModalState.page >= facturePdfModalState.totalPages) return;
      facturePdfModalState.page += 1;
      renderFacturePdfPage();
    });
  }
  if (facturePdfZoomOutBtn) {
    facturePdfZoomOutBtn.addEventListener("click", () => {
      if (facturePdfModalState.loading || !facturePdfModalState.doc) return;
      facturePdfModalState.scale = Math.max(0.55, Math.round((facturePdfModalState.scale - 0.25) * 100) / 100);
      renderFacturePdfPage();
    });
  }
  if (facturePdfZoomInBtn) {
    facturePdfZoomInBtn.addEventListener("click", () => {
      if (facturePdfModalState.loading || !facturePdfModalState.doc) return;
      facturePdfModalState.scale = Math.min(2.8, Math.round((facturePdfModalState.scale + 0.25) * 100) / 100);
      renderFacturePdfPage();
    });
  }

  setActiveView(getViewFromHash());
}

function handleSessionForPage(session) {
  if (isLoginPage()) {
    if (session?.user) goToApp();
    return;
  }

  if (isAppPage()) {
    if (!session?.user) {
      state.dashboardLoadedFor = null;
      goToLogin();
      return;
    }
    updateMenuUser();
    void updateMenuCompanyLogo();
    loadProfile();
    loadDashboard();
  }
}

async function initAuthFlow() {
  const client = ensureAuthClient();
  if (!client) return;

  client.auth.onAuthStateChange((_event, session) => {
    if (!state.sessionReady && !session?.user) return;
    state.session = session ?? null;
    handleSessionForPage(state.session);
  });

  const { data, error } = await client.auth.getSession();
  if (error) {
    setLoginStatus(error.message || "Session invalide.", "bad");
    state.sessionReady = true;
    if (isAppPage()) goToLogin();
    return;
  }

  state.session = data.session ?? null;
  state.sessionReady = true;
  handleSessionForPage(state.session);
}

function bootstrap() {
  initTheme();
  bindLoginUi();
  bindAppUi();
  initAuthFlow();
}

bootstrap();




