(function () {
  "use strict";

  const MAX_IMAGE_SIZE = 10 * 1024 * 1024;
  const ALLOWED_IMAGE_TYPES = new Set(["image/jpeg", "image/png", "image/webp"]);
  const app = document.getElementById("app");
  const toast = document.getElementById("toast");
  const backToTopButton = document.querySelector(".back-to-top");
  const coreDefinitions = window.AppData.coreNoteDefinitions;
  const CUSTOMER_OUTPUT_LABELS = {
    scheduleReason: {
      title: "なぜ・いつ",
      subtitle: "家づくりのきっかけと計画時期",
      point: "いつまでに、なぜ家づくりを進めたいのかを振り返ります。"
    },
    budget: {
      title: "予算",
      subtitle: "無理なく進めるための資金計画",
      point: "総額だけでなく、月々の支払いとのバランスを確認します。"
    },
    building: {
      title: "お家",
      subtitle: "間取り・性能・お家の希望",
      point: "家に求めることを整理します。"
    },
    land: {
      title: "土地",
      subtitle: "どこで、どんな暮らしをするか",
      point: "通勤・通学・周辺環境など、暮らしやすさの条件を確認します。"
    },
    image: {
      title: "イメージ",
      subtitle: "参考にした写真・図面・雰囲気",
      point: "打ち合わせで共有したイメージを、写真や図面で振り返ります。"
    }
  };
  const CUSTOMER_OUTPUT_ORDER = ["scheduleReason", "budget", "building", "land", "image"];
  const CURRENT_STATUS_FIELDS = [
    { key: "currentResidence", label: "今のお住まい" },
    { key: "currentRent", label: "今の家賃", unit: "万円", money: true },
    { key: "ownFunds", label: "自己資金", unit: "万円", money: true },
    { key: "annualIncome", label: "ご年収", unit: "万円", money: true },
    { key: "desiredArea", label: "住みたい場所" }
  ];
  const MONEY_STATUS_KEYS = new Set(CURRENT_STATUS_FIELDS.filter(function (field) { return field.money; }).map(function (field) { return field.key; }));
  const CSV_COLUMNS = [
    "sheetId",
    "customerName",
    "meetingDate",
    "meetingNumber",
    "title",
    "currentResidence",
    "currentRent",
    "ownFunds",
    "annualIncome",
    "desiredArea",
    "scheduleReason",
    "budget",
    "building",
    "land",
    "wife",
    "husband",
    "parents"
  ];

  let state = loadAppData();
  let pendingSaveTimer = null;
  let saveStatusNode = null;
  let currentView = { name: "start", sheetId: null, mode: "input" };
  const issuedIds = new Set();
  const activeObjectUrls = new Set();
  const imageMemoryStore = new Map();
  let imageStoreReadyPromise = null;
  let activeImageRenderPromises = [];

  // ---------------------------------------------------------------------------
  // Data layer. Customer and meeting data is kept only in JavaScript memory
  // while the page is open. CSV import/export is the only durable handoff.
  // ---------------------------------------------------------------------------

  function loadAppData() {
    return window.AppData.createEmptyData();
  }

  function saveAppData() {
    setSaveStatus("ブラウザには保存されません。必要に応じてCSVを書き出してください。", false);
    return true;
  }

  function normalizeAppData(data) {
    if (!data || !Array.isArray(data.sheets)) throw new Error("Invalid app data");
    const normalized = {
      version: 2,
      activeSheetId: data.activeSheetId || null,
      sheets: data.sheets.map(normalizeSheet)
    };
    if (normalized.activeSheetId && !normalized.sheets.some(function (sheet) { return sheet.sheetId === normalized.activeSheetId; })) {
      normalized.activeSheetId = null;
    }
    return normalized;
  }

  function createEmptyCurrentStatus() {
    return CURRENT_STATUS_FIELDS.reduce(function (result, field) {
      result[field.key] = "";
      return result;
    }, {});
  }

  function normalizeCurrentStatus(source) {
    const values = source || {};
    return CURRENT_STATUS_FIELDS.reduce(function (result, field) {
      const value = String(values[field.key] || "");
      result[field.key] = field.money ? normalizeMoneyInput(value) : value;
      return result;
    }, {});
  }

  function normalizeMoneyInput(value) {
    return String(value || "")
      .replace(/[０-９]/g, function (char) { return String.fromCharCode(char.charCodeAt(0) - 0xFEE0); })
      .replace(/万円/g, "")
      .replace(/[,\s]/g, "")
      .replace(/[^\d.]/g, "");
  }

  function formatCurrentStatusValue(key, value) {
    const normalized = MONEY_STATUS_KEYS.has(key) ? normalizeMoneyInput(value) : String(value || "").trim();
    if (!normalized) return "";
    return MONEY_STATUS_KEYS.has(key) ? normalized + "万円" : normalized;
  }

  function normalizeSheet(source) {
    const sheet = Object.assign({}, source);
    sheet.customerName = String(sheet.customerName || "");
    sheet.meetingDate = String(sheet.meetingDate || "");
    sheet.meetingNumber = Number(sheet.meetingNumber || 1);
    sheet.title = String(sheet.title || "");
    sheet.currentStatus = normalizeCurrentStatus(sheet.currentStatus);
    sheet.coreNotes = sheet.coreNotes || {};
    coreDefinitions.forEach(function (definition) {
      const current = sheet.coreNotes[definition.key] || {};
      sheet.coreNotes[definition.key] = { label: definition.label, content: String(current.content || "") };
    });

    const sourceStakeholders = Array.isArray(sheet.stakeholderNotes) ? sheet.stakeholderNotes : [];
    const ordered = [];
    ["wife", "husband", "parents"].forEach(function (key) {
      const found = sourceStakeholders.find(function (item) { return item.key === key; });
      ordered.push(found || {
        id: "ST_" + sheet.sheetId + "_" + key,
        type: "default",
        key: key,
        label: key === "wife" ? "奥様" : key === "husband" ? "ご主人様" : "ご両親",
        content: "",
        deletable: false
      });
    });
    sourceStakeholders.filter(function (item) { return !["wife", "husband", "parents"].includes(item.key); }).forEach(function (item) {
      ordered.push(item);
    });
    sheet.stakeholderNotes = ordered.map(function (item) {
      return {
        id: String(item.id),
        type: item.type === "custom" ? "custom" : "default",
        key: String(item.key),
        label: String(item.label || "補足項目"),
        content: String(item.content || ""),
        deletable: item.type === "custom" || item.deletable === true
      };
    });

    sheet.attachments = (Array.isArray(sheet.attachments) ? sheet.attachments : [])
      .filter(function (attachment) { return attachment && isCoreKey(attachment.categoryKey); })
      .map(function (attachment, index) {
        return {
          imageId: String(attachment.imageId),
          categoryKey: String(attachment.categoryKey),
          fileName: String(attachment.fileName || "画像"),
          mimeType: String(attachment.mimeType || "application/octet-stream"),
          caption: String(attachment.caption || ""),
          includeInCustomerOutput: attachment.includeInCustomerOutput !== false,
          order: Number(attachment.order || index + 1),
          createdAt: attachment.createdAt || new Date().toISOString()
        };
      });

    const preview = sheet.customerPreview || {};
    sheet.customerPreview = window.AppData.createCustomerPreview(
      Array.isArray(preview.visibleCategories) ? preview.visibleCategories.filter(isCoreKey) : [],
      preview.editedContents || {},
      Array.isArray(preview.visibleImageIds)
        ? preview.visibleImageIds.filter(function (imageId) { return sheet.attachments.some(function (attachment) { return attachment.imageId === imageId; }); })
        : sheet.attachments.filter(function (attachment) { return attachment.includeInCustomerOutput; }).map(function (attachment) { return attachment.imageId; })
    );
    sheet.customerPreview.visibilityTouched = preview.visibilityTouched === true;
    sheet.createdAt = sheet.createdAt || new Date().toISOString();
    sheet.updatedAt = sheet.updatedAt || sheet.createdAt;
    ensureCustomerPreviewShape(sheet);
    syncAutomaticCustomerPreviewVisibility(sheet);
    return sheet;
  }

  function getActiveSheet() {
    return getSheetById(state.activeSheetId);
  }

  function getSheetById(sheetId) {
    return state.sheets.find(function (sheet) { return sheet.sheetId === sheetId; }) || null;
  }

  function createBlankSheet(data) {
    const now = new Date().toISOString();
    const sheet = {
      sheetId: generateUniqueId("S"),
      customerName: String(data.customerName || ""),
      meetingDate: String(data.meetingDate || todayIso()),
      meetingNumber: Number(data.meetingNumber || 1),
      title: String(data.title || ""),
      currentStatus: createEmptyCurrentStatus(),
      coreNotes: window.AppData.createCoreNotes(),
      stakeholderNotes: createDefaultStakeholders(),
      attachments: [],
      customerPreview: window.AppData.createCustomerPreview([]),
      createdAt: now,
      updatedAt: now
    };
    state.sheets.push(sheet);
    state.activeSheetId = sheet.sheetId;
    saveAppData();
    return sheet;
  }

  function updateSheet(sheetId, data) {
    const sheet = getSheetById(sheetId);
    if (!sheet) return null;
    Object.assign(sheet, data);
    scheduleSave(sheet);
    return sheet;
  }

  async function deleteSheet(sheetId) {
    const sheet = getSheetById(sheetId);
    if (!sheet) return;
    if (sheet.attachments.length) await deleteImagesBySheetId(sheetId);
    state.sheets = state.sheets.filter(function (sheet) { return sheet.sheetId !== sheetId; });
    if (state.activeSheetId === sheetId) state.activeSheetId = null;
    saveAppData();
  }

  function setActiveSheet(sheetId) {
    if (sheetId && !getSheetById(sheetId)) return;
    state.activeSheetId = sheetId || null;
    saveAppData();
  }

  async function clearLocalData() {
    await clearImageMemoryStore();
    issuedIds.clear();
    state = window.AppData.createEmptyData();
  }

  function updateCoreNote(sheetId, key, content) {
    const sheet = getSheetById(sheetId);
    if (!sheet || !isCoreKey(key)) return;
    ensureCustomerPreviewShape(sheet);
    sheet.coreNotes[key].content = content;
    syncAutomaticCustomerPreviewVisibility(sheet);
    scheduleSave(sheet);
  }

  function updateCurrentStatus(sheetId, key, value) {
    const sheet = getSheetById(sheetId);
    if (!sheet || !CURRENT_STATUS_FIELDS.some(function (field) { return field.key === key; })) return;
    sheet.currentStatus = normalizeCurrentStatus(sheet.currentStatus);
    sheet.currentStatus[key] = MONEY_STATUS_KEYS.has(key) ? normalizeMoneyInput(value) : String(value || "");
    scheduleSave(sheet);
  }

  function updateStakeholderNote(sheetId, stakeholderId, content) {
    const stakeholder = findStakeholder(sheetId, stakeholderId);
    if (!stakeholder) return;
    stakeholder.item.content = content;
    scheduleSave(stakeholder.sheet);
  }

  function updateStakeholderLabel(sheetId, stakeholderId, label) {
    const stakeholder = findStakeholder(sheetId, stakeholderId);
    if (!stakeholder || !stakeholder.item.deletable) return;
    stakeholder.item.label = label;
    scheduleSave(stakeholder.sheet);
  }

  function addStakeholderNote(sheetId, label) {
    const sheet = getSheetById(sheetId);
    if (!sheet) return null;
    const id = generateUniqueId("ST");
    const item = { id: id, type: "custom", key: "custom_" + id, label: label, content: "", deletable: true };
    sheet.stakeholderNotes.push(item);
    scheduleSave(sheet);
    return item;
  }

  function deleteStakeholderNote(sheetId, stakeholderId) {
    const sheet = getSheetById(sheetId);
    if (!sheet) return;
    const item = sheet.stakeholderNotes.find(function (stakeholder) { return stakeholder.id === stakeholderId; });
    if (!item || !item.deletable) return;
    sheet.stakeholderNotes = sheet.stakeholderNotes.filter(function (stakeholder) { return stakeholder.id !== stakeholderId; });
    scheduleSave(sheet);
  }

  function initializeCustomerPreview(sheetId) {
    const sheet = getSheetById(sheetId);
    if (!sheet) return;
    ensureCustomerPreviewShape(sheet);
    syncAutomaticCustomerPreviewVisibility(sheet);
  }

  function updateCustomerPreviewContent(sheetId, key, content) {
    const sheet = getSheetById(sheetId);
    if (!sheet || !isCoreKey(key)) return;
    ensureCustomerPreviewShape(sheet);
    const original = getCoreNoteContent(sheet, key);
    const nextContent = String(content || "");
    sheet.customerPreview.editedContents[key] = hasText(nextContent) && nextContent !== original ? nextContent : "";
    syncAutomaticCustomerPreviewVisibility(sheet);
    scheduleSave(sheet);
  }

  function clearCustomerPreviewContent(sheetId, key) {
    const sheet = getSheetById(sheetId);
    if (!sheet || !isCoreKey(key)) return "";
    ensureCustomerPreviewShape(sheet);
    sheet.customerPreview.editedContents[key] = "";
    syncAutomaticCustomerPreviewVisibility(sheet);
    scheduleSave(sheet);
    return getCustomerPreviewContentFromSheet(sheet, key);
  }

  function toggleCustomerPreviewCategory(sheetId, key, visible) {
    const sheet = getSheetById(sheetId);
    if (!sheet || !isCoreKey(key)) return;
    ensureCustomerPreviewShape(sheet);
    const categories = new Set(sheet.customerPreview.visibleCategories);
    if (visible) categories.add(key);
    else categories.delete(key);
    sheet.customerPreview.visibleCategories = coreDefinitions.map(function (definition) { return definition.key; }).filter(function (item) { return categories.has(item); });
    sheet.customerPreview.visibilityTouched = true;
    scheduleSave(sheet);
  }

  function getCustomerPreviewContent(sheetId, key) {
    const sheet = getSheetById(sheetId);
    if (!sheet || !isCoreKey(key)) return "";
    ensureCustomerPreviewShape(sheet);
    return getCustomerPreviewContentFromSheet(sheet, key);
  }

  function getCustomerPreviewContentFromSheet(sheet, key) {
    if (!sheet || !isCoreKey(key)) return "";
    ensureCustomerPreviewShape(sheet);
    const edited = sheet.customerPreview.editedContents[key];
    if (hasText(edited)) return String(edited);
    return getCoreNoteContent(sheet, key);
  }

  function hasCustomerPreviewOverride(sheet, key) {
    if (!sheet || !isCoreKey(key)) return false;
    ensureCustomerPreviewShape(sheet);
    return hasText(sheet.customerPreview.editedContents[key]);
  }

  function getCoreNoteContent(sheet, key) {
    return sheet && sheet.coreNotes && sheet.coreNotes[key] ? String(sheet.coreNotes[key].content || "") : "";
  }

  function getAutomaticCustomerPreviewCategories(sheet) {
    return coreDefinitions
      .map(function (definition) { return definition.key; })
      .filter(function (key) { return hasText(getCustomerPreviewContentFromSheet(sheet, key)); });
  }

  function syncAutomaticCustomerPreviewVisibility(sheet) {
    ensureCustomerPreviewShape(sheet);
    if (sheet.customerPreview.visibilityTouched) return;
    const automaticCategories = getAutomaticCustomerPreviewCategories(sheet);
    sheet.customerPreview.visibleCategories = automaticCategories;
  }

  function ensureCustomerPreviewShape(sheet) {
    if (!sheet) return;
    const preview = sheet.customerPreview || {};
    const defaults = window.AppData.createCustomerPreview([]);
    const validImageIds = new Set((Array.isArray(sheet.attachments) ? sheet.attachments : []).map(function (attachment) { return attachment.imageId; }));
    sheet.customerPreview = {
      visibleCategories: Array.isArray(preview.visibleCategories) ? preview.visibleCategories.filter(isCoreKey) : [],
      editedContents: Object.assign(defaults.editedContents, preview.editedContents || {}),
      visibleImageIds: Array.isArray(preview.visibleImageIds)
        ? preview.visibleImageIds.filter(function (imageId) { return validImageIds.has(imageId); })
        : (Array.isArray(sheet.attachments) ? sheet.attachments : [])
          .filter(function (attachment) { return attachment.includeInCustomerOutput; })
          .map(function (attachment) { return attachment.imageId; }),
      visibilityTouched: preview.visibilityTouched === true
    };
  }

  function updateImageCaption(sheetId, imageId, caption) {
    const attachment = findAttachment(sheetId, imageId);
    if (!attachment) return;
    attachment.item.caption = caption;
    scheduleSave(attachment.sheet);
  }

  function toggleImageCustomerOutput(sheetId, imageId, enabled) {
    const attachment = findAttachment(sheetId, imageId);
    if (!attachment) return;
    attachment.item.includeInCustomerOutput = Boolean(enabled);
    const visible = new Set(attachment.sheet.customerPreview.visibleImageIds);
    if (enabled) visible.add(imageId);
    else visible.delete(imageId);
    attachment.sheet.customerPreview.visibleImageIds = Array.from(visible);
    scheduleSave(attachment.sheet);
  }

  function toggleCustomerPreviewImage(sheetId, imageId, visible) {
    const attachment = findAttachment(sheetId, imageId);
    if (!attachment || !attachment.item.includeInCustomerOutput) return;
    const visibleIds = new Set(attachment.sheet.customerPreview.visibleImageIds);
    if (visible) visibleIds.add(imageId);
    else visibleIds.delete(imageId);
    attachment.sheet.customerPreview.visibleImageIds = Array.from(visibleIds);
    scheduleSave(attachment.sheet);
  }

  function findAttachment(sheetId, imageId) {
    const sheet = getSheetById(sheetId);
    if (!sheet) return null;
    const item = sheet.attachments.find(function (attachment) { return attachment.imageId === imageId; });
    return item ? { sheet: sheet, item: item } : null;
  }

  async function addImagesToCategory(sheetId, categoryKey, files) {
    const sheet = getSheetById(sheetId);
    if (!sheet || !isCoreKey(categoryKey)) return { added: [], errors: ["画像を追加できる項目ではありません"] };
    const added = [];
    const errors = [];
    let nextOrder = sheet.attachments
      .filter(function (attachment) { return attachment.categoryKey === categoryKey; })
      .reduce(function (max, attachment) { return Math.max(max, attachment.order); }, 0) + 1;

    for (const file of Array.from(files || [])) {
      if (!ALLOWED_IMAGE_TYPES.has(file.type)) {
        errors.push(file.name + "：JPEG、PNG、WebPの画像を選択してください");
        continue;
      }
      if (file.size > MAX_IMAGE_SIZE) {
        errors.push(file.name + "：10MB以下の画像を選択してください");
        continue;
      }
      const imageId = generateUniqueId("IMG");
      const now = new Date().toISOString();
      try {
        await saveImageBlob({ imageId: imageId, sheetId: sheetId, categoryKey: categoryKey, blob: file, createdAt: now });
        const metadata = {
          imageId: imageId,
          categoryKey: categoryKey,
          fileName: file.name,
          mimeType: file.type,
          caption: "",
          includeInCustomerOutput: true,
          order: nextOrder,
          createdAt: now
        };
        nextOrder += 1;
        sheet.attachments.push(metadata);
        sheet.customerPreview.visibleImageIds.push(imageId);
        added.push(metadata);
      } catch (error) {
        console.error("画像を追加できませんでした。", error);
        errors.push(file.name + "：画像を追加できませんでした");
      }
    }
    if (added.length) {
      sheet.updatedAt = new Date().toISOString();
      saveAppData();
    }
    return { added: added, errors: errors };
  }

  async function deleteImageAttachment(sheetId, imageId) {
    const attachment = findAttachment(sheetId, imageId);
    if (!attachment) return;
    await deleteImageBlob(imageId);
    attachment.sheet.attachments = attachment.sheet.attachments.filter(function (item) { return item.imageId !== imageId; });
    attachment.sheet.customerPreview.visibleImageIds = attachment.sheet.customerPreview.visibleImageIds.filter(function (id) { return id !== imageId; });
    attachment.sheet.updatedAt = new Date().toISOString();
    saveAppData();
  }

  // ---------------------------------------------------------------------------
  // In-memory image repository. Images are available only while this page is open.
  // ---------------------------------------------------------------------------

  function openImageMemoryStore() {
    if (!imageStoreReadyPromise) imageStoreReadyPromise = Promise.resolve({ type: "memory" });
    return imageStoreReadyPromise;
  }

  async function saveImageBlob(imageData) {
    await openImageMemoryStore();
    imageMemoryStore.set(imageData.imageId, imageData);
    return imageData;
  }

  async function getImageBlob(imageId) {
    await openImageMemoryStore();
    return imageMemoryStore.get(imageId) || null;
  }

  async function getImagesBySheetId(sheetId) {
    await openImageMemoryStore();
    return Array.from(imageMemoryStore.values()).filter(function (item) { return item.sheetId === sheetId; });
  }

  async function deleteImageBlob(imageId) {
    await openImageMemoryStore();
    imageMemoryStore.delete(imageId);
  }

  async function deleteImagesBySheetId(sheetId) {
    await openImageMemoryStore();
    Array.from(imageMemoryStore.keys()).forEach(function (imageId) {
      const item = imageMemoryStore.get(imageId);
      if (item && item.sheetId === sheetId) imageMemoryStore.delete(imageId);
    });
  }

  async function clearImageMemoryStore() {
    imageMemoryStore.clear();
    imageStoreReadyPromise = null;
  }

  async function createImageObjectUrl(imageId) {
    const record = await getImageBlob(imageId);
    if (!record || !record.blob) return null;
    const url = URL.createObjectURL(record.blob);
    activeObjectUrls.add(url);
    return url;
  }

  function revokeImageObjectUrls() {
    activeObjectUrls.forEach(function (url) { URL.revokeObjectURL(url); });
    activeObjectUrls.clear();
    activeImageRenderPromises = [];
  }

  function findStakeholder(sheetId, stakeholderId) {
    const sheet = getSheetById(sheetId);
    if (!sheet) return null;
    const item = sheet.stakeholderNotes.find(function (stakeholder) { return stakeholder.id === stakeholderId; });
    return item ? { sheet: sheet, item: item } : null;
  }

  function createDefaultStakeholders() {
    return [
      { id: generateUniqueId("ST"), type: "default", key: "wife", label: "奥様", content: "", deletable: false },
      { id: generateUniqueId("ST"), type: "default", key: "husband", label: "ご主人様", content: "", deletable: false },
      { id: generateUniqueId("ST"), type: "default", key: "parents", label: "ご両親", content: "", deletable: false }
    ];
  }

  function generateUniqueId(prefix) {
    const ids = (prefix === "S"
      ? state.sheets.map(function (sheet) { return sheet.sheetId; })
      : prefix === "IMG"
        ? state.sheets.flatMap(function (sheet) { return sheet.attachments.map(function (item) { return item.imageId; }); })
        : state.sheets.flatMap(function (sheet) { return sheet.stakeholderNotes.map(function (item) { return item.id; }); }))
      .concat(Array.from(issuedIds).filter(function (id) { return id.startsWith(prefix); }));
    const max = ids.reduce(function (current, id) {
      const number = Number(String(id || "").replace(/\D/g, ""));
      return Number.isFinite(number) ? Math.max(current, number) : current;
    }, 0);
    const id = prefix + String(max + 1).padStart(4, "0");
    issuedIds.add(id);
    return id;
  }

  function isCoreKey(key) {
    return coreDefinitions.some(function (definition) { return definition.key === key; });
  }

  function hasText(value) {
    return String(value || "").trim() !== "";
  }

  function scheduleSave(sheet) {
    sheet.updatedAt = new Date().toISOString();
    setSaveStatus("メモリ保持中：CSV未書き出しの内容は再読み込みで消えます", false);
    window.clearTimeout(pendingSaveTimer);
    pendingSaveTimer = null;
  }

  function flushPendingSave() {
    if (!pendingSaveTimer) return;
    window.clearTimeout(pendingSaveTimer);
    pendingSaveTimer = null;
  }

  // Public functions make the data boundary easy to inspect and replace.
  window.AppStore = {
    loadAppData: loadAppData,
    saveAppData: saveAppData,
    getActiveSheet: getActiveSheet,
    getSheetById: getSheetById,
    createBlankSheet: createBlankSheet,
    updateSheet: updateSheet,
    deleteSheet: deleteSheet,
    setActiveSheet: setActiveSheet,
    clearLocalData: clearLocalData,
    updateCoreNote: updateCoreNote,
    updateCurrentStatus: updateCurrentStatus,
    updateStakeholderNote: updateStakeholderNote,
    updateStakeholderLabel: updateStakeholderLabel,
    addStakeholderNote: addStakeholderNote,
    deleteStakeholderNote: deleteStakeholderNote,
    initializeCustomerPreview: initializeCustomerPreview,
    updateCustomerPreviewContent: updateCustomerPreviewContent,
    clearCustomerPreviewContent: clearCustomerPreviewContent,
    toggleCustomerPreviewCategory: toggleCustomerPreviewCategory,
    getCustomerPreviewContent: getCustomerPreviewContent,
    getAutomaticCustomerPreviewCategories: getAutomaticCustomerPreviewCategories,
    createPractitionerCsv: createPractitionerCsv,
    createCsvFileName: createCsvFileName,
    importPractitionerCsvText: importPractitionerCsvText,
    importPractitionerCsvFile: importPractitionerCsvFile,
    exportPractitionerCsv: exportPractitionerCsv,
    openImageMemoryStore: openImageMemoryStore,
    saveImageBlob: saveImageBlob,
    getImageBlob: getImageBlob,
    getImagesBySheetId: getImagesBySheetId,
    deleteImageBlob: deleteImageBlob,
    deleteImagesBySheetId: deleteImagesBySheetId,
    clearImageMemoryStore: clearImageMemoryStore,
    addImagesToCategory: addImagesToCategory,
    updateImageCaption: updateImageCaption,
    toggleImageCustomerOutput: toggleImageCustomerOutput,
    toggleCustomerPreviewImage: toggleCustomerPreviewImage,
    deleteImageAttachment: deleteImageAttachment,
    createImageObjectUrl: createImageObjectUrl,
    revokeImageObjectUrls: revokeImageObjectUrls
  };

  // ---------------------------------------------------------------------------
  // Safe DOM and formatting helpers. User input is never inserted into innerHTML.
  // ---------------------------------------------------------------------------

  function el(tagName, options) {
    const element = document.createElement(tagName);
    const config = options || {};
    if (config.className) element.className = config.className;
    if (Object.prototype.hasOwnProperty.call(config, "text")) element.textContent = String(config.text);
    if (config.id) element.id = config.id;
    if (config.href) element.setAttribute("href", config.href);
    if (config.type) element.setAttribute("type", config.type);
    if (config.name) element.setAttribute("name", config.name);
    if (config.value !== undefined) element.value = config.value;
    if (config.checked !== undefined) element.checked = config.checked;
    if (config.placeholder) element.setAttribute("placeholder", config.placeholder);
    if (config.required) element.required = true;
    if (config.min !== undefined) element.min = config.min;
    if (config.attrs) Object.keys(config.attrs).forEach(function (name) { element.setAttribute(name, config.attrs[name]); });
    if (config.on) Object.keys(config.on).forEach(function (name) { element.addEventListener(name, config.on[name]); });
    for (let index = 2; index < arguments.length; index += 1) appendChild(element, arguments[index]);
    return element;
  }

  function appendChild(parent, child) {
    if (child === null || child === undefined || child === false) return;
    if (Array.isArray(child)) return child.forEach(function (item) { appendChild(parent, item); });
    parent.appendChild(child instanceof Node ? child : document.createTextNode(String(child)));
  }

  function todayIso() {
    const date = new Date();
    return date.getFullYear() + "-" + String(date.getMonth() + 1).padStart(2, "0") + "-" + String(date.getDate()).padStart(2, "0");
  }

  function formatDate(value) {
    if (!value) return "未設定";
    const date = new Date(String(value).slice(0, 10) + "T00:00:00");
    if (Number.isNaN(date.getTime())) return String(value);
    return new Intl.DateTimeFormat("ja-JP", { year: "numeric", month: "long", day: "numeric" }).format(date);
  }

  function formatDateTime(value) {
    if (!value) return "未更新";
    const date = new Date(value);
    if (Number.isNaN(date.getTime())) return String(value);
    return new Intl.DateTimeFormat("ja-JP", {
      year: "numeric", month: "numeric", day: "numeric", hour: "2-digit", minute: "2-digit"
    }).format(date);
  }

  function setSaveStatus(message, saved) {
    if (!saveStatusNode) return;
    saveStatusNode.textContent = message;
    saveStatusNode.classList.toggle("is-saved", Boolean(saved));
  }

  function showToast(message, isError) {
    toast.textContent = message;
    toast.classList.toggle("is-error", Boolean(isError));
    toast.classList.add("is-visible");
    window.clearTimeout(showToast.timer);
    showToast.timer = window.setTimeout(function () { toast.classList.remove("is-visible"); }, 2600);
  }

  function createStorageNotice() {
    return el("aside", { className: "panel storage-notice no-print" },
      el("strong", { text: "このシートの内容はブラウザに保存されません。" }),
      el("p", { text: "後日編集する場合は、実務者用CSVを書き出して保管してください。" })
    );
  }

  function createCsvSecurityNotice() {
    return el("p", { className: "csv-security-note no-print", text: "CSVにはお客様情報が含まれます。保管・共有先にご注意ください。" });
  }

  function getStakeholderContent(sheet, key) {
    const stakeholder = sheet.stakeholderNotes.find(function (item) { return item.key === key; });
    return stakeholder ? stakeholder.content : "";
  }

  function getCsvValue(sheet, key) {
    if (key === "sheetId") return sheet.sheetId;
    if (key === "customerName") return sheet.customerName;
    if (key === "meetingDate") return sheet.meetingDate;
    if (key === "meetingNumber") return String(sheet.meetingNumber || 1);
    if (key === "title") return sheet.title;
    if (CURRENT_STATUS_FIELDS.some(function (field) { return field.key === key; })) {
      const value = sheet.currentStatus ? sheet.currentStatus[key] : "";
      return MONEY_STATUS_KEYS.has(key) ? normalizeMoneyInput(value) : String(value || "");
    }
    if (isCoreKey(key)) return getCoreNoteContent(sheet, key);
    if (["wife", "husband", "parents"].includes(key)) return getStakeholderContent(sheet, key);
    return "";
  }

  function escapeCsvValue(value) {
    return "\"" + String(value === undefined || value === null ? "" : value).replace(/"/g, "\"\"") + "\"";
  }

  function createPractitionerCsv(sheet) {
    const header = CSV_COLUMNS.join(",");
    const values = CSV_COLUMNS.map(function (key) { return escapeCsvValue(getCsvValue(sheet, key)); }).join(",");
    return "\uFEFF" + header + "\r\n" + values + "\r\n";
  }

  function safeFileNamePart(value, fallback) {
    const text = String(value || "").trim() || fallback;
    return text.replace(/[\\/:*?"<>|]/g, "_").replace(/\s+/g, "_");
  }

  function createCsvFileName(sheet) {
    return "ベクトルシート_" + safeFileNamePart(sheet.customerName, "お客様名未設定") + "_" + safeFileNamePart(sheet.meetingDate, todayIso()) + ".csv";
  }

  function downloadTextFile(fileName, text, type) {
    const blob = new Blob([text], { type: type || "text/plain;charset=utf-8" });
    const url = URL.createObjectURL(blob);
    activeObjectUrls.add(url);
    const link = el("a", { href: url, attrs: { download: fileName } });
    document.body.appendChild(link);
    link.click();
    link.remove();
  }

  function exportPractitionerCsv(sheetId) {
    flushPendingSave();
    const sheet = getSheetById(sheetId);
    if (!sheet) return;
    const csv = createPractitionerCsv(sheet);
    downloadTextFile(createCsvFileName(sheet), csv, "text/csv;charset=utf-8");
    showToast("実務者用CSVを書き出しました。");
  }

  function parseCsv(text) {
    const rows = [];
    let row = [];
    let value = "";
    let quoted = false;
    const source = String(text || "").replace(/^\uFEFF/, "");
    for (let index = 0; index < source.length; index += 1) {
      const char = source[index];
      const next = source[index + 1];
      if (quoted) {
        if (char === "\"" && next === "\"") {
          value += "\"";
          index += 1;
        } else if (char === "\"") {
          quoted = false;
        } else {
          value += char;
        }
      } else if (char === "\"") {
        quoted = true;
      } else if (char === ",") {
        row.push(value);
        value = "";
      } else if (char === "\n") {
        row.push(value);
        rows.push(row);
        row = [];
        value = "";
      } else if (char !== "\r") {
        value += char;
      }
    }
    if (value !== "" || row.length) {
      row.push(value);
      rows.push(row);
    }
    return rows;
  }

  function createSheetFromCsvRecord(record) {
    const now = new Date().toISOString();
    const sheet = {
      sheetId: String(record.sheetId || generateUniqueId("S")),
      customerName: String(record.customerName || ""),
      meetingDate: String(record.meetingDate || todayIso()),
      meetingNumber: Math.max(1, Number(record.meetingNumber || 1)),
      title: String(record.title || ""),
      currentStatus: {
        currentResidence: String(record.currentResidence || ""),
        currentRent: normalizeMoneyInput(record.currentRent || ""),
        ownFunds: normalizeMoneyInput(record.ownFunds || ""),
        annualIncome: normalizeMoneyInput(record.annualIncome || ""),
        desiredArea: String(record.desiredArea || "")
      },
      coreNotes: window.AppData.createCoreNotes({
        scheduleReason: record.scheduleReason || "",
        budget: record.budget || "",
        building: record.building || "",
        land: record.land || ""
      }),
      stakeholderNotes: createDefaultStakeholders(),
      attachments: [],
      customerPreview: window.AppData.createCustomerPreview([]),
      createdAt: now,
      updatedAt: now
    };
    ["wife", "husband", "parents"].forEach(function (key) {
      const stakeholder = sheet.stakeholderNotes.find(function (item) { return item.key === key; });
      if (stakeholder) stakeholder.content = String(record[key] || "");
    });
    return normalizeSheet(sheet);
  }

  function importPractitionerCsvText(text) {
    const rows = parseCsv(text).filter(function (row) {
      return row.some(function (value) { return String(value || "").trim() !== ""; });
    });
    if (rows.length < 2) throw new Error("CSVにデータ行がありません。");
    const headers = rows[0].map(function (header) { return String(header || "").trim(); });
    const values = rows[1] || [];
    const record = {};
    headers.forEach(function (header, index) {
      if (header) record[header] = values[index] || "";
    });
    const sheet = createSheetFromCsvRecord(record);
    state.sheets = [sheet];
    state.activeSheetId = sheet.sheetId;
    issuedIds.add(sheet.sheetId);
    ensureCustomerPreviewShape(sheet);
    syncAutomaticCustomerPreviewVisibility(sheet);
    return sheet;
  }

  async function importPractitionerCsvFile(file) {
    if (!file) return null;
    const text = await file.text();
    const sheet = importPractitionerCsvText(text);
    showToast("実務者用CSVを読み込みました。");
    navigateToSheet(sheet.sheetId, "input");
    return sheet;
  }

  function createCsvImportControl(label) {
    const inputId = "csv-import-" + Math.random().toString(36).slice(2, 8);
    const fileInput = el("input", {
      id: inputId,
      type: "file",
      className: "csv-file-input",
      attrs: { accept: ".csv,text/csv" }
    });
    const button = el("button", { className: "button button-secondary", type: "button", text: label || "実務者用CSVを読み込む" });
    button.addEventListener("click", function () { fileInput.click(); });
    fileInput.addEventListener("change", async function () {
      const file = fileInput.files && fileInput.files[0];
      fileInput.value = "";
      if (!file) return;
      button.disabled = true;
      try {
        await importPractitionerCsvFile(file);
      } catch (error) {
        console.error("実務者用CSVを読み込めませんでした。", error);
        showToast("実務者用CSVを読み込めませんでした。形式を確認してください。", true);
      } finally {
        button.disabled = false;
      }
    });
    return el("span", { className: "csv-import-control" }, fileInput, button);
  }

  function createCsvExportButton(sheetId, small) {
    const button = el("button", {
      className: "button button-secondary" + (small ? " button-small" : ""),
      type: "button",
      text: "実務者用CSVを書き出す"
    });
    button.addEventListener("click", function () { exportPractitionerCsv(sheetId); });
    return button;
  }

  function updateBackToTopVisibility() {
    if (!backToTopButton) return;
    backToTopButton.classList.toggle("is-visible", window.scrollY > 400);
  }

  function scrollToPageTop() {
    window.scrollTo({ top: 0, behavior: "smooth" });
  }

  function navigateToStart() {
    currentView = { name: "start", sheetId: null, mode: "input" };
    renderApp();
  }

  function navigateToNewSheet() {
    currentView = { name: "new", sheetId: null, mode: "input" };
    renderApp();
  }

  function navigateToSheet(sheetId, mode) {
    currentView = { name: "sheet", sheetId: sheetId, mode: mode || "input" };
    renderApp();
  }

  function createNavButton(className, text, onClick) {
    const button = el("button", { className: className, type: "button", text: text });
    button.addEventListener("click", onClick);
    return button;
  }

  function confirmAction(message) {
    return window.confirm(message);
  }

  async function printCustomerPreview() {
    flushPendingSave();
    await Promise.allSettled(activeImageRenderPromises);
    window.print();
  }

  function setPageTitle(title) {
    document.title = title + "｜ベクトルシート";
  }

  function makeField(labelText, name, type, value, options) {
    const config = options || {};
    const id = config.id || "field-" + name + "-" + Math.random().toString(36).slice(2, 7);
    const input = el("input", {
      id: id, name: name, type: type, value: value === undefined ? "" : value,
      placeholder: config.placeholder || "", required: config.required, min: config.min,
      attrs: config.attrs || null
    });
    return {
      input: input,
      wrapper: el("div", { className: "field" }, el("label", { text: labelText, attrs: { for: id } }), input)
    };
  }

  function appendFieldError(field, message) {
    const error = el("p", { className: "field-error", text: message, attrs: { role: "alert", hidden: "" } });
    appendChild(field.wrapper, error);
    return error;
  }

  function createPageHeading(title, description, action) {
    return el("div", { className: "page-heading" },
      el("div", {}, el("p", { className: "eyebrow", text: "VECTOR SHEET" }), el("h1", { text: title }), description ? el("p", { text: description }) : null),
      action || null
    );
  }

  function modeLabel(mode) {
    return mode === "input" ? "入力・社内確認モード" : mode === "internal" ? "入力・社内確認モード" : "お客様用確認・PDF出力モード";
  }

  // ---------------------------------------------------------------------------
  // Sheet list and blank sheet creation
  // ---------------------------------------------------------------------------

  function renderSheetList() {
    revokeImageObjectUrls();
    saveStatusNode = null;
    setPageTitle("開始");
    const page = el("section", { className: "page" });
    appendChild(page, createPageHeading("ベクトルシート", "ブラウザには保存せず、必要な引き継ぎは実務者用CSVで行います。"));

    appendChild(page, el("div", { className: "panel empty-state empty-state-welcome" },
      el("p", { className: "empty-state-kicker", text: "START A MEETING NOTE" }),
      el("h2", { text: "新規入力、または実務者用CSVから開始します" }),
      el("p", { text: "このシートの内容はブラウザに保存されません。後日編集する場合は、実務者用CSVを書き出して保管してください。" }),
      el("div", { className: "start-actions" },
        createNavButton("button empty-state-action", "新規入力", navigateToNewSheet),
        createCsvImportControl("実務者用CSVを読み込む")
      ),
      createCsvSecurityNotice(),
      el("ol", { className: "empty-state-flow", attrs: { "aria-label": "運用の流れ" } },
        el("li", {}, el("span", { text: "01" }), el("strong", { text: "入力・確認" })),
        el("li", {}, el("span", { text: "02" }), el("strong", { text: "PDF出力" })),
        el("li", {}, el("span", { text: "03" }), el("strong", { text: "CSV保管" }))
      )
    ));
    app.replaceChildren(page);
  }

  function sheetDatum(label, value) {
    return el("div", { className: "sheet-datum" }, el("small", { text: label }), el("span", { text: value }));
  }

  function renderNewSheetForm() {
    revokeImageObjectUrls();
    saveStatusNode = null;
    setPageTitle("新規入力");
    const page = el("section", { className: "page page-narrow" });
    appendChild(page, createNavButton("back-link nav-text-button", "← 開始画面へ戻る", navigateToStart));
    appendChild(page, createPageHeading("新規入力", "基本情報を入力すると、項目が空欄のシートを作成します。"));
    appendChild(page, createStorageNotice());
    const form = el("form", { className: "panel form-panel", attrs: { novalidate: "" } });
    const customer = makeField("お施主様名", "customerName", "text", "", { required: true, placeholder: "お施主様名を入力" });
    const date = makeField("商談日", "meetingDate", "date", todayIso(), { required: true });
    const number = makeField("商談回数", "meetingNumber", "number", 1, { required: true, min: 1 });
    const title = makeField("商談タイトル（任意）", "title", "text", "", { placeholder: "今回の打ち合わせ内容を入力" });
    const customerError = appendFieldError(customer, "お施主様名を入力してください。");
    const dateError = appendFieldError(date, "商談日を入力してください。");
    const numberError = appendFieldError(number, "商談回数は1以上で入力してください。");
    appendChild(form, el("div", { className: "field-grid" }, customer.wrapper, date.wrapper, number.wrapper, title.wrapper));
    appendChild(form, el("div", { className: "form-actions" },
      createNavButton("button button-secondary", "開始画面へ戻る", navigateToStart),
      el("button", { className: "button", type: "submit", text: "空の商談シートを作成" })
    ));
    form.addEventListener("submit", function (event) {
      event.preventDefault();
      const validCustomer = Boolean(customer.input.value.trim());
      const validDate = Boolean(date.input.value);
      const validNumber = Number(number.input.value) >= 1;
      customerError.hidden = validCustomer;
      dateError.hidden = validDate;
      numberError.hidden = validNumber;
      if (!validCustomer || !validDate || !validNumber) {
        const firstInvalid = !validCustomer ? customer.input : !validDate ? date.input : number.input;
        firstInvalid.focus();
        return;
      }
      const sheet = createBlankSheet({
        customerName: customer.input.value.trim(),
        meetingDate: date.input.value,
        meetingNumber: Number(number.input.value),
        title: title.input.value.trim()
      });
      showToast("新しい商談シートを作成しました。");
      navigateToSheet(sheet.sheetId, "input");
    });
    appendChild(page, form);
    app.replaceChildren(page);
    customer.input.focus();
  }

  // ---------------------------------------------------------------------------
  // Shared sheet header and input mode
  // ---------------------------------------------------------------------------

  function createWorkspaceHeader(sheet, mode) {
    saveStatusNode = el("span", { className: "save-status", text: "ブラウザ保存なし：CSVを書き出して保管してください" });
    const actions = el("div", { className: "workspace-actions no-print" });
    if (mode === "input") appendChild(actions, createCsvImportControl("CSV読み込み"));
    if (mode === "input" || mode === "internal") appendChild(actions, createCsvExportButton(sheet.sheetId, true));
    appendChild(actions, modeButton(sheet.sheetId, "input", "入力・社内確認", mode === "internal" ? "input" : mode));
    appendChild(actions, modeButton(sheet.sheetId, "customer", "お客様用で確認・出力", mode));
    appendChild(actions, createNavButton("button button-quiet button-small", "開始画面へ戻る", navigateToStart));

    return el("header", { className: "workspace-header" },
      el("div", { className: "workspace-title" },
        el("div", { className: "mode-line" },
          el("span", { className: "mode-badge", text: modeLabel(mode) })
        ),
        el("h1", { text: sheet.customerName + "　第" + sheet.meetingNumber + "回" }),
        el("p", { text: formatDate(sheet.meetingDate) + "　｜　" + (sheet.title || "タイトル未設定") }),
        saveStatusNode
      ),
      actions
    );
  }

  function modeButton(sheetId, targetMode, label, currentMode) {
    const button = el("button", {
      className: "button button-small " + (targetMode === currentMode ? "button-active" : "button-secondary") + (targetMode === "customer" ? " button-output" : ""),
      type: "button",
      text: label,
      attrs: { "aria-current": targetMode === currentMode ? "page" : "false" }
    });
    button.disabled = targetMode === currentMode;
    button.addEventListener("click", function () {
      flushPendingSave();
      navigateToSheet(sheetId, targetMode);
    });
    return button;
  }

  function createCurrentStatusEditor(sheetId, sheet) {
    sheet.currentStatus = normalizeCurrentStatus(sheet.currentStatus);
    const panel = el("section", { className: "panel form-panel current-status-editor" },
      el("div", { className: "section-heading" },
        el("div", {}, el("p", { className: "section-number", text: "02" }), el("h2", { text: "現在のご状況" })),
        el("span", { className: "field-hint", text: "お客様用の振り返りシートに表示する基本状況です。未入力の項目は出力されません。" })
      )
    );
    const fields = CURRENT_STATUS_FIELDS.map(function (field) {
      const formField = makeField(field.label, "currentStatus-" + field.key, "text", sheet.currentStatus[field.key], field.money ? {
        attrs: { inputmode: "decimal", pattern: "[0-9.]*" }
      } : {});
      if (field.money) {
        formField.wrapper.classList.add("field-with-unit");
        formField.input.addEventListener("blur", function () {
          formField.input.value = normalizeMoneyInput(formField.input.value);
          updateCurrentStatus(sheetId, field.key, formField.input.value);
        });
        appendChild(formField.wrapper, el("span", { className: "field-unit", text: field.unit }));
      }
      formField.input.addEventListener("input", function () {
        updateCurrentStatus(sheetId, field.key, formField.input.value);
      });
      return formField.wrapper;
    });
    appendChild(panel, el("div", { className: "field-grid current-status-grid" }, fields));
    return panel;
  }

  function renderInputMode(sheetId) {
    revokeImageObjectUrls();
    const sheet = getSheetById(sheetId);
    if (!sheet) return renderNotFound();
    state.activeSheetId = sheetId;
    setPageTitle(sheet.customerName + "・入力モード");
    const page = el("section", { className: "page workspace-page" });
    appendChild(page, createWorkspaceHeader(sheet, "input"));
    appendChild(page, createStorageNotice());
    appendChild(page, el("section", { className: "panel csv-panel no-print" },
      el("div", {}, el("h2", { text: "実務者用CSV" }), createCsvSecurityNotice()),
      el("div", { className: "csv-actions" },
        createCsvImportControl("実務者用CSVを読み込む"),
        createCsvExportButton(sheetId)
      )
    ));

    const basicPanel = el("section", { className: "panel form-panel" },
      el("div", { className: "section-heading" },
        el("div", {}, el("p", { className: "section-number", text: "01" }), el("h2", { text: "商談基本情報" })),
        el("span", { className: "field-hint", text: "変更はこのシートだけに反映されます" })
      )
    );
    const customer = makeField("お施主様名", "customerName", "text", sheet.customerName, { required: true });
    const date = makeField("商談日", "meetingDate", "date", sheet.meetingDate, { required: true });
    const number = makeField("商談回数", "meetingNumber", "number", sheet.meetingNumber, { required: true, min: 1 });
    const title = makeField("商談タイトル（任意）", "title", "text", sheet.title, { placeholder: "今回の打ち合わせ内容を入力" });
    [
      [customer.input, "customerName", function (value) { return value; }],
      [date.input, "meetingDate", function (value) { return value; }],
      [number.input, "meetingNumber", function (value) { return Math.max(1, Number(value || 1)); }],
      [title.input, "title", function (value) { return value; }]
    ].forEach(function (entry) {
      entry[0].addEventListener("input", function () { updateSheet(sheetId, { [entry[1]]: entry[2](entry[0].value) }); });
    });
    appendChild(basicPanel, el("div", { className: "field-grid four-fields" }, customer.wrapper, date.wrapper, number.wrapper, title.wrapper));
    appendChild(page, basicPanel);
    appendChild(page, createCurrentStatusEditor(sheetId, sheet));

    appendChild(page, el("div", { className: "input-group-heading" },
      el("div", {}, el("p", { className: "section-number", text: "03–06" }), el("h2", { text: "家づくりに関する主要4項目" })),
      el("p", { text: "入力した文章は、このページを開いている間だけ保持されます。後日編集する場合はCSVを書き出してください。" })
    ));
    coreDefinitions.forEach(function (definition, index) {
      const editor = createNoteEditor({
        index: index + 3,
        label: definition.label,
        hint: definition.hint,
        content: sheet.coreNotes[definition.key].content,
        onInput: function (content) { updateCoreNote(sheetId, definition.key, content); }
      });
      appendChild(editor, createImageManager(sheet, definition));
      appendChild(page, editor);
    });

    appendChild(page, el("div", { className: "input-group-heading stakeholder-heading" },
      el("div", {}, el("p", { className: "section-number", text: "07–" }), el("h2", { text: "関係者ごとの補足" })),
      el("p", { text: "この内容は社内用表示だけに含まれ、お客様用資料には表示されません。" })
    ));
    sheet.stakeholderNotes.forEach(function (stakeholder, index) {
      appendChild(page, createStakeholderEditor(sheet, stakeholder, index + 7));
    });
    appendChild(page, createStakeholderAddForm(sheetId));
    appendChild(page, el("div", { className: "bottom-actions no-print" },
      createCsvExportButton(sheetId),
      modeButton(sheetId, "customer", "お客様用で確認・出力", "input"),
      createNavButton("button button-secondary", "開始画面へ戻る", navigateToStart)
    ));
    app.replaceChildren(page);
  }

  function createNoteEditor(config) {
    const textareaId = "note-" + String(config.index) + "-" + Math.random().toString(36).slice(2, 6);
    const textarea = el("textarea", { id: textareaId, value: config.content, attrs: { rows: "6" }, placeholder: config.hint });
    const count = el("span", { className: "character-count", text: config.content.length + "文字" });
    const status = el("span", { className: "entry-status" + (config.content.trim() ? " is-filled" : ""), text: config.content.trim() ? "入力済み" : "未入力" });
    const clear = el("button", { className: "text-button", type: "button", text: "内容をクリア" });
    function updateVisuals() {
      count.textContent = textarea.value.length + "文字";
      const filled = Boolean(textarea.value.trim());
      status.textContent = filled ? "入力済み" : "未入力";
      status.classList.toggle("is-filled", filled);
    }
    textarea.addEventListener("input", function () {
      updateVisuals();
      config.onInput(textarea.value);
    });
    clear.addEventListener("click", function () {
      if (textarea.value && !confirmAction(config.label + "の内容を空欄にしますか？")) return;
      textarea.value = "";
      updateVisuals();
      config.onInput("");
      textarea.focus();
    });
    return el("section", { className: "panel note-editor" },
      el("div", { className: "note-editor-head" },
        el("div", {}, el("p", { className: "section-number", text: String(config.index).padStart(2, "0") }), el("h2", { text: config.label }), el("p", { text: config.hint })),
        status
      ),
      el("label", { className: "sr-only", text: config.label + "の内容", attrs: { for: textareaId } }),
      textarea,
      el("div", { className: "note-editor-footer" }, count, clear)
    );
  }

  function createImageManager(sheet, definition) {
    const inputId = "image-upload-" + definition.key;
    const fileInput = el("input", {
      id: inputId,
      type: "file",
      attrs: { accept: "image/jpeg,image/png,image/webp", multiple: "" }
    });
    const uploadLabel = el("label", {
      className: "button button-secondary image-upload-button",
      text: definition.key === "building" ? "＋ 建物の参考画像を追加" : "＋ 画像を追加",
      attrs: { for: inputId }
    });
    const message = el("p", { className: "image-upload-hint", text: "JPEG・PNG・WebP／1ファイル10MB以下／複数選択可" });
    const gallery = el("div", { className: "image-gallery image-gallery-input" });
    const renderPromise = renderCategoryImages(sheet.sheetId, definition.key, "input", gallery);
    activeImageRenderPromises.push(renderPromise);

    fileInput.addEventListener("change", async function () {
      if (!fileInput.files || !fileInput.files.length) return;
      uploadLabel.classList.add("is-busy");
      uploadLabel.textContent = "画像を読み込み中…";
      const result = await addImagesToCategory(sheet.sheetId, definition.key, fileInput.files);
      if (result.added.length) showToast(result.added.length + "件の画像を追加しました。");
      if (result.errors.length) showToast(result.errors.join("　"), true);
      fileInput.value = "";
      renderInputMode(sheet.sheetId);
    });

    return el("div", { className: "image-manager" + (definition.key === "building" ? " is-building" : "") },
      el("div", { className: "image-manager-head" },
        el("div", {}, el("h3", { text: "添付画像" }), message),
        el("div", {}, fileInput, uploadLabel)
      ),
      gallery
    );
  }

  async function renderCategoryImages(sheetId, categoryKey, mode, container) {
    const sheet = getSheetById(sheetId);
    if (!sheet) return;
    let attachments = sheet.attachments
      .filter(function (attachment) { return attachment.categoryKey === categoryKey; })
      .sort(function (a, b) { return a.order - b.order; });
    if (mode === "customer") {
      attachments = attachments.filter(function (attachment) {
        return attachment.includeInCustomerOutput;
      });
    }
    if (!attachments.length) return;

    const figures = await Promise.all(attachments.map(function (attachment) {
      return createImageFigure(sheet, attachment, mode);
    }));
    figures.forEach(function (figure) { appendChild(container, figure); });
  }

  function getCustomerOutputAttachments(sheet) {
    if (!sheet) return [];
    const categoryOrder = new Map(coreDefinitions.map(function (definition, index) { return [definition.key, index]; }));
    return sheet.attachments
      .filter(function (attachment) {
        return attachment.includeInCustomerOutput && isCoreKey(attachment.categoryKey);
      })
      .sort(function (a, b) {
        const categoryDiff = (categoryOrder.get(a.categoryKey) || 0) - (categoryOrder.get(b.categoryKey) || 0);
        return categoryDiff || a.order - b.order;
      });
  }

  function getCustomerImageGroupTitle(categoryKey) {
    return getCustomerOutputLabel(categoryKey).title;
  }

  function getCustomerOutputLabel(key) {
    return CUSTOMER_OUTPUT_LABELS[key] || {
      title: "イメージ",
      subtitle: "",
      point: ""
    };
  }

  function getCustomerSectionNumber(key) {
    const index = CUSTOMER_OUTPUT_ORDER.indexOf(key);
    return index >= 0 ? String(index + 1).padStart(2, "0") : "";
  }

  function getCustomerSectionAreaClass(key) {
    if (key === "scheduleReason") return "reason-section";
    if (key === "budget") return "budget-section";
    if (key === "building") return "building-section lifestyle-section";
    if (key === "land") return "land-section";
    return "";
  }

  function getCurrentStatusRows(sheet) {
    const status = normalizeCurrentStatus(sheet && sheet.currentStatus);
    return CURRENT_STATUS_FIELDS
      .map(function (field) {
        return {
          key: field.key,
          label: field.label,
          value: formatCurrentStatusValue(field.key, status[field.key])
        };
      })
      .filter(function (row) { return row.value !== ""; });
  }

  function createCustomerCurrentStatusArea(sheet) {
    const rows = getCurrentStatusRows(sheet);
    if (!rows.length) return null;
    return el("section", { className: "current-status-area", attrs: { "aria-label": "現在のご状況" } },
      el("div", { className: "current-status-table-wrap" },
        el("table", { className: "current-status-table" },
          el("tbody", {},
            rows.map(function (row) {
              return el("tr", {},
                el("th", { text: row.label }),
                el("td", { text: row.value })
              );
            })
          )
        )
      )
    );
  }

  function getCustomerTopPhotoAttachment(sheet) {
    const attachments = getCustomerOutputAttachments(sheet);
    if (!attachments.length) return null;
    return attachments.find(function (attachment) { return attachment.categoryKey === "building"; }) || attachments[0];
  }

  function createCustomerTopPhotoArea(sheet) {
    const attachment = getCustomerTopPhotoAttachment(sheet);
    if (!attachment) return null;
    const area = el("section", { className: "top-photo-area", attrs: { "aria-label": "参考写真" } },
      el("div", { className: "top-photo-frame" },
        el("span", { className: "top-photo-loading", text: "参考写真を読み込み中" })
      )
    );
    activeImageRenderPromises.push(renderCustomerTopPhoto(sheet, attachment, area));
    return area;
  }

  async function renderCustomerTopPhoto(sheet, attachment, area) {
    const frame = area.querySelector(".top-photo-frame");
    if (!frame) return;
    try {
      const objectUrl = await createImageObjectUrl(attachment.imageId);
      if (!objectUrl) {
        area.hidden = true;
        return;
      }
      const image = el("img", {
        className: "top-reference-photo",
        attrs: { src: objectUrl, alt: attachment.caption || attachment.fileName || "参考写真" }
      });
      frame.replaceChildren(image);
      if (typeof image.decode === "function") await image.decode().catch(function () {});
    } catch (error) {
      console.error("参考写真を表示できませんでした。", error);
      area.hidden = true;
    }
  }

  function createCustomerTopInfoRow(sheet, currentStatusArea) {
    const photoArea = createCustomerTopPhotoArea(sheet);
    if (!photoArea && !currentStatusArea) return null;
    return el("div", { className: "top-info-row" },
      photoArea || el("div", { className: "top-photo-area top-photo-area-empty", attrs: { "aria-hidden": "true" } }),
      currentStatusArea || el("div", { className: "current-status-area current-status-area-empty", attrs: { "aria-hidden": "true" } })
    );
  }

  function createCustomerGridSectionHeader(key) {
    const label = getCustomerOutputLabel(key);
    return el("header", { className: "customer-card-header" },
      el("span", { className: "customer-section-number", text: getCustomerSectionNumber(key) }),
      el("h2", { className: "customer-section-title", text: label.title })
    );
  }

  function createCustomerSectionHeader(key) {
    const label = getCustomerOutputLabel(key);
    return el("header", { className: "customer-section-header" },
      el("h2", { className: "customer-section-title", text: label.title }),
      el("p", { className: "customer-section-subtitle", text: label.subtitle }),
      el("p", { className: "customer-section-point", text: label.point })
    );
  }

  function createCustomerSectionIndex(key) {
    return el("div", { className: "customer-section-index", text: getCustomerSectionNumber(key) });
  }

  function groupCustomerOutputAttachments(sheet) {
    const groups = new Map();
    getCustomerOutputAttachments(sheet).forEach(function (attachment) {
      if (!groups.has(attachment.categoryKey)) groups.set(attachment.categoryKey, []);
      groups.get(attachment.categoryKey).push(attachment);
    });
    return coreDefinitions
      .map(function (definition) {
        return {
          key: definition.key,
          title: getCustomerImageGroupTitle(definition.key),
          attachments: groups.get(definition.key) || []
        };
      })
      .filter(function (group) { return group.attachments.length; });
  }

  async function renderCustomerImageSection(sheetId, section) {
    const sheet = getSheetById(sheetId);
    if (!sheet) return;
    const groups = groupCustomerOutputAttachments(sheet);
    for (const group of groups) {
      const grid = el("div", {
        className: "customer-image-grid",
        attrs: { "data-count": String(group.attachments.length) }
      });
      const figures = await Promise.all(group.attachments.map(function (attachment) {
        return createImageFigure(sheet, attachment, "customer");
      }));
      figures.forEach(function (figure) { appendChild(grid, figure); });
      appendChild(section, el("section", { className: "customer-image-group" },
        el("h3", { className: "customer-image-group-title", text: group.title }),
        grid
      ));
    }
  }

  async function createImageFigure(sheet, attachment, mode) {
    const figure = el("figure", {
      className: (mode === "customer" ? "customer-image-figure" : "customer-image-figure image-figure") + " mode-" + mode,
      attrs: { "data-image-id": attachment.imageId }
    });
    const frame = mode === "customer" ? el("div", { className: "customer-image-frame" }) : null;
    try {
      const objectUrl = await createImageObjectUrl(attachment.imageId);
      if (objectUrl) {
        const image = el("img", {
          className: "customer-image",
          attrs: { src: objectUrl, alt: attachment.caption || attachment.fileName }
        });
        appendChild(frame || figure, image);
        if (typeof image.decode === "function") await image.decode().catch(function () {});
      } else {
        appendChild(frame || figure, el("div", { className: "image-load-error", text: "画像データが見つかりません" }));
      }
    } catch (error) {
      console.error("画像を表示できませんでした。", error);
      appendChild(frame || figure, el("div", { className: "image-load-error", text: "画像を読み込めませんでした" }));
    }
    if (mode === "customer") {
      appendChild(figure, frame);
      appendChild(figure, el("figcaption", {
        className: "customer-image-caption",
        text: attachment.caption || attachment.fileName
      }));
      return figure;
    }

    if (mode === "input") {
      const captionId = "caption-" + attachment.imageId;
      const caption = el("input", {
        id: captionId,
        type: "text",
        value: attachment.caption,
        placeholder: "任意のキャプション"
      });
      caption.addEventListener("input", function () { updateImageCaption(sheet.sheetId, attachment.imageId, caption.value); });
      const includeId = "include-image-" + attachment.imageId;
      const include = el("input", { id: includeId, type: "checkbox", checked: attachment.includeInCustomerOutput });
      include.addEventListener("change", function () {
        toggleImageCustomerOutput(sheet.sheetId, attachment.imageId, include.checked);
      });
      const deleteButton = el("button", { className: "button button-danger-quiet button-small", type: "button", text: "画像を削除" });
      deleteButton.addEventListener("click", async function () {
        if (!confirmAction("画像「" + (attachment.caption || attachment.fileName) + "」を削除しますか？")) return;
        deleteButton.disabled = true;
        try {
          await deleteImageAttachment(sheet.sheetId, attachment.imageId);
          showToast("画像を削除しました。");
          renderInputMode(sheet.sheetId);
        } catch (error) {
          console.error("画像を削除できませんでした。", error);
          deleteButton.disabled = false;
          showToast("画像を削除できませんでした。もう一度お試しください。", true);
        }
      });
      appendChild(figure, el("figcaption", { className: "image-figcaption" },
        el("p", { className: "image-file-name", text: attachment.fileName }),
        el("div", { className: "field" }, el("label", { text: "キャプション", attrs: { for: captionId } }), caption),
        el("div", { className: "image-figure-actions" },
          el("label", { className: "preview-choice", attrs: { for: includeId } }, include, el("span", { text: "お客様用資料にも表示する" })),
          deleteButton
        )
      ));
    } else if (mode === "internal") {
      appendChild(figure, el("figcaption", { className: "image-figcaption" },
        el("p", { text: attachment.caption || attachment.fileName }),
        el("span", {
          className: "image-output-state" + (attachment.includeInCustomerOutput ? " is-enabled" : ""),
          text: attachment.includeInCustomerOutput ? "お客様用：表示" : "お客様用：非表示"
        })
      ));
    }
    return figure;
  }

  function createStakeholderEditor(sheet, stakeholder, index) {
    const editor = createNoteEditor({
      index: index,
      label: stakeholder.label,
      hint: "ご希望や気になっていることを入力",
      content: stakeholder.content,
      onInput: function (content) { updateStakeholderNote(sheet.sheetId, stakeholder.id, content); }
    });
    if (!stakeholder.deletable) return editor;

    const heading = editor.querySelector(".note-editor-head > div");
    const headingElement = heading.querySelector("h2");
    const labelId = "stakeholder-label-" + stakeholder.id;
    const labelInput = el("input", { id: labelId, className: "stakeholder-label-input", type: "text", value: stakeholder.label, attrs: { "aria-label": "補足項目名" } });
    headingElement.replaceWith(labelInput);
    labelInput.addEventListener("input", function () {
      updateStakeholderLabel(sheet.sheetId, stakeholder.id, labelInput.value);
    });
    const deleteButton = el("button", { className: "button button-danger-quiet button-small", type: "button", text: "この補足項目を削除" });
    deleteButton.addEventListener("click", function () {
      if (!confirmAction("補足項目「" + (stakeholder.label || "名称未設定") + "」と入力内容を削除しますか？")) return;
      deleteStakeholderNote(sheet.sheetId, stakeholder.id);
      flushPendingSave();
      showToast("補足項目を削除しました。");
      renderInputMode(sheet.sheetId);
    });
    appendChild(editor.querySelector(".note-editor-footer"), deleteButton);
    return editor;
  }

  function createStakeholderAddForm(sheetId) {
    const form = el("form", { className: "panel add-stakeholder" });
    const field = makeField("追加する補足項目名", "stakeholderLabel", "text", "", { required: true, placeholder: "例：お子様、同居予定のご家族" });
    appendChild(form, el("div", {}, el("p", { className: "section-number", text: "＋" }), el("h2", { text: "補足項目を追加" }), el("p", { className: "field-hint", text: "追加した項目は、このページを開いている間だけ保持されます。必要な場合はCSVを書き出してください。" })));
    appendChild(form, el("div", { className: "add-stakeholder-row" }, field.wrapper, el("button", { className: "button", type: "submit", text: "補足項目を追加" })));
    form.addEventListener("submit", function (event) {
      event.preventDefault();
      const label = field.input.value.trim();
      if (!label) return;
      addStakeholderNote(sheetId, label);
      flushPendingSave();
      showToast("補足項目を追加しました。");
      renderInputMode(sheetId);
    });
    return form;
  }

  // ---------------------------------------------------------------------------
  // Internal review mode
  // ---------------------------------------------------------------------------

  function renderInternalReviewMode(sheetId) {
    revokeImageObjectUrls();
    const sheet = getSheetById(sheetId);
    if (!sheet) return renderNotFound();
    state.activeSheetId = sheetId;
    setPageTitle(sheet.customerName + "・社内用振り返り");
    const page = el("section", { className: "page workspace-page" });
    appendChild(page, createWorkspaceHeader(sheet, "internal"));
    appendChild(page, createStorageNotice());
    appendChild(page, el("section", { className: "panel review-intro" },
      el("div", {}, el("p", { className: "eyebrow", text: "INTERNAL REVIEW" }), el("h2", { text: "一回の商談内容を社内用に振り返る" }), el("p", { text: "主要項目と関係者ごとの補足を、入力欄ではなく読みやすい文章で一覧表示しています。" })),
      el("dl", { className: "basic-summary" },
        summaryPair("お施主様名", sheet.customerName),
        summaryPair("商談日", formatDate(sheet.meetingDate)),
        summaryPair("商談回数", "第" + sheet.meetingNumber + "回"),
        summaryPair("商談タイトル", sheet.title || "未入力"),
        summaryPair("最終更新", formatDateTime(sheet.updatedAt))
      )
    ));
    const reviewList = el("div", { className: "review-list" });
    coreDefinitions.forEach(function (definition, index) {
      const section = reviewSection(index + 1, definition.label, sheet.coreNotes[definition.key].content);
      const gallery = el("div", { className: "image-gallery image-gallery-review" });
      appendChild(section, gallery);
      activeImageRenderPromises.push(renderCategoryImages(sheetId, definition.key, "internal", gallery));
      appendChild(reviewList, section);
    });
    sheet.stakeholderNotes.forEach(function (stakeholder, index) {
      appendChild(reviewList, reviewSection(index + 5, stakeholder.label, stakeholder.content));
    });
    appendChild(page, reviewList);
    appendChild(page, el("div", { className: "bottom-actions no-print" },
      createCsvExportButton(sheetId),
      modeButton(sheetId, "input", "入力モードへ戻る", "internal"),
      modeButton(sheetId, "customer", "お客様用で確認", "internal")
    ));
    app.replaceChildren(page);
  }

  function summaryPair(label, value) {
    return el("div", {}, el("dt", { text: label }), el("dd", { text: value }));
  }

  function reviewSection(index, label, content) {
    const filled = Boolean(content.trim());
    return el("section", { className: "panel review-section" + (filled ? "" : " is-empty") },
      el("div", { className: "review-heading" },
        el("div", {}, el("span", { className: "section-number", text: String(index).padStart(2, "0") }), el("h2", { text: label })),
        el("span", { className: "entry-status" + (filled ? " is-filled" : ""), text: filled ? "入力済み" : "未入力" })
      ),
      el("p", { className: "review-content", text: filled ? content : "未入力" })
    );
  }

  // ---------------------------------------------------------------------------
  // Customer preview and print mode. Only the four core categories are rendered.
  // ---------------------------------------------------------------------------

  function renderCustomerPreviewMode(sheetId) {
    revokeImageObjectUrls();
    const sheet = getSheetById(sheetId);
    if (!sheet) return renderNotFound();
    state.activeSheetId = sheetId;
    initializeCustomerPreview(sheetId);
    setPageTitle(sheet.customerName + "・お客様用プレビュー");
    const page = el("section", { className: "page workspace-page customer-mode-page" });
    appendChild(page, createWorkspaceHeader(sheet, "customer"));

    const controls = el("aside", { className: "panel preview-controls no-print" },
      el("div", {}, el("p", { className: "eyebrow", text: "OUTPUT SETTINGS" }), el("h2", { text: "表示する項目" }), el("p", { text: "チェックした4項目だけをPDFに出力します。文章の編集は元の社内記録を変更しません。" }))
    );
    function createCustomerPrintBrand() {
      return el("div", { className: "customer-print-brand" },
        el("span", { className: "customer-print-logo-crop" },
          el("img", { className: "customer-print-logo", attrs: { src: "logo_clr.png", alt: "NEST HOUSE" } })
        ),
        el("p", { text: 'Make your everyday life a fun "always"' }),
        el("p", { text: "いつもの毎日を楽しい“いつも”に" })
      );
    }
    function createCustomerDocumentHeader(includeMeta, includeTitle, titleText, subtitleText) {
      return el("header", { className: "customer-document-header" },
        includeTitle ? el("div", { className: "customer-document-title-block" },
          el("p", { className: "document-kicker", text: "MEETING SHEET" }),
          el("h1", { text: titleText || "前回のお打ち合わせ内容" }),
          subtitleText ? el("p", { className: "customer-document-subtitle", text: subtitleText }) : null
        ) : el("div", { className: "customer-document-title-block customer-document-title-spacer", attrs: { "aria-hidden": "true" } }),
        createCustomerPrintBrand(),
        includeMeta ? el("dl", { className: "document-meta" },
          summaryPair("お客様名", sheet.customerName),
          summaryPair("前回のお打ち合わせ日", formatDate(sheet.meetingDate)),
          summaryPair("商談回数", "第" + sheet.meetingNumber + "回"),
          summaryPair("商談タイトル", sheet.title || "—"),
          summaryPair("出力日", formatDate(todayIso()))
        ) : null
      );
    }
    function createCustomerDocumentPage(className, includeMeta, includeTitle, label, titleText, subtitleText) {
      return el("article", { className: "customer-document " + className, attrs: { "aria-label": label } },
        createCustomerDocumentHeader(includeMeta, includeTitle, titleText, subtitleText)
      );
    }
    function sectionLengthClass(value) {
      const length = String(value || "").length;
      if (length >= 360) return " is-very-long";
      if (length >= 180) return " is-long";
      return "";
    }
    function syncSectionLengthClass(section, value) {
      const length = String(value || "").length;
      section.classList.toggle("is-long", length >= 180 && length < 360);
      section.classList.toggle("is-very-long", length >= 360);
    }
    function resizePreviewTextarea(textarea) {
      textarea.style.height = "auto";
      textarea.style.height = Math.max(textarea.scrollHeight, 150) + "px";
    }
    const textDocument = createCustomerDocumentPage("customer-text-page customer-document-text", true, true, "お客様用資料プレビュー 1ページ目", "前回のお打ち合わせ内容");
    const currentStatusArea = createCustomerCurrentStatusArea(sheet);
    const topInfoRow = createCustomerTopInfoRow(sheet, currentStatusArea);
    const customerDocumentSections = el("div", { className: "customer-output-body" + (topInfoRow ? " has-top-info" : "") });
    if (topInfoRow) appendChild(customerDocumentSections, topInfoRow);
    const customerLowerGrid = el("div", { className: "customer-lower-grid" });
    const customerRightStack = el("div", { className: "customer-right-stack" });
    const previewEmpty = el("div", { className: "customer-preview-empty no-print" },
      el("p", { className: "empty-state-kicker", text: "NO OUTPUT CONTENT" }),
      el("h2", { text: "お客様用に表示できる内容がまだ入力されていません" }),
      el("p", { text: "入力モードで主要4項目を入力すると、このプレビューに自動反映されます。" }),
      createNavButton("button", "入力モードへ戻る", function () { navigateToSheet(sheetId, "input"); })
    );
    let printButton = null;
    function hasVisiblePreviewContent() {
      return coreDefinitions.some(function (definition) {
        return sheet.customerPreview.visibleCategories.includes(definition.key) && hasText(getCustomerPreviewContentFromSheet(sheet, definition.key));
      });
    }
    function hasAnyPreviewContent() {
      return coreDefinitions.some(function (definition) {
        return hasText(getCustomerPreviewContentFromSheet(sheet, definition.key));
      });
    }
    function updatePreviewEmptyState() {
      const hasCurrentStatus = getCurrentStatusRows(sheet).length > 0;
      const hasTopPhoto = Boolean(getCustomerTopPhotoAttachment(sheet));
      const isEmpty = !hasAnyPreviewContent() && !hasCurrentStatus && !hasTopPhoto;
      previewEmpty.hidden = !isEmpty;
      if (printButton) printButton.disabled = !hasVisiblePreviewContent() && !hasCurrentStatus && !hasTopPhoto;
    }

    coreDefinitions.forEach(function (definition, index) {
      const visible = sheet.customerPreview.visibleCategories.includes(definition.key);
      const outputLabel = getCustomerOutputLabel(definition.key);
      const checkboxId = "visible-" + definition.key;
      const checkbox = el("input", { id: checkboxId, type: "checkbox", checked: visible });
      appendChild(controls, el("label", { className: "preview-choice", attrs: { for: checkboxId } }, checkbox, el("span", { text: outputLabel.title })));

      const value = getCustomerPreviewContent(sheetId, definition.key);
      const textareaId = "preview-" + definition.key;
      const textarea = el("textarea", { id: textareaId, className: "customer-section-body-input", value: value, attrs: { rows: "5" }, placeholder: "お客様へ共有する文章を入力してください" });
      const printText = el("p", { className: "preview-print-text customer-section-body print-only", text: value });
      const characterCount = el("span", { className: "character-count", text: value.length + "文字" });
      const overrideNotice = el("p", {
        className: "preview-origin-note",
        text: "入力内容が更新されています。お客様用文章を元の入力内容に戻すこともできます。",
        attrs: hasCustomerPreviewOverride(sheet, definition.key) ? {} : { hidden: "" }
      });
      const reflectButton = el("button", { className: "text-button", type: "button", text: "元の入力内容を反映" });
      const clearEditButton = el("button", { className: "text-button", type: "button", text: "お客様用編集をクリア" });
      const editActions = el("div", { className: "preview-edit-actions", attrs: hasCustomerPreviewOverride(sheet, definition.key) ? {} : { hidden: "" } }, reflectButton, clearEditButton);
      function syncPreviewSectionText() {
        const current = getCustomerPreviewContent(sheetId, definition.key);
        textarea.value = current;
        printText.textContent = current;
        characterCount.textContent = current.length + "文字";
        syncSectionLengthClass(section, current);
        resizePreviewTextarea(textarea);
        const hasOverride = hasCustomerPreviewOverride(sheet, definition.key);
        overrideNotice.hidden = !hasOverride;
        editActions.hidden = !hasOverride;
        updatePreviewEmptyState();
      }
      function clearPreviewOverride() {
        clearCustomerPreviewContent(sheetId, definition.key);
        syncPreviewSectionText();
      }
      reflectButton.addEventListener("click", clearPreviewOverride);
      clearEditButton.addEventListener("click", clearPreviewOverride);
      const section = el("section", { className: "customer-grid-section " + getCustomerSectionAreaClass(definition.key) + sectionLengthClass(value), attrs: visible ? {} : { hidden: "" } },
        el("div", { className: "customer-section-card" },
          createCustomerGridSectionHeader(definition.key),
          el("label", { className: "sr-only", text: outputLabel.title + "の文章", attrs: { for: textareaId } }),
          textarea,
          printText,
          el("div", { className: "preview-edit-meta no-print" },
            el("div", {}, characterCount, overrideNotice),
            editActions
          )
        )
      );
      checkbox.addEventListener("change", function () {
        toggleCustomerPreviewCategory(sheetId, definition.key, checkbox.checked);
        renderCustomerPreviewMode(sheetId);
      });
      textarea.addEventListener("input", function () {
        updateCustomerPreviewContent(sheetId, definition.key, textarea.value);
        if (!hasCustomerPreviewOverride(sheet, definition.key)) {
          textarea.value = getCustomerPreviewContent(sheetId, definition.key);
        }
        printText.textContent = textarea.value;
        characterCount.textContent = textarea.value.length + "文字";
        syncSectionLengthClass(section, textarea.value);
        resizePreviewTextarea(textarea);
        const hasOverride = hasCustomerPreviewOverride(sheet, definition.key);
        overrideNotice.hidden = !hasOverride;
        editActions.hidden = !hasOverride;
        updatePreviewEmptyState();
      });
      if (definition.key === "scheduleReason") {
        appendChild(customerDocumentSections, section);
      } else if (definition.key === "budget") {
        appendChild(customerLowerGrid, section);
      } else if (definition.key === "building" || definition.key === "land") {
        appendChild(customerRightStack, section);
      } else {
        appendChild(customerDocumentSections, section);
      }
      window.requestAnimationFrame(function () { resizePreviewTextarea(textarea); });
    });
    appendChild(customerLowerGrid, customerRightStack);
    appendChild(customerDocumentSections, customerLowerGrid);
    appendChild(customerDocumentSections, previewEmpty);

    printButton = el("button", { className: "button", type: "button", text: "印刷・PDF保存" });
    printButton.addEventListener("click", printCustomerPreview);
    appendChild(controls, printButton);
    appendChild(textDocument, customerDocumentSections);
    appendChild(textDocument, el("footer", { className: "customer-document-footer" },
      el("p", { text: "お打ち合わせ内容に相違がないか、次回の冒頭でご一緒に確認させてください。" })
    ));
    updatePreviewEmptyState();
    appendChild(page, el("div", { className: "preview-layout" }, controls, el("div", { className: "customer-preview-pages" }, textDocument)));
    app.replaceChildren(page);
  }

  // ---------------------------------------------------------------------------
  // Router, data controls, and error state
  // ---------------------------------------------------------------------------

  function renderNotFound() {
    revokeImageObjectUrls();
    saveStatusNode = null;
    setPageTitle("シートが見つかりません");
    app.replaceChildren(el("section", { className: "page page-narrow" },
      el("div", { className: "panel empty-state" },
        el("h1", { text: "商談シートが見つかりません" }),
        el("p", { text: "現在の入力内容が破棄されたか、表示対象が見つかりません。" }),
        createNavButton("button", "開始画面へ戻る", navigateToStart)
      )
    ));
  }

  function renderApp() {
    window.scrollTo(0, 0);
    window.requestAnimationFrame(updateBackToTopVisibility);
    if (currentView.name === "new") return renderNewSheetForm();
    if (currentView.name === "sheet" && currentView.sheetId) {
      if (currentView.mode === "customer") return renderCustomerPreviewMode(currentView.sheetId);
      return renderInputMode(currentView.sheetId);
    }
    return renderSheetList();
  }

  document.getElementById("clear-data").addEventListener("click", async function () {
    if (!confirmAction("現在の入力内容を破棄しますか？CSVを書き出していない内容は失われます。")) return;
    window.clearTimeout(pendingSaveTimer);
    pendingSaveTimer = null;
    try {
      await clearLocalData();
      showToast("現在の入力内容を破棄しました。");
      navigateToStart();
    } catch (error) {
      console.error("入力内容を破棄できませんでした。", error);
      showToast("入力内容を破棄できませんでした。もう一度お試しください。", true);
    }
  });

  ["brand-start", "nav-start"].forEach(function (id) {
    const button = document.getElementById(id);
    if (button) button.addEventListener("click", navigateToStart);
  });

  window.addEventListener("scroll", updateBackToTopVisibility, { passive: true });
  if (backToTopButton) backToTopButton.addEventListener("click", scrollToPageTop);
  clearImageMemoryStore().finally(function () {
    openImageMemoryStore()
      .finally(function () {
        renderApp();
        updateBackToTopVisibility();
      });
  });
})();

