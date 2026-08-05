(function () {
  "use strict";

  const coreNoteDefinitions = [
    { key: "scheduleReason", label: "計画時期・建築理由", customerLabel: "計画時期・家づくりのきっかけ", hint: "ご希望や気になっていることを入力" },
    { key: "budget", label: "予算", customerLabel: "資金計画", hint: "ご希望や気になっていることを入力" },
    { key: "building", label: "建物", customerLabel: "建物についてのご希望", hint: "ご希望や気になっていることを入力" },
    { key: "land", label: "土地", customerLabel: "土地についてのご希望", hint: "ご希望や気になっていることを入力" }
  ];

  function coreNotes(contents) {
    const values = contents || {};
    const result = {};
    coreNoteDefinitions.forEach(function (definition) {
      result[definition.key] = {
        label: definition.label,
        content: values[definition.key] || ""
      };
    });
    return result;
  }

  function customerPreview(visibleCategories, editedContents, visibleImageIds) {
    return {
      visibleCategories: visibleCategories || [],
      editedContents: Object.assign({
        scheduleReason: "",
        budget: "",
        building: "",
        land: ""
      }, editedContents || {}),
      visibleImageIds: visibleImageIds || [],
      visibilityTouched: false
    };
  }

  const emptyData = {
    version: 2,
    activeSheetId: null,
    sheets: []
  };

  window.AppData = {
    coreNoteDefinitions: coreNoteDefinitions,
    createCoreNotes: function (contents) { return coreNotes(contents); },
    createCustomerPreview: function (visible, edited, visibleImages) { return customerPreview(visible, edited, visibleImages); },
    createEmptyData: function () { return JSON.parse(JSON.stringify(emptyData)); }
  };
})();
