(function (globalScope) {
  'use strict';

  function clamp(value, minimum, maximum) {
    return Math.max(minimum, Math.min(value, maximum));
  }

  function normalizeRowStyleNames(rawNames) {
    if (!Array.isArray(rawNames)) {
      return [];
    }

    const seen = new Set();
    const normalized = [];

    for (const rawName of rawNames) {
      if (typeof rawName !== 'string') {
        continue;
      }

      const name = rawName.trim();
      if (!name) {
        continue;
      }

      const key = name.toLowerCase();
      if (seen.has(key)) {
        continue;
      }

      seen.add(key);
      normalized.push(name);
    }

    return normalized;
  }

  function computeMetadataTitleSync(sheetName, currentTitle, lastAutoSyncedTitle) {
    const safeSheetName = typeof sheetName === 'string' ? sheetName : '';
    const safeCurrentTitle = typeof currentTitle === 'string' ? currentTitle : '';
    const safeLastAutoSyncedTitle = typeof lastAutoSyncedTitle === 'string' ? lastAutoSyncedTitle : '';

    const titleIsEmpty = safeCurrentTitle.trim().length === 0;
    const titleMatchesLastAuto = safeCurrentTitle === safeLastAutoSyncedTitle;

    if (titleIsEmpty || titleMatchesLastAuto) {
      return {
        title: safeSheetName,
        lastAutoSyncedTitle: safeSheetName,
        wasAutoSynced: true
      };
    }

    return {
      title: safeCurrentTitle,
      lastAutoSyncedTitle: safeLastAutoSyncedTitle,
      wasAutoSynced: false
    };
  }

  function computeInsertRange(scriptLength, cursorIndex, selectedStart, selectedEnd) {
    const safeLength = Number.isFinite(scriptLength) ? Math.max(0, scriptLength) : 0;

    if (Number.isFinite(selectedStart) && Number.isFinite(selectedEnd) && selectedEnd > selectedStart) {
      const replaceStart = clamp(selectedStart, 0, safeLength);
      const replaceEnd = clamp(selectedEnd, replaceStart, safeLength);
      return {
        start: replaceStart,
        end: replaceEnd,
        isReplace: true
      };
    }

    const insertIndex = Number.isFinite(cursorIndex) ? clamp(cursorIndex, 0, safeLength) : 0;
    return {
      start: insertIndex,
      end: insertIndex,
      isReplace: false
    };
  }

  function computeSubmenuLayout(options) {
    const settings = options || {};
    const viewportPadding = Number.isFinite(settings.viewportPadding) ? settings.viewportPadding : 8;
    const minimumHeight = Number.isFinite(settings.minimumHeight) ? settings.minimumHeight : 120;

    const menuRight = Number.isFinite(settings.menuRight) ? settings.menuRight : 0;
    const wrapperTop = Number.isFinite(settings.wrapperTop) ? settings.wrapperTop : 0;
    const submenuWidth = Number.isFinite(settings.submenuWidth) ? settings.submenuWidth : 190;
    const submenuHeight = Number.isFinite(settings.submenuHeight) ? settings.submenuHeight : 0;
    const viewportWidth = Number.isFinite(settings.viewportWidth) ? settings.viewportWidth : 1024;
    const viewportHeight = Number.isFinite(settings.viewportHeight) ? settings.viewportHeight : 768;

    const openLeft = viewportWidth - menuRight < submenuWidth + 16;

    let submenuTop = 0;
    const spaceBelowWrapper = viewportHeight - wrapperTop - viewportPadding;
    if (submenuHeight > spaceBelowWrapper) {
      submenuTop = spaceBelowWrapper - submenuHeight;
    }

    const projectedTop = wrapperTop + submenuTop;
    if (projectedTop < viewportPadding) {
      submenuTop += viewportPadding - projectedTop;
    }

    const adjustedTop = wrapperTop + submenuTop;
    const maxHeight = Math.max(minimumHeight, Math.floor(viewportHeight - adjustedTop - viewportPadding));

    return {
      openLeft,
      submenuTop,
      maxHeight
    };
  }

  const exported = {
    normalizeRowStyleNames,
    computeMetadataTitleSync,
    computeInsertRange,
    computeSubmenuLayout
  };

  if (typeof module !== 'undefined' && module.exports) {
    module.exports = exported;
  }

  globalScope.Query2ExcelWebUtils = exported;
})(typeof window !== 'undefined' ? window : globalThis);
