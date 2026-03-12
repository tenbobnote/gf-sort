/**
 * app.js — Main entry point. Handles initialization, UI events, and
 * orchestrates section detection, sorting, column conversion, and
 * comment preservation.
 */
var GF = GF || {};

GF.App = (function () {
  "use strict";

  var _currentSection = null;
  var _currentColumnCount = null;
  var _detectionTimeout = null;
  var _monosCount = 0;
  var BUILD_VERSION = "__BUILD_VERSION__";

  // ── Initialization ──

  function init() {
    // Bind UI events
    document.getElementById("sort-btn").addEventListener("click", onSortClick);
    document
      .getElementById("convert-btn")
      .addEventListener("click", onConvertClick);
    document
      .getElementById("sort-all-btn")
      .addEventListener("click", onSortAllClick);
    document
      .getElementById("add-mono-btn")
      .addEventListener("click", onAddMonoClick);
    document
      .getElementById("mono-input")
      .addEventListener("keypress", function (e) {
        if (e.key === "Enter") onAddMonoClick();
      });
    document
      .getElementById("import-monos-btn")
      .addEventListener("click", onImportMonosClick);
    document
      .getElementById("import-file")
      .addEventListener("change", onImportFileSelected);

    // Listen for selection changes to auto-detect section
    Office.context.document.addHandlerAsync(
      Office.EventType.DocumentSelectionChanged,
      onSelectionChanged
    );

    // Initial detection
    updateSectionDisplay();

    // Load monos count
    loadMonosInfo();

    // Version display and refresh
    var versionEl = document.getElementById("version-text");
    if (BUILD_VERSION !== "__BUILD_" + "VERSION__") {
      versionEl.textContent = "v" + BUILD_VERSION;
    } else {
      versionEl.textContent = "dev";
    }
    document
      .getElementById("refresh-btn")
      .addEventListener("click", function () {
        window.location.reload(true);
      });

    setStatus("Ready");
  }

  // ── Section Detection ──

  function onSelectionChanged() {
    clearTimeout(_detectionTimeout);
    _detectionTimeout = setTimeout(updateSectionDisplay, 500);
  }

  async function updateSectionDisplay() {
    try {
      var result = await GF.Document.detectSection();

      if (result.error) {
        showNoSection(result.error);
        _currentSection = null;
        _currentColumnCount = null;
        return;
      }

      _currentSection = result;

      // Get column count for convert button label
      var colCount = await GF.Document.getColumnCount();
      _currentColumnCount = colCount;

      showSection(result, colCount);
    } catch (e) {
      showNoSection("Place your cursor in a noun-group table.");
      _currentSection = null;
      _currentColumnCount = null;
    }
  }

  function showSection(section, colCount) {
    document.getElementById("no-section").classList.add("hidden");
    document.getElementById("current-section").classList.remove("hidden");
    document.getElementById("category-name").textContent = section.pageTitle;
    document.getElementById("section-type").textContent = section.sectionType;

    var genderLabel = { m: "Masculine", f: "Feminine", n: "Neuter" };
    document.getElementById("main-gender").textContent =
      genderLabel[section.mainGender] || "";

    // Show column count
    var colCountEl = document.getElementById("column-count");
    if (colCountEl && colCount) {
      colCountEl.textContent = colCount + " columns";
    }

    // Enable sort button
    document.getElementById("sort-btn").disabled = false;

    // Update convert button label and visibility
    var convertBtn = document.getElementById("convert-btn");
    if (colCount === 4) {
      convertBtn.textContent = "Convert to 3 Columns";
      convertBtn.classList.remove("hidden");
      convertBtn.disabled = false;
    } else if (colCount === 3) {
      convertBtn.textContent = "Convert to 4 Columns";
      convertBtn.classList.remove("hidden");
      convertBtn.disabled = false;
    } else {
      convertBtn.classList.add("hidden");
    }
  }

  function showNoSection(message) {
    document.getElementById("no-section").classList.remove("hidden");
    document.getElementById("no-section-msg").textContent = message;
    document.getElementById("current-section").classList.add("hidden");
    document.getElementById("sort-btn").disabled = true;
    document.getElementById("convert-btn").classList.add("hidden");
    document.getElementById("uncolored-warning").classList.add("hidden");
  }

  // ── Sort Operation ──

  async function onSortClick() {
    if (!_currentSection) return;

    var sortBtn = document.getElementById("sort-btn");
    var convertBtn = document.getElementById("convert-btn");
    sortBtn.disabled = true;
    convertBtn.disabled = true;
    setStatus("Reading table...");

    try {
      // 1. Read items from the table
      var tableData = await GF.Document.readTableItems();
      if (tableData.error) {
        setStatus("Error: " + tableData.error);
        sortBtn.disabled = false;
        convertBtn.disabled = false;
        return;
      }

      // 2. Check for uncolored words
      if (tableData.uncolored.length > 0) {
        showUncoloredWarning(tableData.uncolored);
        setStatus("Fix uncolored words before sorting.");
        sortBtn.disabled = false;
        convertBtn.disabled = false;
        return;
      }
      hideUncoloredWarning();

      // 3. Capture comments (if supported)
      setStatus("Checking comments...");
      var commentMap = {};
      var commentCount = 0;
      var tableCommentCount = await GF.Comments.countCommentsInTable();

      if (tableCommentCount > 0) {
        commentMap = await GF.Comments.captureComments();
        commentCount = GF.Comments.countComments(commentMap);

        if (commentCount > 0) {
          setStatus(
            "Found " +
              commentCount +
              " comment(s). Tracking for restoration..."
          );
        } else {
          setStatus(
            "Warning: " +
              tableCommentCount +
              " comment(s) detected but could not be captured. Sort aborted."
          );
          sortBtn.disabled = false;
          convertBtn.disabled = false;
          return;
        }
      } else if (tableCommentCount === -1) {
        setStatus("Comments API not available. Comments may be lost.");
      }

      // 4. Load monos (clear cache to ensure fresh read from document)
      setStatus("Loading reference data...");
      GF.Monos.clearCache();
      var monosSet = await GF.Monos.getMonosSet();

      // 5. Sort
      setStatus("Sorting " + tableData.items.length + " items...");
      var sorted = GF.Sorting.sortItems(
        tableData.items,
        _currentSection.sectionType,
        _currentSection.mainGender,
        monosSet
      );

      // 6. Check if order actually changed
      var changed = false;
      for (var i = 0; i < sorted.length; i++) {
        if (sorted[i].word !== tableData.items[i].word) {
          changed = true;
          break;
        }
      }

      if (!changed) {
        // Even if order hasn't changed, still apply column alignment
        setStatus("Order unchanged. Applying column alignment...");
      }

      // 7. Delete original comments (before rewriting cells)
      if (commentCount > 0) {
        setStatus("Removing original comments...");
        await GF.Comments.deleteOriginalComments(commentMap);
      }

      // 8. Write sorted items with page-level column alignment
      setStatus("Writing sorted items + aligning columns...");
      var result = await GF.Document.sortWithAlignment(
        sorted,
        tableData.numGroups
      );

      // 9. Restore comments in new positions
      if (commentCount > 0) {
        setStatus("Restoring comments...");
        await GF.Comments.restoreComments(commentMap);
      }

      var statusMsg =
        "Sorted " + sorted.length + " items. Use Ctrl+Z to undo.";
      if (result.shrinkCount > 0) {
        statusMsg += " " + result.shrinkCount + " cell(s) font-reduced.";
      }
      if (result.pageTableCount > 1) {
        statusMsg +=
          " Aligned " + result.pageTableCount + " tables on page.";
      }
      if (commentCount > 0) {
        statusMsg += " " + commentCount + " comment(s) preserved.";
      }
      setStatus(statusMsg);
    } catch (e) {
      console.error("Sort error:", e);
      setStatus("Error: " + e.message);
    }

    sortBtn.disabled = false;
    convertBtn.disabled = false;
  }

  // ── Convert Operation ──

  async function onConvertClick() {
    if (!_currentSection || !_currentColumnCount) return;

    var targetGroups = _currentColumnCount === 4 ? 3 : 4;
    var sortBtn = document.getElementById("sort-btn");
    var convertBtn = document.getElementById("convert-btn");
    sortBtn.disabled = true;
    convertBtn.disabled = true;

    setStatus("Reading table...");

    try {
      // 1. Read items from the table
      var tableData = await GF.Document.readTableItems();
      if (tableData.error) {
        setStatus("Error: " + tableData.error);
        sortBtn.disabled = false;
        convertBtn.disabled = false;
        return;
      }

      // 2. Check for uncolored words
      if (tableData.uncolored.length > 0) {
        showUncoloredWarning(tableData.uncolored);
        setStatus("Fix uncolored words before converting.");
        sortBtn.disabled = false;
        convertBtn.disabled = false;
        return;
      }
      hideUncoloredWarning();

      // 3. Capture comments
      setStatus("Checking comments...");
      var commentMap = {};
      var commentCount = 0;
      var tableCommentCount = await GF.Comments.countCommentsInTable();

      if (tableCommentCount > 0) {
        commentMap = await GF.Comments.captureComments();
        commentCount = GF.Comments.countComments(commentMap);

        if (commentCount > 0) {
          setStatus(
            "Found " + commentCount + " comment(s). Tracking for restoration..."
          );
        } else {
          setStatus(
            "Warning: " +
              tableCommentCount +
              " comment(s) detected but could not be captured. Convert aborted."
          );
          sortBtn.disabled = false;
          convertBtn.disabled = false;
          return;
        }
      } else if (tableCommentCount === -1) {
        setStatus("Comments API not available. Comments may be lost.");
      }

      // 4. Load monos
      setStatus("Loading reference data...");
      GF.Monos.clearCache();
      var monosSet = await GF.Monos.getMonosSet();

      // 5. Sort items
      setStatus("Sorting " + tableData.items.length + " items...");
      var sorted = GF.Sorting.sortItems(
        tableData.items,
        _currentSection.sectionType,
        _currentSection.mainGender,
        monosSet
      );

      // 6. Delete original comments
      if (commentCount > 0) {
        setStatus("Removing original comments...");
        await GF.Comments.deleteOriginalComments(commentMap);
      }

      // 7. Convert: delete old table, create new with target columns, align
      setStatus(
        "Converting to " +
          targetGroups +
          " columns + sorting + aligning..."
      );
      var result = await GF.Document.convertAndSort(sorted, targetGroups);

      // 8. Restore comments in the new table
      if (commentCount > 0) {
        setStatus("Restoring comments...");
        await GF.Comments.restoreComments(commentMap);
      }

      // Update column count state
      _currentColumnCount = targetGroups;

      var statusMsg =
        "Converted to " +
        targetGroups +
        " columns. " +
        sorted.length +
        " items sorted. Use Ctrl+Z to undo.";
      if (result.shrinkCount > 0) {
        statusMsg += " " + result.shrinkCount + " cell(s) font-reduced.";
      }
      if (result.pageTableCount > 1) {
        statusMsg +=
          " Aligned " + result.pageTableCount + " tables on page.";
      }
      if (commentCount > 0) {
        statusMsg += " " + commentCount + " comment(s) preserved.";
      }
      setStatus(statusMsg);

      // Refresh section display (column count changed)
      updateSectionDisplay();
    } catch (e) {
      console.error("Convert error:", e);
      setStatus("Error: " + e.message);
    }

    sortBtn.disabled = false;
    convertBtn.disabled = false;
  }

  // ── Sort All Operation ──

  async function onSortAllClick() {
    var sortAllBtn = document.getElementById("sort-all-btn");
    var sortBtn = document.getElementById("sort-btn");
    var convertBtn = document.getElementById("convert-btn");
    sortAllBtn.disabled = true;
    sortBtn.disabled = true;
    convertBtn.disabled = true;

    try {
      setStatus("Loading reference data...");
      GF.Monos.clearCache();
      var monosSet = await GF.Monos.getMonosSet();

      var tableCount = await GF.Document.getTableCount();
      var sorted = 0;
      var skipped = 0;
      var totalShrinks = 0;

      for (var t = 0; t < tableCount; t++) {
        setStatus(
          "Processing table " + (t + 1) + " of " + tableCount + "..."
        );

        // Select this table
        var selected = await GF.Document.selectTable(t);
        if (!selected) {
          skipped++;
          continue;
        }

        // Detect section
        var section = await GF.Document.detectSection();
        if (section.error) {
          skipped++;
          continue;
        }

        // Read items
        var tableData = await GF.Document.readTableItems();
        if (tableData.error || tableData.items.length === 0) {
          skipped++;
          continue;
        }
        if (tableData.uncolored.length > 0) {
          skipped++;
          continue;
        }

        // Capture comments
        var commentMap = {};
        var commentCount = 0;
        var tableCommentCount = await GF.Comments.countCommentsInTable();
        if (tableCommentCount > 0) {
          commentMap = await GF.Comments.captureComments();
          commentCount = GF.Comments.countComments(commentMap);
          if (commentCount > 0) {
            await GF.Comments.deleteOriginalComments(commentMap);
          }
        }

        // Sort
        var sortedItems = GF.Sorting.sortItems(
          tableData.items,
          section.sectionType,
          section.mainGender,
          monosSet
        );

        // Write with alignment
        var result = await GF.Document.sortWithAlignment(
          sortedItems,
          tableData.numGroups
        );
        totalShrinks += result.shrinkCount;

        // Restore comments
        if (commentCount > 0) {
          await GF.Comments.restoreComments(commentMap);
        }

        sorted++;
      }

      var msg = "Done! Sorted " + sorted + " tables, skipped " + skipped + ".";
      if (totalShrinks > 0) {
        msg += " " + totalShrinks + " cell(s) font-reduced.";
      }
      msg += " Use Ctrl+Z to undo.";
      setStatus(msg);
    } catch (e) {
      console.error("Sort all error:", e);
      setStatus("Error: " + e.message);
    }

    sortAllBtn.disabled = false;
    sortBtn.disabled = false;
    convertBtn.disabled = false;
  }

  // ── Uncolored Words Warning ──

  function showUncoloredWarning(words) {
    var container = document.getElementById("uncolored-warning");
    var list = document.getElementById("uncolored-list");
    list.innerHTML = "";
    for (var i = 0; i < words.length; i++) {
      var li = document.createElement("li");
      li.textContent = words[i];
      list.appendChild(li);
    }
    container.classList.remove("hidden");
  }

  function hideUncoloredWarning() {
    document.getElementById("uncolored-warning").classList.add("hidden");
  }

  // ── Monos Management ──

  async function loadMonosInfo() {
    try {
      var monos = await GF.Monos.getMonosSet();
      _monosCount = monos.size;
      document.getElementById("mono-count").textContent =
        _monosCount + " words";

      if (_monosCount === 0) {
        document.getElementById("monos-empty").classList.remove("hidden");
      } else {
        document.getElementById("monos-empty").classList.add("hidden");
      }
    } catch (e) {
      document.getElementById("mono-count").textContent = "Error loading";
    }
  }

  async function onAddMonoClick() {
    var input = document.getElementById("mono-input");
    var word = input.value.trim();
    if (!word) return;

    var addBtn = document.getElementById("add-mono-btn");
    addBtn.disabled = true;

    try {
      GF.Monos.clearCache();
      var monos = await GF.Monos.getMonosSet();
      var beforeCount = monos.size;

      if (monos.has(word)) {
        setStatus(
          '"' +
            word +
            '" is already in the monos list (' +
            beforeCount +
            " words)."
        );
        addBtn.disabled = false;
        return;
      }

      monos.add(word);
      await GF.Monos.saveMonosSet(monos);

      input.value = "";
      _monosCount = monos.size;
      document.getElementById("mono-count").textContent =
        _monosCount + " words";
      document.getElementById("monos-empty").classList.add("hidden");
      setStatus(
        'Added "' +
          word +
          '": ' +
          beforeCount +
          " \u2192 " +
          _monosCount +
          " monos."
      );
    } catch (e) {
      console.error("Failed to add mono:", e);
      setStatus("FAILED at last step. Error: " + e.message);
    }

    addBtn.disabled = false;
  }

  function onImportMonosClick() {
    document.getElementById("import-file").click();
  }

  async function onImportFileSelected(e) {
    var file = e.target.files[0];
    if (!file) return;

    var reader = new FileReader();
    reader.onload = async function (evt) {
      try {
        var monos = await GF.Monos.importFromCsv(evt.target.result);
        _monosCount = monos.size;
        document.getElementById("mono-count").textContent =
          _monosCount + " words";
        document.getElementById("monos-empty").classList.add("hidden");
        setStatus("Imported " + _monosCount + " monos from file.");
      } catch (err) {
        setStatus("Error importing: " + err.message);
      }
    };
    reader.readAsText(file);

    // Reset file input so the same file can be re-selected
    e.target.value = "";
  }

  // ── Status ──

  function setStatus(text) {
    document.getElementById("status-text").textContent = text;
  }

  return {
    init: init,
  };
})();

// ── Office.js entry point ──
Office.onReady(function (info) {
  if (info.host === Office.HostType.Word) {
    GF.App.init();
  }
});
