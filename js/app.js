/**
 * app.js — Main entry point. Handles initialization, UI events, and
 * orchestrates section detection, sorting, and comment preservation.
 */
var GF = GF || {};

GF.App = (function () {
  "use strict";

  var _currentSection = null;
  var _detectionTimeout = null;
  var _monosCount = 0;

  // ── Initialization ──

  function init() {
    // Bind UI events
    document.getElementById("sort-btn").addEventListener("click", onSortClick);
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
        return;
      }

      _currentSection = result;
      showSection(result);
    } catch (e) {
      showNoSection("Place your cursor in a noun-group table.");
      _currentSection = null;
    }
  }

  function showSection(section) {
    document.getElementById("no-section").classList.add("hidden");
    document.getElementById("current-section").classList.remove("hidden");
    document.getElementById("category-name").textContent = section.pageTitle;
    document.getElementById("section-type").textContent = section.sectionType;

    var genderLabel = { m: "Masculine", f: "Feminine", n: "Neuter" };
    document.getElementById("main-gender").textContent =
      genderLabel[section.mainGender] || "";

    document.getElementById("sort-btn").disabled = false;
  }

  function showNoSection(message) {
    document.getElementById("no-section").classList.remove("hidden");
    document.getElementById("no-section-msg").textContent = message;
    document.getElementById("current-section").classList.add("hidden");
    document.getElementById("sort-btn").disabled = true;
    document.getElementById("uncolored-warning").classList.add("hidden");
  }

  // ── Sort Operation ──

  async function onSortClick() {
    if (!_currentSection) return;

    var sortBtn = document.getElementById("sort-btn");
    sortBtn.disabled = true;
    setStatus("Reading table...");

    try {
      // 1. Read items from the table
      var tableData = await GF.Document.readTableItems();
      if (tableData.error) {
        setStatus("Error: " + tableData.error);
        sortBtn.disabled = false;
        return;
      }

      // 2. Check for uncolored words
      if (tableData.uncolored.length > 0) {
        showUncoloredWarning(tableData.uncolored);
        setStatus("Fix uncolored words before sorting.");
        sortBtn.disabled = false;
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
          // Comments exist in the table but capture failed
          setStatus(
            "Warning: " +
              tableCommentCount +
              " comment(s) detected but could not be captured. Sort aborted to protect comments."
          );
          sortBtn.disabled = false;
          return;
        }
      } else if (tableCommentCount === -1) {
        // Comments API not available — warn but allow sort
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
        setStatus("Already in correct order. No changes needed.");
        sortBtn.disabled = false;
        return;
      }

      // 7. Delete original comments (before rewriting cells)
      if (commentCount > 0) {
        setStatus("Removing original comments...");
        await GF.Comments.deleteOriginalComments(commentMap);
      }

      // 8. Write sorted items back
      setStatus("Writing sorted items...");
      await GF.Document.writeTableItems(sorted, tableData.numGroups);

      // 9. Restore comments in new positions
      if (commentCount > 0) {
        setStatus("Restoring comments...");
        await GF.Comments.restoreComments(commentMap);
      }

      setStatus(
        "Sorted " +
          sorted.length +
          " items. Use Ctrl+Z to undo." +
          (commentCount > 0
            ? " " + commentCount + " comment(s) preserved."
            : "")
      );
    } catch (e) {
      console.error("Sort error:", e);
      setStatus("Error: " + e.message);
    }

    sortBtn.disabled = false;
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
        setStatus('"' + word + '" is already in the monos list (' + beforeCount + ' words).');
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
        'Added "' + word + '": ' + beforeCount + " → " + _monosCount + " monos."
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
