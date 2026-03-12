/**
 * document.js — Section detection + table read/write using Word JS API.
 * Handles reading items from tables, detecting the current section,
 * writing sorted items back, column layout computation, page-level
 * alignment, and column count conversion.
 */
var GF = GF || {};

GF.Document = (function () {
  "use strict";

  // Categories that support sorting (standard noun-group sections)
  var SORTABLE_CATEGORIES = {
    Animals: "m",
    Beverages: "m",
    Fabrics: "m",
    Finance: "m",
    Landscape: "m",
    Plants: "m",
    Rocks: "m",
    Weather: "m",
    "Chemical Elements": "n",
    Measurements: "n",
    Colors: "n",
    Gerunds: "n",
    "Metals & Materials": "n",
    "Fruits & Nuts": "f",
    Numbers: "f",
  };

  var CATEGORY_NAMES = Object.keys(SORTABLE_CATEGORIES);

  // ── Layout constants ──
  var DEFAULT_FONT_SIZE = 11;
  var MIN_FONT_SIZE = 8;
  var CELL_PADDING_PT = 8; // estimated cell padding (left + right margins)
  var WIDTH_SAFETY = 1.12; // safety factor for width estimation
  var TABLE_WIDTH_PT = 525.6; // 7.3 inches * 72 pt/inch

  // =========================================================================
  // SECTION DETECTION (unchanged)
  // =========================================================================

  /**
   * Detect the current section based on cursor position.
   * Returns: { pageTitle, sectionType, mainGender } or { error }.
   */
  async function detectSection() {
    return Word.run(async function (context) {
      var selection = context.document.getSelection();
      var parentTable = selection.parentTableOrNullObject;
      await context.sync();

      if (parentTable.isNullObject) {
        return { error: "Cursor is not inside a table." };
      }

      // Get range from document start to this table
      var tableStart = parentTable.getRange("Start");
      var beforeTable = context.document.body
        .getRange("Start")
        .expandTo(tableStart);

      // Search for "Predictable" and "Unpredictable" before the table
      var predResults = beforeTable.search("Predictable", {
        matchWholeWord: true,
        matchCase: true,
      });
      var unpredResults = beforeTable.search("Unpredictable", {
        matchWholeWord: true,
        matchCase: true,
      });
      predResults.load("items");
      unpredResults.load("items");
      await context.sync();

      // Find the closest heading (last match = closest to the table)
      var lastPred =
        predResults.items.length > 0
          ? predResults.items[predResults.items.length - 1]
          : null;
      var lastUnpred =
        unpredResults.items.length > 0
          ? unpredResults.items[unpredResults.items.length - 1]
          : null;

      var sectionType = null;

      if (lastPred && lastUnpred) {
        var comp = lastPred.compareLocationWith(lastUnpred);
        await context.sync();
        sectionType = comp.value === "After" ? "Predictable" : "Unpredictable";
      } else if (lastPred) {
        sectionType = "Predictable";
      } else if (lastUnpred) {
        sectionType = "Unpredictable";
      } else {
        return { error: "No section heading found before this table." };
      }

      // Find the category name by searching before the section heading
      var headingRange =
        sectionType === "Predictable" ? lastPred : lastUnpred;
      var beforeHeading = context.document.body
        .getRange("Start")
        .expandTo(headingRange.getRange("Start"));

      // Search for each known category name
      var searches = {};
      for (var i = 0; i < CATEGORY_NAMES.length; i++) {
        var cat = CATEGORY_NAMES[i];
        searches[cat] = beforeHeading.search(cat, { matchCase: false });
        searches[cat].load("items");
      }
      await context.sync();

      // Find the closest category (last match among all searches)
      var candidates = [];
      for (var cat in searches) {
        var results = searches[cat];
        if (results.items.length > 0) {
          candidates.push({
            category: cat,
            range: results.items[results.items.length - 1],
          });
        }
      }

      if (candidates.length === 0) {
        return { error: "No category title found. Is this a noun-group page?" };
      }

      // Find the latest candidate (closest to the heading)
      var closest = candidates[0];
      for (var c = 1; c < candidates.length; c++) {
        var cmp = candidates[c].range.compareLocationWith(closest.range);
        await context.sync();
        if (cmp.value === "After") {
          closest = candidates[c];
        }
      }

      return {
        pageTitle: closest.category,
        sectionType: sectionType,
        mainGender: SORTABLE_CATEGORIES[closest.category],
      };
    });
  }

  // =========================================================================
  // TABLE READING
  // =========================================================================

  /**
   * Read all word-translation items from the table the cursor is in.
   * Returns: { items: [{word, translation, gender, color}], numGroups, uncolored }
   */
  async function readTableItems() {
    return Word.run(async function (context) {
      var selection = context.document.getSelection();
      var parentTable = selection.parentTableOrNullObject;
      await context.sync();

      if (parentTable.isNullObject) {
        return { error: "Cursor is not inside a table." };
      }

      // Load table rows
      var rows = parentTable.rows;
      rows.load("items");
      await context.sync();

      // Load cells for each row
      for (var r = 0; r < rows.items.length; r++) {
        rows.items[r].cells.load("items");
      }
      await context.sync();

      var numCols = rows.items[0].cells.items.length;
      var numGroups = Math.floor(numCols / 2);

      // Load text and font color for each cell
      var cellRanges = [];
      for (var r = 0; r < rows.items.length; r++) {
        var rowRanges = [];
        for (var c = 0; c < numCols; c++) {
          var range = rows.items[r].cells.items[c].body.getRange("Whole");
          range.load("text");
          range.font.load("color");
          rowRanges.push(range);
        }
        cellRanges.push(rowRanges);
      }
      await context.sync();

      // Extract items from word columns (even indices)
      var items = [];
      var uncolored = [];

      for (var r = 0; r < cellRanges.length; r++) {
        for (var g = 0; g < numGroups; g++) {
          var wordRange = cellRanges[r][g * 2];
          var transRange = cellRanges[r][g * 2 + 1];

          var wordText = (wordRange.text || "").replace(/\r/g, "").trim();
          if (!wordText) continue;

          var colorHex = wordRange.font.color;
          var genderName = GF.ColorMap.identifyGender(colorHex);
          var genderCode = GF.ColorMap.genderToCode(genderName);

          if (!genderCode) {
            uncolored.push(wordText);
          }

          var transText = (transRange.text || "").replace(/\r/g, "").trim();
          // Strip parentheses if present
          if (transText.startsWith("(") && transText.endsWith(")")) {
            transText = transText.slice(1, -1);
          }

          items.push({
            word: wordText,
            translation: transText,
            gender: genderCode || "m", // fallback; will be flagged
            color: colorHex,
          });
        }
      }

      return {
        items: items,
        numGroups: numGroups,
        uncolored: uncolored,
      };
    });
  }

  /**
   * Get the column count (number of column groups) of the current table.
   */
  async function getColumnCount() {
    return Word.run(async function (context) {
      var selection = context.document.getSelection();
      var parentTable = selection.parentTableOrNullObject;
      await context.sync();
      if (parentTable.isNullObject) return null;

      var firstRow = parentTable.rows.getFirst();
      firstRow.load("cellCount");
      await context.sync();

      return Math.floor(firstRow.cellCount / 2);
    });
  }

  // =========================================================================
  // LEGACY WRITE (kept for backward compatibility)
  // =========================================================================

  /**
   * Write sorted items back to the table the cursor is in.
   * Redistributes items across column groups and handles row count changes.
   * NOTE: Use sortWithAlignment() for new code — it adds column width
   * computation and page-level alignment.
   */
  async function writeTableItems(sortedItems, numGroups) {
    return Word.run(async function (context) {
      var selection = context.document.getSelection();
      var parentTable = selection.parentTableOrNullObject;
      await context.sync();

      if (parentTable.isNullObject) {
        throw new Error("Cursor is no longer inside a table.");
      }

      // Distribute items across columns
      var dist = GF.Sorting.distributeItems(sortedItems.length, numGroups);
      var groups = [];
      var idx = 0;
      for (var g = 0; g < numGroups; g++) {
        groups.push(sortedItems.slice(idx, idx + dist[g]));
        idx += dist[g];
      }

      var maxRows = Math.max.apply(null, dist);

      // Load current rows
      var rows = parentTable.rows;
      rows.load("items");
      await context.sync();

      var currentRowCount = rows.items.length;

      // Add rows if needed
      if (maxRows > currentRowCount) {
        for (var i = 0; i < maxRows - currentRowCount; i++) {
          parentTable.addRows("End", 1);
        }
        await context.sync();
        rows = parentTable.rows;
        rows.load("items");
        await context.sync();
      }

      // Load cells
      for (var r = 0; r < Math.min(maxRows, rows.items.length); r++) {
        rows.items[r].cells.load("items");
      }
      await context.sync();

      // Write data to cells
      for (var r = 0; r < maxRows; r++) {
        var row = rows.items[r];
        for (var g = 0; g < numGroups; g++) {
          var wordCell = row.cells.items[g * 2];
          var transCell = row.cells.items[g * 2 + 1];

          var item = groups[g][r];
          if (item) {
            wordCell.body.clear();
            var wordPara = wordCell.body.paragraphs.getFirst();
            var wordRange = wordPara.insertText(item.word, "Start");
            wordRange.font.color = item.color;
            wordRange.font.name = "Avenir Book";
            wordRange.font.size = 11;
            wordRange.font.bold = false;
            wordRange.font.italic = false;
            wordPara.alignment = "Right";
            wordPara.spaceAfter = 0;
            wordPara.spaceBefore = 0;

            transCell.body.clear();
            var transPara = transCell.body.paragraphs.getFirst();
            var transRange = transPara.insertText(
              "(" + item.translation + ")",
              "Start"
            );
            transRange.font.color = "#555555";
            transRange.font.name = "Avenir Book";
            transRange.font.size = 11;
            transRange.font.bold = false;
            transRange.font.italic = false;
            transPara.alignment = "Left";
            transPara.spaceAfter = 0;
            transPara.spaceBefore = 0;
          } else {
            wordCell.body.clear();
            transCell.body.clear();
          }
        }
      }

      // Delete extra rows at the bottom (if item count decreased)
      if (currentRowCount > maxRows) {
        rows = parentTable.rows;
        rows.load("items");
        await context.sync();

        for (var r = rows.items.length - 1; r >= maxRows; r--) {
          rows.items[r].delete();
        }
      }

      await context.sync();
    });
  }

  // =========================================================================
  // PAGE TABLE DETECTION (internal)
  // =========================================================================

  /**
   * Find all tables on the same page as currentTable with the given column count.
   * Uses manual page break (^m) and section break (^b) detection.
   * Falls back to all matching tables if page breaks can't be found.
   */
  async function _findPageTables(context, currentTable, numGroups) {
    var body = context.document.body;
    var allTables = body.tables;
    allTables.load("items");
    await context.sync();

    if (allTables.items.length <= 1) return allTables.items;

    // Load first row cell count for each table
    var firstRows = [];
    for (var i = 0; i < allTables.items.length; i++) {
      var firstRow = allTables.items[i].rows.getFirst();
      firstRow.load("cellCount");
      firstRows.push(firstRow);
    }
    await context.sync();

    // Filter tables by column count
    var targetCols = numGroups * 2;
    var matchingTables = [];
    for (var i = 0; i < allTables.items.length; i++) {
      if (firstRows[i].cellCount === targetCols) {
        matchingTables.push(allTables.items[i]);
      }
    }

    if (matchingTables.length <= 1) return matchingTables;

    // Search for page breaks (manual + section breaks)
    var pb1, pb2;
    try {
      pb1 = body.search("^m");
      pb2 = body.search("^b");
      pb1.load("items");
      pb2.load("items");
      await context.sync();
    } catch (e) {
      // Page break search not supported — return all matching tables
      return matchingTables;
    }

    var allBreaks = pb1.items.concat(pb2.items);
    if (allBreaks.length === 0) {
      // No page breaks found — return all matching tables
      return matchingTables;
    }

    // Get ranges for comparison
    var tableRanges = [];
    for (var i = 0; i < matchingTables.length; i++) {
      tableRanges.push(matchingTables[i].getRange("Whole"));
    }
    var pbRanges = [];
    for (var i = 0; i < allBreaks.length; i++) {
      pbRanges.push(allBreaks[i].getRange());
    }
    var currentRange = currentTable.getRange("Whole");
    await context.sync();

    // Batch all comparisons: each table vs each page break
    var comparisons = [];
    for (var t = 0; t < matchingTables.length; t++) {
      comparisons[t] = [];
      for (var pb = 0; pb < pbRanges.length; pb++) {
        comparisons[t][pb] = tableRanges[t].compareLocationWith(pbRanges[pb]);
      }
    }
    var currentComps = [];
    for (var pb = 0; pb < pbRanges.length; pb++) {
      currentComps[pb] = currentRange.compareLocationWith(pbRanges[pb]);
    }
    await context.sync();

    // Count how many page breaks come before each table to determine its "page"
    function getPageNumber(comps) {
      var page = 0;
      for (var i = 0; i < comps.length; i++) {
        var val = comps[i].value;
        if (val === "After" || val === "AdjacentAfter") {
          page++;
        }
      }
      return page;
    }

    var currentPage = getPageNumber(currentComps);
    var samePageTables = [];
    for (var t = 0; t < matchingTables.length; t++) {
      if (getPageNumber(comparisons[t]) === currentPage) {
        samePageTables.push(matchingTables[t]);
      }
    }

    return samePageTables;
  }

  // =========================================================================
  // TABLE GROUP READING (internal — for width computation)
  // =========================================================================

  /**
   * Read items from a table grouped by column position. Only needs text
   * (no color/gender) since this is used for width computation.
   */
  async function _readTableGroups(context, table) {
    var rows = table.rows;
    rows.load("items");
    await context.sync();

    for (var r = 0; r < rows.items.length; r++) {
      rows.items[r].cells.load("items");
    }
    await context.sync();

    var numCols = rows.items[0].cells.items.length;
    var numGroups = Math.floor(numCols / 2);

    // Batch-load all cell text
    var allRanges = [];
    for (var r = 0; r < rows.items.length; r++) {
      var rowRanges = [];
      for (var c = 0; c < numCols; c++) {
        var range = rows.items[r].cells.items[c].body.getRange("Whole");
        range.load("text");
        rowRanges.push(range);
      }
      allRanges.push(rowRanges);
    }
    await context.sync();

    var groups = [];
    for (var g = 0; g < numGroups; g++) {
      var groupItems = [];
      for (var r = 0; r < allRanges.length; r++) {
        var wordText = (allRanges[r][g * 2].text || "").replace(/\r/g, "").trim();
        if (!wordText) continue;
        var transText = (allRanges[r][g * 2 + 1].text || "")
          .replace(/\r/g, "")
          .trim();
        if (transText.startsWith("(") && transText.endsWith(")")) {
          transText = transText.slice(1, -1);
        }
        groupItems.push({ word: wordText, translation: transText });
      }
      groups.push(groupItems);
    }

    return { groups: groups, numGroups: numGroups };
  }

  // =========================================================================
  // COLUMN LAYOUT COMPUTATION (pure functions — no API calls)
  // =========================================================================

  /**
   * Find the minimum font size at which text fits within the given width.
   * Returns a size between MIN_FONT_SIZE and DEFAULT_FONT_SIZE.
   */
  function _findFitFontSize(text, availableWidthPt) {
    for (var size = DEFAULT_FONT_SIZE; size >= MIN_FONT_SIZE; size--) {
      var est =
        GF.Sorting.estimateTextWidthPt(text, size) * WIDTH_SAFETY +
        CELL_PADDING_PT;
      if (est <= availableWidthPt) {
        return size;
      }
    }
    return MIN_FONT_SIZE;
  }

  /**
   * Estimate the minimum column width needed for text at the given font size.
   */
  function _neededWidth(text, fontSize) {
    return (
      GF.Sorting.estimateTextWidthPt(text, fontSize) * WIDTH_SAFETY +
      CELL_PADDING_PT
    );
  }

  /**
   * Compute optimal column layout for a set of tables on the same page.
   *
   * allTableGroups: [{ groups: [[{word,translation},...], ...] }, ...]
   * numGroups: 3 or 4
   *
   * Returns: {
   *   setWidths: [pt per column set],
   *   wordWidths: [pt per word column],
   *   transWidths: [pt per translation column]
   * }
   */
  function _computeColumnLayout(allTableGroups, numGroups) {
    var tableWidthPt = TABLE_WIDTH_PT;

    // For each column position, find max word and trans width across all tables
    var maxWordPt = [];
    var maxTransPt = [];
    for (var g = 0; g < numGroups; g++) {
      var mw = 0;
      var mt = 0;
      for (var t = 0; t < allTableGroups.length; t++) {
        var items = allTableGroups[t].groups[g] || [];
        for (var i = 0; i < items.length; i++) {
          var ww = _neededWidth(items[i].word, DEFAULT_FONT_SIZE);
          var tw = _neededWidth(
            "(" + items[i].translation + ")",
            DEFAULT_FONT_SIZE
          );
          mw = Math.max(mw, ww);
          mt = Math.max(mt, tw);
        }
      }
      maxWordPt.push(mw);
      maxTransPt.push(mt);
    }

    // Total content width (tight-fit, no visual gaps)
    var totalContent = 0;
    for (var g = 0; g < numGroups; g++) {
      totalContent += maxWordPt[g] + maxTransPt[g];
    }

    var totalGap = tableWidthPt - totalContent;
    var wordWidths, transWidths, setWidths;

    if (totalGap <= 0) {
      // Content exceeds table width — distribute proportionally.
      // Font shrinking will handle the overflow.
      wordWidths = [];
      transWidths = [];
      setWidths = [];
      for (var g = 0; g < numGroups; g++) {
        var ratio = (maxWordPt[g] + maxTransPt[g]) / totalContent;
        var setW = ratio * tableWidthPt;
        var wordRatio = maxWordPt[g] / (maxWordPt[g] + maxTransPt[g]);
        var ww = wordRatio * setW;
        wordWidths.push(ww);
        transWidths.push(setW - ww);
        setWidths.push(setW);
      }
    } else {
      // Size each column to its content, then distribute remaining space
      // as equal visual gaps between column sets, centered on the page.
      //
      // N column sets → N-1 inner gaps + 2 outer margins.
      // outerMargin = interGap / 2 centers the group.
      // Visual gap between adjacent sets = interGap (half absorbed by each side).
      var interGap = totalGap / numGroups;
      var outerMargin = interGap / 2;

      // Start with content-fitted widths
      wordWidths = maxWordPt.slice();
      transWidths = maxTransPt.slice();

      // Left outer margin → extra space in first word column (right-aligned text)
      wordWidths[0] += outerMargin;
      // Right outer margin → extra space in last trans column (left-aligned text)
      transWidths[numGroups - 1] += outerMargin;

      // Inner gaps: split equally between adjacent trans and word columns
      for (var g = 0; g < numGroups - 1; g++) {
        transWidths[g] += interGap / 2;
        wordWidths[g + 1] += interGap / 2;
      }

      setWidths = [];
      for (var g = 0; g < numGroups; g++) {
        setWidths.push(wordWidths[g] + transWidths[g]);
      }
    }

    return {
      setWidths: setWidths,
      wordWidths: wordWidths,
      transWidths: transWidths,
    };
  }

  /**
   * Attempt to resolve column overflow by widening sets that need more space.
   * Falls back to proportional distribution if cascading overflow occurs.
   */
  function _resolveOverflow(
    numGroups,
    tableWidthPt,
    equalSetWidth,
    neededAtDefault,
    allTableGroups
  ) {
    var setWidths = [];
    for (var g = 0; g < numGroups; g++) {
      setWidths.push(equalSetWidth);
    }

    // Identify overflow sets
    var overflowSets = [];
    var totalOverflow = 0;
    for (var g = 0; g < numGroups; g++) {
      if (neededAtDefault[g] > equalSetWidth) {
        overflowSets.push(g);
        totalOverflow += neededAtDefault[g] - equalSetWidth;
      }
    }

    var nonOverflowCount = numGroups - overflowSets.length;

    if (nonOverflowCount === 0) {
      // All sets overflow — use proportional distribution
      return _proportionalWidths(numGroups, tableWidthPt, neededAtDefault);
    }

    // Can non-overflow sets absorb the extra? (min 60% of equal width)
    var minNonOverflow = equalSetWidth * 0.6;
    var maxShrinkPerNonOverflow = equalSetWidth - minNonOverflow;
    var maxTotalShrink = maxShrinkPerNonOverflow * nonOverflowCount;

    if (totalOverflow <= maxTotalShrink) {
      // Absorb: give overflow sets what they need, shrink others equally
      var shrinkEach = totalOverflow / nonOverflowCount;
      for (var i = 0; i < overflowSets.length; i++) {
        setWidths[overflowSets[i]] = neededAtDefault[overflowSets[i]];
      }
      for (var g = 0; g < numGroups; g++) {
        if (overflowSets.indexOf(g) === -1) {
          setWidths[g] = equalSetWidth - shrinkEach;
        }
      }

      // Check for cascading overflow
      var cascading = false;
      for (var g = 0; g < numGroups; g++) {
        if (overflowSets.indexOf(g) === -1) {
          // Check if any content in this column now overflows at MIN_FONT_SIZE
          var minNeeded = _minNeededWidthForColumn(g, allTableGroups);
          if (minNeeded > setWidths[g]) {
            cascading = true;
            break;
          }
        }
      }

      if (cascading) {
        return _proportionalWidths(numGroups, tableWidthPt, neededAtDefault);
      }

      return setWidths;
    }

    // Can't absorb — use proportional distribution
    return _proportionalWidths(numGroups, tableWidthPt, neededAtDefault);
  }

  /**
   * Compute the minimum width needed for a column at MIN_FONT_SIZE.
   * This is the absolute minimum — content fits at 8pt with font shrinking.
   */
  function _minNeededWidthForColumn(g, allTableGroups) {
    var maxNeeded = 0;
    for (var t = 0; t < allTableGroups.length; t++) {
      var items = allTableGroups[t].groups[g] || [];
      for (var i = 0; i < items.length; i++) {
        var ww = _neededWidth(items[i].word, MIN_FONT_SIZE);
        var tw = _neededWidth(
          "(" + items[i].translation + ")",
          MIN_FONT_SIZE
        );
        maxNeeded = Math.max(maxNeeded, ww + tw);
      }
    }
    return maxNeeded;
  }

  /**
   * Distribute table width proportionally to needed widths.
   */
  function _proportionalWidths(numGroups, tableWidthPt, neededWidths) {
    var totalNeeded = 0;
    for (var g = 0; g < numGroups; g++) {
      totalNeeded += Math.max(neededWidths[g], tableWidthPt / numGroups * 0.5);
    }
    var widths = [];
    for (var g = 0; g < numGroups; g++) {
      var w = Math.max(neededWidths[g], tableWidthPt / numGroups * 0.5);
      widths.push((w / totalNeeded) * tableWidthPt);
    }
    return widths;
  }

  /**
   * Compute per-cell font shrinks for a set of table groups at the given layout.
   * Returns: shrinks[tableIdx][groupIdx] = { rowIdx: { wordFont, transFont } }
   */
  function _computeFontShrinks(allTableGroups, layout) {
    var shrinks = [];
    for (var t = 0; t < allTableGroups.length; t++) {
      var tableShrinks = [];
      for (var g = 0; g < layout.wordWidths.length; g++) {
        var groupShrinks = {};
        var items = allTableGroups[t].groups[g] || [];
        for (var i = 0; i < items.length; i++) {
          var item = items[i];
          var wordFont = _findFitFontSize(item.word, layout.wordWidths[g]);
          var transText = "(" + item.translation + ")";
          var transFont = _findFitFontSize(transText, layout.transWidths[g]);

          if (wordFont < DEFAULT_FONT_SIZE || transFont < DEFAULT_FONT_SIZE) {
            groupShrinks[i] = { wordFont: wordFont, transFont: transFont };
          }
        }
        tableShrinks.push(groupShrinks);
      }
      shrinks.push(tableShrinks);
    }
    return shrinks;
  }

  // =========================================================================
  // WRITE WITH LAYOUT (internal)
  // =========================================================================

  /**
   * Write sorted items to a table with column widths and font shrinking.
   * Used by sortWithAlignment and convertAndSort.
   *
   * context: Word.RequestContext (already inside a Word.run)
   * table: Word.Table proxy
   * sortedItems: flat array of items
   * numGroups: target column group count
   * layout: { wordWidths, transWidths } from _computeColumnLayout
   * groupShrinks: shrinks for THIS table (shrinks[tableIdx])
   */
  async function _writeItemsWithLayout(
    context,
    table,
    sortedItems,
    numGroups,
    layout,
    groupShrinks
  ) {
    // Distribute items across columns
    var dist = GF.Sorting.distributeItems(sortedItems.length, numGroups);
    var groups = [];
    var idx = 0;
    for (var g = 0; g < numGroups; g++) {
      groups.push(sortedItems.slice(idx, idx + dist[g]));
      idx += dist[g];
    }

    var maxRows = Math.max.apply(null, dist);

    // Load current rows
    var rows = table.rows;
    rows.load("items");
    await context.sync();

    var currentRowCount = rows.items.length;

    // Add rows if needed
    if (maxRows > currentRowCount) {
      for (var i = 0; i < maxRows - currentRowCount; i++) {
        table.addRows("End", 1);
      }
      await context.sync();
      rows = table.rows;
      rows.load("items");
      await context.sync();
    }

    // Load cells for all rows we'll write to
    for (var r = 0; r < Math.min(maxRows, rows.items.length); r++) {
      rows.items[r].cells.load("items");
    }
    await context.sync();

    // Set column widths on the first row
    var firstRowCells = rows.items[0].cells.items;
    for (var g = 0; g < numGroups; g++) {
      firstRowCells[g * 2].columnWidth = layout.wordWidths[g];
      firstRowCells[g * 2 + 1].columnWidth = layout.transWidths[g];
    }

    // Write data to cells with font sizing
    for (var r = 0; r < maxRows; r++) {
      var row = rows.items[r];
      for (var g = 0; g < numGroups; g++) {
        var wordCell = row.cells.items[g * 2];
        var transCell = row.cells.items[g * 2 + 1];

        var item = groups[g][r];
        if (item) {
          var shrink = groupShrinks[g] && groupShrinks[g][r];
          var wordFontSize = shrink ? shrink.wordFont : DEFAULT_FONT_SIZE;
          var transFontSize = shrink ? shrink.transFont : DEFAULT_FONT_SIZE;

          // Write word
          wordCell.body.clear();
          var wordPara = wordCell.body.paragraphs.getFirst();
          var wordRange = wordPara.insertText(item.word, "Start");
          wordRange.font.color = item.color;
          wordRange.font.name = "Avenir Book";
          wordRange.font.size = wordFontSize;
          wordRange.font.bold = false;
          wordRange.font.italic = false;
          wordPara.alignment = "Right";
          wordPara.spaceAfter = 0;
          wordPara.spaceBefore = 0;

          // Write translation
          transCell.body.clear();
          var transPara = transCell.body.paragraphs.getFirst();
          var transRange = transPara.insertText(
            "(" + item.translation + ")",
            "Start"
          );
          transRange.font.color = "#555555";
          transRange.font.name = "Avenir Book";
          transRange.font.size = transFontSize;
          transRange.font.bold = false;
          transRange.font.italic = false;
          transPara.alignment = "Left";
          transPara.spaceAfter = 0;
          transPara.spaceBefore = 0;
        } else {
          // Clear empty cells
          wordCell.body.clear();
          transCell.body.clear();
        }
      }
    }

    // Delete extra rows at the bottom (if item count decreased)
    if (currentRowCount > maxRows) {
      rows = table.rows;
      rows.load("items");
      await context.sync();

      for (var r = rows.items.length - 1; r >= maxRows; r--) {
        rows.items[r].delete();
      }
    }

    await context.sync();
  }

  /**
   * Apply column widths and font shrinks to an existing table (other tables
   * on the same page). Does NOT change content — only widths and font sizes.
   */
  async function _applyLayoutToTable(
    context,
    table,
    numGroups,
    layout,
    groupShrinks,
    tableGroups
  ) {
    var rows = table.rows;
    rows.load("items");
    await context.sync();

    for (var r = 0; r < rows.items.length; r++) {
      rows.items[r].cells.load("items");
    }
    await context.sync();

    // Set column widths on the first row
    var firstRowCells = rows.items[0].cells.items;
    for (var g = 0; g < numGroups; g++) {
      firstRowCells[g * 2].columnWidth = layout.wordWidths[g];
      firstRowCells[g * 2 + 1].columnWidth = layout.transWidths[g];
    }

    // Apply font sizes — reset to default, then shrink where needed
    for (var r = 0; r < rows.items.length; r++) {
      for (var g = 0; g < numGroups; g++) {
        var wordRange = rows.items[r].cells.items[g * 2].body.getRange("Whole");
        var transRange =
          rows.items[r].cells.items[g * 2 + 1].body.getRange("Whole");

        // Check if this row has an item (not an empty cell)
        var itemIdx = r; // row index = item index within the group
        var items = tableGroups.groups[g] || [];
        if (itemIdx >= items.length) continue; // empty cell

        var shrink = groupShrinks[g] && groupShrinks[g][itemIdx];
        wordRange.font.size = shrink ? shrink.wordFont : DEFAULT_FONT_SIZE;
        transRange.font.size = shrink ? shrink.transFont : DEFAULT_FONT_SIZE;
      }
    }

    await context.sync();
  }

  // =========================================================================
  // SORT WITH ALIGNMENT (main entry point)
  // =========================================================================

  /**
   * Sort items, compute page-level column layout, and apply to all
   * same-page tables with matching column count.
   */
  async function sortWithAlignment(sortedItems, numGroups) {
    return Word.run(async function (context) {
      // 1. Get current table
      var selection = context.document.getSelection();
      var currentTable = selection.parentTableOrNullObject;
      await context.sync();
      if (currentTable.isNullObject) {
        throw new Error("Cursor is no longer inside a table.");
      }

      // 2. Distribute sorted items into column groups (for width computation)
      var dist = GF.Sorting.distributeItems(sortedItems.length, numGroups);
      var currentGroups = [];
      var idx = 0;
      for (var g = 0; g < numGroups; g++) {
        currentGroups.push(sortedItems.slice(idx, idx + dist[g]));
        idx += dist[g];
      }

      // 3. Find same-page tables with matching column count
      var pageTables = await _findPageTables(context, currentTable, numGroups);

      // 4. Separate current table from other tables
      var currentRange = currentTable.getRange("Whole");
      var compResults = [];
      for (var i = 0; i < pageTables.length; i++) {
        compResults.push(
          pageTables[i].getRange("Whole").compareLocationWith(currentRange)
        );
      }
      await context.sync();

      var otherTables = [];
      for (var i = 0; i < pageTables.length; i++) {
        if (compResults[i].value !== "Equal") {
          otherTables.push(pageTables[i]);
        }
      }

      // 5. Read items from other tables for width computation
      var allTableGroups = [{ groups: currentGroups }]; // current table first
      var otherTableGroups = [];
      for (var i = 0; i < otherTables.length; i++) {
        var tg = await _readTableGroups(context, otherTables[i]);
        if (tg.numGroups === numGroups) {
          allTableGroups.push(tg);
          otherTableGroups.push(tg);
        }
      }

      // 6. Compute column layout across all page tables
      var layout = _computeColumnLayout(allTableGroups, numGroups);

      // 7. Compute font shrinks for all tables
      var shrinks = _computeFontShrinks(allTableGroups, layout);

      // 8. Write sorted items to current table with layout
      await _writeItemsWithLayout(
        context,
        currentTable,
        sortedItems,
        numGroups,
        layout,
        shrinks[0]
      );

      // 9. Apply layout to other page tables (widths + font sizes only)
      for (var i = 0; i < otherTables.length; i++) {
        await _applyLayoutToTable(
          context,
          otherTables[i],
          numGroups,
          layout,
          shrinks[i + 1],
          otherTableGroups[i]
        );
      }

      return {
        shrinkCount: _countShrinks(shrinks),
        pageTableCount: pageTables.length,
      };
    });
  }

  /**
   * Count how many cells were font-shrunk across all tables.
   */
  function _countShrinks(shrinks) {
    var count = 0;
    for (var t = 0; t < shrinks.length; t++) {
      for (var g = 0; g < shrinks[t].length; g++) {
        count += Object.keys(shrinks[t][g]).length;
      }
    }
    return count;
  }

  // =========================================================================
  // COLUMN CONVERSION
  // =========================================================================

  /**
   * Convert the current table to a different column count, sort, and align
   * with other same-column-count tables on the page.
   *
   * Deletes the existing table and creates a new one with targetGroups columns.
   * Returns the new table's range so the caller can restore comments.
   */
  async function convertAndSort(sortedItems, targetGroups) {
    return Word.run(async function (context) {
      // 1. Get current table and a reference paragraph before it
      var selection = context.document.getSelection();
      var currentTable = selection.parentTableOrNullObject;
      await context.sync();
      if (currentTable.isNullObject) {
        throw new Error("Cursor is no longer inside a table.");
      }

      // 2. Compute new table dimensions
      var dist = GF.Sorting.distributeItems(sortedItems.length, targetGroups);
      var maxRows = Math.max.apply(null, dist);
      var numCols = targetGroups * 2;

      // 3. Insert new table AFTER the current one (while it still exists)
      var afterRange = currentTable.getRange("After");
      var newTable = afterRange.insertTable(maxRows, numCols, "After");
      await context.sync();

      // 4. Delete the old table (new table reference remains valid)
      currentTable.delete();
      await context.sync();

      // 5. Remove borders on the new table
      var borderLocations = [
        "Top",
        "Bottom",
        "Left",
        "Right",
        "InsideHorizontal",
        "InsideVertical",
      ];
      for (var i = 0; i < borderLocations.length; i++) {
        var border = newTable.getBorder(borderLocations[i]);
        border.type = "None";
      }
      await context.sync();

      // 6. Find other same-column-count tables on this page for alignment
      var pageTables = await _findPageTables(context, newTable, targetGroups);

      // Distribute items into groups
      var currentGroups = [];
      var idx = 0;
      for (var g = 0; g < targetGroups; g++) {
        currentGroups.push(sortedItems.slice(idx, idx + dist[g]));
        idx += dist[g];
      }

      // Read items from other page tables
      var allTableGroups = [{ groups: currentGroups }];
      var otherTables = [];
      var otherTableGroups = [];

      // Identify other tables (exclude the new table)
      var newRange = newTable.getRange("Whole");
      var compResults = [];
      for (var i = 0; i < pageTables.length; i++) {
        compResults.push(
          pageTables[i].getRange("Whole").compareLocationWith(newRange)
        );
      }
      await context.sync();

      for (var i = 0; i < pageTables.length; i++) {
        if (compResults[i].value !== "Equal") {
          otherTables.push(pageTables[i]);
        }
      }

      for (var i = 0; i < otherTables.length; i++) {
        var tg = await _readTableGroups(context, otherTables[i]);
        if (tg.numGroups === targetGroups) {
          allTableGroups.push(tg);
          otherTableGroups.push(tg);
        }
      }

      // 7. Compute layout and shrinks
      var layout = _computeColumnLayout(allTableGroups, targetGroups);
      var shrinks = _computeFontShrinks(allTableGroups, layout);

      // 8. Write items to the new table with layout
      await _writeItemsWithLayout(
        context,
        newTable,
        sortedItems,
        targetGroups,
        layout,
        shrinks[0]
      );

      // 9. Apply layout to other page tables
      for (var i = 0; i < otherTables.length; i++) {
        await _applyLayoutToTable(
          context,
          otherTables[i],
          targetGroups,
          layout,
          shrinks[i + 1],
          otherTableGroups[i]
        );
      }

      // 10. Place cursor in the new table
      newTable.getRange("Start").select();
      await context.sync();

      return {
        shrinkCount: _countShrinks(shrinks),
        pageTableCount: pageTables.length,
      };
    });
  }

  // =========================================================================
  // PUBLIC API
  // =========================================================================

  return {
    detectSection: detectSection,
    readTableItems: readTableItems,
    writeTableItems: writeTableItems,
    getColumnCount: getColumnCount,
    sortWithAlignment: sortWithAlignment,
    convertAndSort: convertAndSort,
    SORTABLE_CATEGORIES: SORTABLE_CATEGORIES,
  };
})();
