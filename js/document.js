/**
 * document.js — Section detection + table read/write using Word JS API.
 * Handles reading items from tables, detecting the current section, and
 * writing sorted items back.
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

  /**
   * Detect the current section based on cursor position.
   * Returns: { pageTitle, sectionType, mainGender, table } or { error }.
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

  /**
   * Read all word-translation items from the table the cursor is in.
   * Returns: { items: [{word, translation, gender, color}], numGroups, table }
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
   * Write sorted items back to the table the cursor is in.
   * Redistributes items across column groups and handles row count changes.
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

      // Write data to cells — use insertText on existing paragraph
      // to avoid creating a double paragraph (clear leaves one empty
      // paragraph; insertParagraph would add a second, inflating rows)
      for (var r = 0; r < maxRows; r++) {
        var row = rows.items[r];
        for (var g = 0; g < numGroups; g++) {
          var wordCell = row.cells.items[g * 2];
          var transCell = row.cells.items[g * 2 + 1];

          var item = groups[g][r];
          if (item) {
            // Write word — clear then set text on remaining paragraph
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

            // Write translation
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
            // Clear empty cells
            wordCell.body.clear();
            transCell.body.clear();
          }
        }
      }

      // Delete extra rows at the bottom (if item count decreased)
      if (currentRowCount > maxRows) {
        // Reload rows after changes
        rows = parentTable.rows;
        rows.load("items");
        await context.sync();

        // Delete from bottom up to avoid index shifting
        for (var r = rows.items.length - 1; r >= maxRows; r--) {
          rows.items[r].delete();
        }
      }

      await context.sync();
    });
  }

  return {
    detectSection: detectSection,
    readTableItems: readTableItems,
    writeTableItems: writeTableItems,
    SORTABLE_CATEGORIES: SORTABLE_CATEGORIES,
  };
})();
