/**
 * sorting.js — Sorting logic + color-to-gender mapping
 * Ported from process_noun_groups.py. No Office.js dependencies.
 */
var GF = GF || {};

// =============================================================================
// COLOR-TO-GENDER MAPPING
// =============================================================================

GF.ColorMap = (function () {
  "use strict";

  // Known gender colors from the document (RGB values)
  var GENDER_COLORS = {
    masculine: [
      { r: 49, g: 103, b: 210 }, // #3167D2
      { r: 20, g: 104, b: 217 }, // #1468D9
      { r: 20, g: 105, b: 217 }, // #1469D9
    ],
    feminine: [
      { r: 185, g: 51, b: 42 }, // #B9332A
      { r: 201, g: 31, b: 31 }, // #C91F1F
    ],
    neuter: [
      { r: 64, g: 125, b: 77 }, // #407D4D
      { r: 33, g: 127, b: 71 }, // #217F47
      { r: 34, g: 127, b: 72 }, // #227F48
    ],
    plural: [
      { r: 159, g: 84, b: 185 }, // #9F54B9
    ],
  };

  var GENDER_CODES = {
    masculine: "m",
    feminine: "f",
    neuter: "n",
    plural: "p",
  };

  function hexToRgb(hex) {
    if (!hex) return null;
    hex = hex.replace("#", "");
    if (hex.length !== 6) return null;
    return {
      r: parseInt(hex.substr(0, 2), 16),
      g: parseInt(hex.substr(2, 2), 16),
      b: parseInt(hex.substr(4, 2), 16),
    };
  }

  function colorDistance(c1, c2) {
    return Math.sqrt(
      Math.pow(c1.r - c2.r, 2) +
        Math.pow(c1.g - c2.g, 2) +
        Math.pow(c1.b - c2.b, 2)
    );
  }

  /**
   * Map a hex color string to a gender name.
   * Returns 'masculine', 'feminine', 'neuter', 'plural', or null.
   */
  function identifyGender(hexColor) {
    var rgb = hexToRgb(hexColor);
    if (!rgb) return null;

    var TOLERANCE = 40;
    var bestMatch = null;
    var bestDist = Infinity;

    for (var gender in GENDER_COLORS) {
      var colors = GENDER_COLORS[gender];
      for (var i = 0; i < colors.length; i++) {
        var dist = colorDistance(rgb, colors[i]);
        if (dist < bestDist && dist <= TOLERANCE) {
          bestDist = dist;
          bestMatch = gender;
        }
      }
    }

    return bestMatch;
  }

  /**
   * Map gender name to single-letter code: 'm', 'f', 'n', 'p', or null.
   */
  function genderToCode(genderName) {
    return GENDER_CODES[genderName] || null;
  }

  return {
    identifyGender: identifyGender,
    genderToCode: genderToCode,
    hexToRgb: hexToRgb,
  };
})();

// =============================================================================
// SORTING LOGIC
// =============================================================================

GF.Sorting = (function () {
  "use strict";

  // Endings sorted longest-first (critical for matching priority)
  var ENDINGS = [
    { ending: "schaft", gender: "feminine" },
    { ending: "heit", gender: "feminine" },
    { ending: "keit", gender: "feminine" },
    { ending: "ling", gender: "masculine" },
    { ending: "chen", gender: "neuter" },
    { ending: "tät", gender: "feminine" },
    { ending: "enz", gender: "feminine" },
    { ending: "ung", gender: "feminine" },
    { ending: "anz", gender: "feminine" },
    { ending: "ion", gender: "feminine" },
    { ending: "nis", gender: "neuter" },
    { ending: "or", gender: "masculine" },
    { ending: "us", gender: "masculine" },
    { ending: "um", gender: "neuter" },
    { ending: "in", gender: "feminine" },
    { ending: "ik", gender: "feminine" },
    { ending: "ei", gender: "feminine" },
    { ending: "ur", gender: "feminine" },
    { ending: "ie", gender: "feminine" },
    { ending: "el", gender: "masculine" },
    { ending: "en", gender: "masculine" },
    { ending: "er", gender: "masculine" },
    { ending: "e", gender: "feminine" },
  ];

  // If word ends with exception, skip the base ending match
  var ENDING_EXCEPTIONS = {
    en: ["chen"],
    e: ["ie"],
  };

  // Ge- prefix overrides these weak endings
  var WEAK_ENDINGS = { el: true, en: true, er: true, e: true };

  // Currency symbols in translations
  var CURRENCY_TRANSLATIONS = {
    "\u0E3F": true, // ฿
    "\u00A2": true, // ¢
    "\u20AC": true, // €
    Fr: true,
    $: true,
    p: true,
    "\u00A5": true, // ¥
    "\u00A3": true, // £
    "\u20BA": true, // ₺
    Rp: true,
  };

  /**
   * Check if a word has an ending on Laura's list.
   * Returns the ending string if found, null otherwise.
   */
  function findEnding(word) {
    var lower = word.toLowerCase();

    for (var i = 0; i < ENDINGS.length; i++) {
      var ending = ENDINGS[i].ending;
      if (lower.length > ending.length && lower.endsWith(ending)) {
        // Check exceptions
        var exceptions = ENDING_EXCEPTIONS[ending];
        if (exceptions) {
          var skip = false;
          for (var j = 0; j < exceptions.length; j++) {
            if (lower.endsWith(exceptions[j])) {
              skip = true;
              break;
            }
          }
          if (skip) continue;
        }

        // Ge- prefix overrides weak endings
        if (word.startsWith("Ge") && WEAK_ENDINGS[ending]) {
          return null;
        }

        return ending;
      }
    }

    return null;
  }

  function isCurrency(translation) {
    return CURRENCY_TRANSLATIONS[translation.trim()] === true;
  }

  /**
   * Compute gender sort position.
   * Normal order: m=0, n=1, f=2
   * mainLast: main gender sorts to position 3
   * mainFirst: main gender sorts to position -1
   */
  function genderOrder(gender, mainGender, mainLast, mainFirst) {
    var base = { m: 0, n: 1, f: 2 };
    if (mainLast && gender === mainGender) return 3;
    if (mainFirst && gender === mainGender) return -1;
    return base[gender] !== undefined ? base[gender] : 2;
  }

  function predictableSortKey(item, mainGender, monosSet) {
    return [
      isCurrency(item.translation) ? 1 : 0,
      !monosSet.has(item.word) && findEnding(item.word) !== null ? 0 : 1,
      genderOrder(item.gender, mainGender, true, false),
      monosSet.has(item.word) ? 1 : 0,
      item.word.toLowerCase(),
    ];
  }

  function unpredictableSortKey(item, mainGender, monosSet) {
    return [
      isCurrency(item.translation) ? 1 : 0,
      genderOrder(item.gender, mainGender, false, true),
      monosSet.has(item.word) ? 1 : 0,
      item.word.toLowerCase(),
    ];
  }

  function compareKeys(a, b) {
    for (var i = 0; i < a.length; i++) {
      if (a[i] < b[i]) return -1;
      if (a[i] > b[i]) return 1;
    }
    return 0;
  }

  // ── Custom sort: Fractions (sort by numeric fraction value, smallest→largest) ──

  /**
   * Parse a fraction string like "1/8", "3/4", or "1 1/2" into a numeric value.
   * Returns NaN if unparseable.
   */
  function parseFraction(text) {
    text = text.trim();
    // Mixed number: "1 1/2" → 1 + 0.5
    var mixedMatch = text.match(/^(\d+)\s+(\d+)\/(\d+)$/);
    if (mixedMatch) {
      return (
        parseInt(mixedMatch[1], 10) +
        parseInt(mixedMatch[2], 10) / parseInt(mixedMatch[3], 10)
      );
    }
    // Simple fraction: "3/4" → 0.75
    var fracMatch = text.match(/^(\d+)\/(\d+)$/);
    if (fracMatch) {
      return parseInt(fracMatch[1], 10) / parseInt(fracMatch[2], 10);
    }
    // Plain number fallback
    var num = parseFloat(text);
    return isNaN(num) ? Infinity : num;
  }

  function fractionSortKey(item) {
    return [parseFraction(item.translation)];
  }

  // ── Custom sort: Numbers (sort by numeric value; "(number)" first) ──

  /**
   * Parse a number translation like "8", "80,000", "1,000,000,000".
   * Returns NaN if unparseable.
   */
  function parseNumber(text) {
    text = text.trim().replace(/,/g, "");
    var num = parseFloat(text);
    return isNaN(num) ? Infinity : num;
  }

  function numberSortKey(item) {
    var trans = item.translation.trim().toLowerCase();
    // "(number)" or "number" sorts first
    if (trans === "number" || trans === "(number)") {
      return [-1];
    }
    return [parseNumber(item.translation)];
  }

  // ── Custom sort: Gerunds (alphabetical, empty translations last) ──

  function gerundSortKey(item) {
    var hasTranslation = item.translation && item.translation.trim() !== "" ? 0 : 1;
    return [hasTranslation, item.word.toLowerCase()];
  }

  // Categories with custom sort logic (bypass standard Predictable/Unpredictable sort)
  var CUSTOM_SORT_CATEGORIES = {
    Fractions: fractionSortKey,
    Numbers: numberSortKey,
    Gerunds: gerundSortKey,
  };

  // Default alphabetical sort for flat-list categories without custom sort
  function alphabeticalSortKey(item) {
    return [item.word.toLowerCase()];
  }

  /**
   * Sort items according to the section type's hierarchy.
   * For flat-list categories: uses custom sort if defined, else alphabetical.
   * For sectioned categories: uses Predictable/Unpredictable sort hierarchy.
   * Returns a new sorted array (does not mutate input).
   */
  function sortItems(items, sectionType, mainGender, monosSet, categoryName) {
    var customKeyFn = categoryName && CUSTOM_SORT_CATEGORIES[categoryName];
    if (customKeyFn) {
      return items.slice().sort(function (a, b) {
        return compareKeys(customKeyFn(a), customKeyFn(b));
      });
    }

    // Flat-list category with no custom sort — sort alphabetically
    if (!sectionType) {
      return items.slice().sort(function (a, b) {
        return compareKeys(alphabeticalSortKey(a), alphabeticalSortKey(b));
      });
    }

    var keyFn =
      sectionType === "Predictable" ? predictableSortKey : unpredictableSortKey;

    return items.slice().sort(function (a, b) {
      return compareKeys(
        keyFn(a, mainGender, monosSet),
        keyFn(b, mainGender, monosSet)
      );
    });
  }

  /**
   * Balanced item distribution across columns.
   * e.g. 7 items, 4 cols -> [2, 2, 2, 1]
   */
  function distributeItems(totalItems, numGroups) {
    var base = Math.floor(totalItems / numGroups);
    var remainder = totalItems % numGroups;
    var dist = [];
    for (var g = 0; g < numGroups; g++) {
      dist.push(base + (g < remainder ? 1 : 0));
    }
    return dist;
  }

  /**
   * Estimate text width in points using per-character width ratios for Avenir Book.
   * Ported from Python components.py estimate_text_width_pt().
   */
  function estimateTextWidthPt(text, fontSizePt) {
    fontSizePt = fontSizePt || 11;
    var width = 0;
    for (var i = 0; i < text.length; i++) {
      var ch = text[i];
      if (ch === "m" || ch === "w" || ch === "M" || ch === "W") {
        width += 0.75;
      } else if (ch >= "A" && ch <= "Z") {
        width += 0.65;
      } else if ("il|!.,;:'\"()".indexOf(ch) !== -1) {
        width += 0.3;
      } else if ("fjrt".indexOf(ch) !== -1) {
        width += 0.4;
      } else if (ch === " ") {
        width += 0.25;
      } else if ("\u00F6\u00E4\u00FC\u00D6\u00C4\u00DC\u00DF".indexOf(ch) !== -1) {
        // öäüÖÄÜß
        width += 0.6;
      } else {
        width += 0.55;
      }
    }
    return width * fontSizePt;
  }

  return {
    findEnding: findEnding,
    isCurrency: isCurrency,
    sortItems: sortItems,
    distributeItems: distributeItems,
    estimateTextWidthPt: estimateTextWidthPt,
    ENDINGS: ENDINGS,
  };
})();
