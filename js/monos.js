/**
 * monos.js — Manage the monosyllabic words list via Custom XML Parts.
 * Uses the Office Common API (works across all Office versions with add-in support).
 */
var GF = GF || {};

GF.Monos = (function () {
  "use strict";

  var NAMESPACE = "urn:german-foundations:monos";
  var _cachedMonos = null;

  // ── Promise wrappers for the callback-based Common API ──

  function getPartsByNamespace(ns) {
    return new Promise(function (resolve, reject) {
      Office.context.document.customXmlParts.getByNamespaceAsync(
        ns,
        function (result) {
          if (result.status === Office.AsyncResultStatus.Succeeded) {
            resolve(result.value);
          } else {
            reject(new Error(result.error.message));
          }
        }
      );
    });
  }

  function getPartXml(part) {
    return new Promise(function (resolve, reject) {
      part.getXmlAsync(function (result) {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          resolve(result.value);
        } else {
          reject(new Error(result.error.message));
        }
      });
    });
  }

  function addPart(xml) {
    return new Promise(function (resolve, reject) {
      Office.context.document.customXmlParts.addAsync(
        xml,
        function (result) {
          if (result.status === Office.AsyncResultStatus.Succeeded) {
            resolve(result.value);
          } else {
            reject(new Error(result.error.message));
          }
        }
      );
    });
  }

  function deletePart(part) {
    return new Promise(function (resolve, reject) {
      part.deleteAsync(function (result) {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          resolve();
        } else {
          reject(new Error(result.error.message));
        }
      });
    });
  }

  // ── XML serialization ──

  function escapeXml(str) {
    return str
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;");
  }

  function buildXml(monosSet) {
    var words = Array.from(monosSet).sort();
    var items = words
      .map(function (w) {
        return "  <Word>" + escapeXml(w) + "</Word>";
      })
      .join("\n");
    return (
      '<?xml version="1.0" encoding="UTF-8"?>\n' +
      '<Monos xmlns="' +
      NAMESPACE +
      '">\n' +
      items +
      "\n</Monos>"
    );
  }

  function parseXml(xmlString) {
    var parser = new DOMParser();
    var doc = parser.parseFromString(xmlString, "text/xml");
    var words = doc.getElementsByTagNameNS(NAMESPACE, "Word");
    var monos = new Set();
    for (var i = 0; i < words.length; i++) {
      var text = words[i].textContent.trim();
      if (text) monos.add(text);
    }
    return monos;
  }

  // ── Public API ──

  /**
   * Load the monos set from the document's Custom XML Part.
   * Returns a Set of word strings.
   */
  async function getMonosSet() {
    if (_cachedMonos) return new Set(_cachedMonos);

    try {
      var parts = await getPartsByNamespace(NAMESPACE);
      if (parts.length > 0) {
        var xml = await getPartXml(parts[0]);
        _cachedMonos = parseXml(xml);
        return new Set(_cachedMonos);
      }
    } catch (e) {
      console.warn("Failed to load monos Custom XML Part:", e);
    }

    return new Set();
  }

  /**
   * Save the full monos set, replacing any existing Custom XML Part.
   */
  async function saveMonosSet(monosSet) {
    // Delete existing parts
    var parts;
    try {
      parts = await getPartsByNamespace(NAMESPACE);
    } catch (e) {
      parts = [];
    }

    for (var i = 0; i < parts.length; i++) {
      try {
        await deletePart(parts[i]);
      } catch (e) {
        console.warn("Failed to delete existing monos part:", e);
      }
    }

    var xml = buildXml(monosSet);
    await addPart(xml);
    _cachedMonos = new Set(monosSet);
  }

  async function addWord(word) {
    var monos = await getMonosSet();
    monos.add(word);
    await saveMonosSet(monos);

    // Verify save by re-reading from document
    _cachedMonos = null;
    var verified = await getMonosSet();
    if (!verified.has(word)) {
      throw new Error(
        "Word was not persisted to document. Try saving the document first."
      );
    }
    return verified;
  }

  async function removeWord(word) {
    var monos = await getMonosSet();
    monos.delete(word);
    await saveMonosSet(monos);

    // Verify save by re-reading from document
    _cachedMonos = null;
    var verified = await getMonosSet();
    return verified;
  }

  /**
   * Import monos from CSV text (one word per line, first line may be "Monos" header).
   */
  async function importFromCsv(csvText) {
    var words = csvText
      .split("\n")
      .map(function (line) {
        return line.trim();
      })
      .filter(function (line) {
        return line && line !== "Monos";
      });
    var monos = new Set(words);
    await saveMonosSet(monos);
    return monos;
  }

  function clearCache() {
    _cachedMonos = null;
  }

  return {
    getMonosSet: getMonosSet,
    saveMonosSet: saveMonosSet,
    addWord: addWord,
    removeWord: removeWord,
    importFromCsv: importFromCsv,
    clearCache: clearCache,
    NAMESPACE: NAMESPACE,
  };
})();
