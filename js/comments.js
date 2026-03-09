/**
 * comments.js — Track and restore comments during table sorting.
 * Requires WordApi 1.4+ for comment support.
 * Falls back gracefully if comments API is unavailable.
 */
var GF = GF || {};

GF.Comments = (function () {
  "use strict";

  /**
   * Check if the Comments API is available in this Word version.
   */
  function isSupported() {
    return (
      typeof Word !== "undefined" &&
      Word.run &&
      Office.context.requirements.isSetSupported("WordApi", "1.4")
    );
  }

  /**
   * Capture all comments inside the given table.
   * Returns a map: { "wordText": [{ text, author, date }] }
   */
  async function captureComments() {
    if (!isSupported()) return {};

    return Word.run(async function (context) {
      var selection = context.document.getSelection();
      var parentTable = selection.parentTableOrNullObject;
      await context.sync();

      if (parentTable.isNullObject) return {};

      // Get comments scoped to the table range (not entire body)
      var tableRange = parentTable.getRange("Whole");
      var tableComments;
      try {
        tableComments = tableRange.getComments();
        tableComments.load("items");
        await context.sync();
      } catch (e) {
        console.warn("Comments API not available:", e);
        return {};
      }

      if (tableComments.items.length === 0) return {};

      // Load properties for each comment
      for (var i = 0; i < tableComments.items.length; i++) {
        var comment = tableComments.items[i];
        comment.load("content,authorName,creationDate,id");
      }
      await context.sync();

      // Get the range (anchored text) for each comment
      var commentData = [];
      for (var i = 0; i < tableComments.items.length; i++) {
        var comment = tableComments.items[i];
        try {
          var commentRange = comment.getRange();
          commentRange.load("text");
          commentData.push({ comment: comment, range: commentRange });
        } catch (e) {
          // Some comments may not have accessible ranges
          continue;
        }
      }
      await context.sync();

      // Build map keyed by the word the comment is attached to
      var commentMap = {};
      for (var i = 0; i < commentData.length; i++) {
        var cd = commentData[i];
        var word = (cd.range.text || "").trim();
        if (!commentMap[word]) {
          commentMap[word] = [];
        }
        commentMap[word].push({
          text: cd.comment.content,
          author: cd.comment.authorName,
          date: cd.comment.creationDate,
          id: cd.comment.id,
        });
      }

      return commentMap;
    });
  }

  /**
   * Delete original comments that were captured (by ID).
   */
  async function deleteOriginalComments(commentMap) {
    if (!isSupported() || Object.keys(commentMap).length === 0) return;

    var idsToDelete = {};
    for (var word in commentMap) {
      for (var i = 0; i < commentMap[word].length; i++) {
        idsToDelete[commentMap[word][i].id] = true;
      }
    }

    return Word.run(async function (context) {
      var allComments = context.document.body.getComments();
      allComments.load("items");
      await context.sync();

      for (var i = 0; i < allComments.items.length; i++) {
        allComments.items[i].load("id");
      }
      await context.sync();

      for (var i = 0; i < allComments.items.length; i++) {
        if (idsToDelete[allComments.items[i].id]) {
          allComments.items[i].delete();
        }
      }
      await context.sync();
    });
  }

  /**
   * Restore comments after sorting by finding words in their new positions.
   * Comments are re-created with an author attribution prefix.
   */
  async function restoreComments(commentMap) {
    if (!isSupported() || Object.keys(commentMap).length === 0) return;

    return Word.run(async function (context) {
      var selection = context.document.getSelection();
      var parentTable = selection.parentTableOrNullObject;
      await context.sync();

      if (parentTable.isNullObject) return;

      var tableRange = parentTable.getRange("Whole");

      for (var word in commentMap) {
        var comments = commentMap[word];

        // Search for the word in the table
        var results = tableRange.search(word, {
          matchWholeWord: true,
          matchCase: true,
        });
        results.load("items");
        await context.sync();

        if (results.items.length > 0) {
          var targetRange = results.items[0];

          for (var i = 0; i < comments.length; i++) {
            var c = comments[i];
            var prefix = "[" + c.author + ", " + formatDate(c.date) + "] ";
            targetRange.insertComment(prefix + c.text);
          }
          await context.sync();
        }
      }
    });
  }

  function formatDate(dateString) {
    if (!dateString) return "unknown date";
    var d = new Date(dateString);
    return d.toLocaleDateString("en-US", {
      month: "short",
      day: "numeric",
      year: "numeric",
    });
  }

  /**
   * Count how many comments are in the comment map.
   */
  function countComments(commentMap) {
    var count = 0;
    for (var word in commentMap) {
      count += commentMap[word].length;
    }
    return count;
  }

  /**
   * Quick count of comments inside the table at the cursor.
   * Used as a safety check before destructive operations.
   */
  async function countCommentsInTable() {
    if (!isSupported()) return -1; // -1 = can't check

    try {
      return await Word.run(async function (context) {
        var selection = context.document.getSelection();
        var parentTable = selection.parentTableOrNullObject;
        await context.sync();

        if (parentTable.isNullObject) return 0;

        var tableRange = parentTable.getRange("Whole");
        var tableComments = tableRange.getComments();
        tableComments.load("items");
        await context.sync();

        return tableComments.items.length;
      });
    } catch (e) {
      console.warn("countCommentsInTable failed:", e);
      return -1; // -1 = can't check
    }
  }

  return {
    isSupported: isSupported,
    captureComments: captureComments,
    deleteOriginalComments: deleteOriginalComments,
    restoreComments: restoreComments,
    countComments: countComments,
    countCommentsInTable: countCommentsInTable,
  };
})();
