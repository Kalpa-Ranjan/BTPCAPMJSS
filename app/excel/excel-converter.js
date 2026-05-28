// excel/excel-converter.js
// Converts Excel rows into JSON payloads for SAP-style processing.
(function () {
  function normalizeHeader(name) {
    return String(name || '').trim().toLowerCase();
  }

  function findColumnIndex(headers, keys) {
    keys = Array.isArray(keys) ? keys : [keys];
    var lowerKeys = keys.map(function (k) { return k.toLowerCase(); });
    for (var i = 0; i < headers.length; i++) {
      if (lowerKeys.indexOf(normalizeHeader(headers[i])) >= 0) {
        return i;
      }
    }
    return -1;
  }

  function toPayloadValue(value) {
    return String(value || '').trim();
  }

  function normalizeImmediate(value) {
    var normalized = String(value || '').trim().toLowerCase();
    if (normalized === 'yes' || normalized === 'y') {
      return 'Yes';
    }
    if (normalized === 'no' || normalized === 'n') {
      return 'No';
    }
    return String(value || '').trim();
  }

  window.convertRowsToPayloads = function (rows, username) {
    if (!Array.isArray(rows) || rows.length < 2) {
      return [];
    }

    if (typeof username === 'undefined' || username === null) {
      username = '';
    }

    var headerRow = rows[0].map(function (cell) { return String(cell || '').trim(); });
    var jobNameIdx = findColumnIndex(headerRow, ['JobName', 'Job Name']);
    var programIdx = findColumnIndex(headerRow, ['Program']);
    var variantIdx = findColumnIndex(headerRow, ['VARIANT', 'Variant']);
    var dateIdx = findColumnIndex(headerRow, ['Date']);
    var timeIdx = findColumnIndex(headerRow, ['Time']);
    var immediateIdx = findColumnIndex(headerRow, ['Immediate']);

    if (jobNameIdx < 0 || programIdx < 0 || variantIdx < 0 || immediateIdx < 0) {
      throw new Error('Missing required Excel columns: JobName, Program, VARIANT, or Immediate.');
    }

    var payloads = [];
    for (var r = 1; r < rows.length; r++) {
      var row = rows[r] || [];
      var jobName = toPayloadValue(row[jobNameIdx]);
      var program = toPayloadValue(row[programIdx]);
      var variant = toPayloadValue(row[variantIdx]);
      var date = toPayloadValue(row[dateIdx]);
      var time = toPayloadValue(row[timeIdx]);
      var immediate = normalizeImmediate(row[immediateIdx]);

      if (!jobName && !program && !variant && !immediate) {
        continue; // skip empty rows
      }

      payloads.push({
        JobName: jobName,
        Program: program,
        VARIANT: variant,
        Date: date,
        Time: time,
        Immediate: immediate,
        Username: String(username),
        IMMEDIATE: immediate.toLowerCase() === 'yes' ? 'true' : 'false'
      });
    }

    return payloads;
  };

  window.formatPayloadsAsSapList = function (payloads) {
    if (!Array.isArray(payloads) || payloads.length === 0) {
      return 'No payloads generated.';
    }

    return payloads.map(function (payload, index) {
      var lines = ['Payload ' + (index + 1) + ':'];
      Object.keys(payload).forEach(function (key) {
        lines.push('  ' + key + ' = ' + payload[key]);
      });
      return lines.join('\n');
    }).join('\n\n');
  };
})();
