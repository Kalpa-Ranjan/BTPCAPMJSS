// excel/excel-handler.js
// Provides `downloadTemplate()` and `handleFileInput(ev)` globally
(function () {
  // Load SheetJS (xlsx) library dynamically for upload parsing
  function loadXLSX(onload) {
    if (window.XLSX) return onload();
    var s = document.createElement('script');
    s.src = 'https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js';
    s.onload = onload;
    s.onerror = function () { console.error('Failed to load XLSX library'); };
    document.head.appendChild(s);
  }

  // Load ExcelJS library dynamically for styled XLSX template generation
  function loadExcelJS(onload) {
    if (window.ExcelJS) return onload();
    var s = document.createElement('script');
    s.src = 'https://cdn.jsdelivr.net/npm/exceljs/dist/exceljs.min.js';
    s.onload = onload;
    s.onerror = function () { console.error('Failed to load ExcelJS library'); };
    document.head.appendChild(s);
  }

  // Create hidden file input and attach handler
  function ensureFileInput() {
    if (document.getElementById('fileInput')) return;
    var input = document.createElement('input');
    input.type = 'file';
    input.id = 'fileInput';
    input.accept = '.xlsx,.xls';
    input.style.display = 'none';
    input.onchange = function (ev) { window.handleFileInput && window.handleFileInput(ev); };
    document.body.appendChild(input);
  }

  // Expose downloadTemplate
  window.downloadTemplate = function () {
    loadExcelJS(function () {
      var workbook = new ExcelJS.Workbook();
      var worksheet = workbook.addWorksheet('Template');
      worksheet.columns = [
        { header: 'Job Name', key: 'jobName', width: 30 },
        { header: 'Program', key: 'program', width: 20 },
        { header: 'Variant', key: 'variant', width: 15 },
        { header: 'Date', key: 'date', width: 12 },
        { header: 'Time', key: 'time', width: 10 },
        { header: 'Immediate', key: 'immediate', width: 12 }
      ];

      var headerRow = worksheet.getRow(1);
      headerRow.eachCell(function (cell) {
        cell.font = { name: 'Calibri', size: 11, bold: true, color: { argb: 'FF000000' } };
        cell.alignment = { horizontal: 'center', vertical: 'middle' };
        cell.border = {
          top: { style: 'thin', color: { argb: 'FF000000' } },
          bottom: { style: 'thin', color: { argb: 'FF000000' } },
          left: { style: 'thin', color: { argb: 'FF000000' } },
          right: { style: 'thin', color: { argb: 'FF000000' } }
        };
      });

      var dataRow = worksheet.addRow({ jobName: 'Example Job', program: 'ZPROG', variant: 'VAR1', date: '2026-05-28', time: '12:00', immediate: 'No' });
      dataRow.eachCell(function (cell) {
        cell.border = {
          top: { style: 'thin', color: { argb: 'FF000000' } },
          bottom: { style: 'thin', color: { argb: 'FF000000' } },
          left: { style: 'thin', color: { argb: 'FF000000' } },
          right: { style: 'thin', color: { argb: 'FF000000' } }
        };
      });

      workbook.xlsx.writeBuffer().then(function (buffer) {
        var blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
        var url = URL.createObjectURL(blob);
        var a = document.createElement('a');
        a.href = url;
        a.download = 'template.xlsx';
        document.body.appendChild(a);
        a.click();
        a.remove();
        URL.revokeObjectURL(url);
      }).catch(function (err) {
        console.error('ExcelJS write error:', err);
      });
    });
  };

  // Expose handleFileInput
  window.handleFileInput = function (ev) {
    loadXLSX(function () {
      var file = ev.target.files && ev.target.files[0];
      if (!file) return;

      var reader = new FileReader();
      reader.onload = function (e) {
        try {
          var data = e.target.result;
          var workbook = XLSX.read(data, { type: 'array' });
          var firstSheetName = workbook.SheetNames[0];
          var worksheet = workbook.Sheets[firstSheetName];
          var rows = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

          var el = document.getElementById('resultText');
          if (!rows || rows.length === 0) {
            var msg = 'Uploaded file is empty.';
            if (el) el.innerText = msg;
            if (window.sap && sap.m && sap.m.MessageToast) {
              sap.m.MessageToast.show(msg);
            }
            return;
          }

          var headerRow = rows[0].map(function (h) { return String(h || '').trim().toLowerCase(); });
          var findIndex = function (name) {
            return headerRow.findIndex(function (h) { return h === name.toLowerCase(); });
          };

          var dateIdx = findIndex('date');
          var timeIdx = findIndex('time');
          var immediateIdx = findIndex('immediate');

          if (dateIdx < 0 || timeIdx < 0 || immediateIdx < 0) {
            var missing = [];
            if (dateIdx < 0) missing.push('Date');
            if (timeIdx < 0) missing.push('Time');
            if (immediateIdx < 0) missing.push('Immediate');
            var msg = 'Missing required columns: ' + missing.join(', ');
            if (el) el.innerText = msg;
            if (window.sap && sap.m && sap.m.MessageToast) {
              sap.m.MessageToast.show(msg);
            }
            return;
          }

          var errors = [];
          var validatedRows = [];

          for (var r = 1; r < rows.length; r++) {
            var row = rows[r] || [];
            var immediateValue = String(row[immediateIdx] || '').trim();
            var dateValue = String(row[dateIdx] || '').trim();
            var timeValue = String(row[timeIdx] || '').trim();
            var immediateNorm = immediateValue.toLowerCase();

            if (!immediateValue) {
              errors.push('Row ' + (r + 1) + ': Immediate must be Yes or No.');
              continue;
            }

            if (immediateNorm === 'yes' || immediateNorm === 'y') {
              if (dateValue !== '' || timeValue !== '') {
                errors.push('Row ' + (r + 1) + ': Immediate=Yes requires Date and Time to be blank.');
              }
            } else if (immediateNorm === 'no' || immediateNorm === 'n') {
              if (dateValue === '' || timeValue === '') {
                errors.push('Row ' + (r + 1) + ': Immediate=No requires both Date and Time to be filled.');
              }
            } else {
              errors.push('Row ' + (r + 1) + ': Immediate must be Yes or No.');
            }

            validatedRows.push(row);
          }

          if (errors.length > 0) {
            if (el) el.innerText = 'Validation errors:\n' + errors.join('\n');
            if (window.sap && sap.m && sap.m.MessageToast) {
              sap.m.MessageToast.show('Excel validation failed.');
            }
          } else {
            var payloads = [];
            if (typeof window.convertRowsToPayloads === 'function') {
              try {
                payloads = window.convertRowsToPayloads([rows[0]].concat(validatedRows), 'TESTBEH');
              } catch (convertErr) {
                console.error(convertErr);
              }
            }

            window.latestSchedulePayloads = payloads;

            if (el) {
              if (payloads.length > 0 && typeof window.formatPayloadsAsSapList === 'function') {
                el.innerText = 'Validation passed. JSON payloads:\n\n' + window.formatPayloadsAsSapList(payloads);
              } else {
                var headerText = rows[0].map(function (h) { return String(h || '').trim(); }).join(' | ');
                var separator = headerText.split('').map(function () { return '-'; }).join('');
                var rowText = rows.slice(1).map(function (row) {
                  return row.map(function (cell) { return String(cell || ''); }).join(' | ');
                }).join('\n');
                el.innerText = 'Validation passed.\n' + headerText + '\n' + separator + '\n' + rowText;
              }
            }
            if (window.sap && sap.m && sap.m.MessageToast) {
              sap.m.MessageToast.show('Excel uploaded and validated successfully.');
            }
          }
        } catch (err) {
          console.error(err);
        }
      };
      reader.readAsArrayBuffer(file);
      // reset input so same file can be uploaded again
      ev.target.value = '';
    });
  };

  // Initialize
  ensureFileInput();
})();
