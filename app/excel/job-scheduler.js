// app/excel/job-scheduler.js
(function () {
  
  // Parses localized text values cleanly into proper ISO dates for BTP Job Scheduler
  function parseUtcDateTime(dateValue, timeValue) {
    if (!dateValue || !timeValue) return null;

    var cleanedDate = String(dateValue).trim().replace(/\./g, '-').replace(/\//g, '-');
    var dateParts = cleanedDate.split('-').map(function (part) { return parseInt(part, 10); });
    var year, month, day;

    if (dateParts.length === 3) {
      if (dateParts[0] > 31) {
        year = dateParts[0]; month = dateParts[1]; day = dateParts[2];
      } else {
        day = dateParts[0]; month = dateParts[1]; year = dateParts[2];
      }
    }

    var timeParts = String(timeValue).trim().split(':').map(function (part) { return parseInt(part, 10); });
    var hour = timeParts[0] || 0;
    var minute = timeParts[1] || 0;

    if (!year || !month || !day || isNaN(hour) || isNaN(minute)) return null;

    return new Date(Date.UTC(year, month - 1, day, hour, minute, 0));
  }

  // Wraps your row array entry into the specific format expected by BTP Job Scheduler APIs
  function buildSchedulePayload(payload) {
    var startDate;
    if (payload.Immediate && String(payload.Immediate).toLowerCase() === 'yes') {
      startDate = new Date().toISOString();
    } else {
      var parsed = parseUtcDateTime(payload.Date, payload.Time);
      if (!parsed) throw new Error('Invalid date/time for job payload: ' + payload.JobName);
      startDate = parsed.toISOString();
    }

    return {
      jobName: 'OnPremJobs',
      description: payload.JobName,
      payload: payload,
      startDate: startDate,
      timezone: 'UTC',
      enabled: true
    };
  }

  // Exposed globally to bind seamlessly with your UI5 Execute Button click handler
  window.scheduleOnPremJobs = function () {
    var resultEl = document.getElementById('resultText');
    var payloads = window.latestSchedulePayloads || [];

    if (!payloads || payloads.length === 0) {
      var msg = 'Please upload and validate an Excel file before scheduling jobs.';
      if (resultEl) resultEl.innerText = msg;
      if (window.sap && sap.m && sap.m.MessageToast) sap.m.MessageToast.show(msg);
      return;
    }

    if (resultEl) resultEl.innerText = 'Submitting requests to CAP Backend...';

    var requests = payloads.map(function (payload) {
      try {
        var schedulePayload = buildSchedulePayload(payload);
        
        // CAP action targets the unbound service path /api/scheduleJob
        return fetch('/api/scheduleJob', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ payload: JSON.stringify(schedulePayload) })
        })
        .then(function (res) {
          if (!res.ok) {
            return res.text().then(function (t) { throw new Error(t); });
          }
          return res.json();
        })
        .then(function (json) {
          return { success: true, response: json };
        })
        .catch(function (err) {
          return { success: false, error: err.message || err.toString() };
        });
      } catch (ex) {
        return Promise.resolve({ success: false, error: ex.message });
      }
    });

    // Resolve all background jobs collectively and output logs to screen
    Promise.all(requests).then(function (results) {
      var messages = results.map(function (result, index) {
        if (result.success) return 'Row ' + (index + 1) + ': Scheduled successfully.';
        return 'Row ' + (index + 1) + ': Failed - ' + result.error;
      });
      
      var finalOutput = 'Processing Summary:\n' + messages.join('\n');
      if (resultEl) resultEl.innerText = finalOutput;
      if (window.sap && sap.m && sap.m.MessageToast) sap.m.MessageToast.show('Processing complete');
    });
  };
})();