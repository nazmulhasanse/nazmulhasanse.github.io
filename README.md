<!DOCTYPE html>
<html>
<head>
  <title>Job Card Report - Versioned</title>
  <style>
    body { font-family: Arial, sans-serif; font-size: 12px; }
    .report-box { border: 1px solid #000; padding: 10px; margin-bottom: 40px; }
    .header { text-align: center; font-weight: bold; }
    .middle-table { border-collapse: collapse; width: 100%; margin-top: 10px; border: 1px solid black; }
    .middle-table td, .middle-table th { padding: 5px; text-align: center; border: 1px solid black; }
    .header-table{ border-collapse:collapse;width:100%;margin-top:10px;border:none;} 
    .header-table td,.header-table th { padding:5px;text-align:left;border:none;}
    .summary { margin-top: 10px; display: grid; grid-template-columns: repeat(4, 1fr); gap: 10px; }
    .footer { margin-top: 20px; display: flex; justify-content: space-between; }
    .no-print-controls { margin-bottom: 10px; }
    @media print {
      @page { size: A4 portrait; margin: 10mm; }
      body { zoom: 90%; }
      .no-print { display: none !important; }
      .report-box { page-break-inside: avoid; }
    }
  </style>
</head>
<body>
  <div class="no-print no-print-controls">
    <label for="version">Select Report Version: </label>
    <select id="version">
      <option value="actual">Actual Data</option>
      <option value="ot2">OT capped at 2h/day</option>
      <option value="ot4">OT capped at 4h/day</option>
      <option value="noWeekend">No Working Weekends</option>
    </select>
    <br><br>
    <input id="fileInput" type="file" accept=".xlsx,.xls" onchange="loadExcel(event)">
    <button class="no-print" onclick="rebuildReport()">Rebuild Report</button>
    <button class="no-print" onclick="window.print()">Print Report</button>
  </div>

  <!-- Container for the report -->
  <div id="reports"></div>

  <script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
  <script>
    let cachedRows = null;
    let cachedEmp = null;
    let cachedDateRange = null;

    function loadExcel(event) {
      const version = document.getElementById("version").value;
      if (!version) {
        alert("Please select a version first!");
        event.target.value = "";
        return;
      }

      const file = event.target.files[0];
      const reader = new FileReader();
      reader.onload = function (e) {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });

        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        let rows = XLSX.utils.sheet_to_json(sheet, { defval: "", header: 1, range: 2 });

        const headers = XLSX.utils.sheet_to_json(sheet, { header: 1 })[0] || [];
        rows = rows.map(r => {
          const obj = {};
          headers.forEach((h, i) => obj[h] = r[i] ?? "");
          return obj;
        });

        if (!rows.length) return;

        cachedRows = rows;

        const first = rows[0] || {};
        cachedEmp = {
          id: first["Employee Id"] || first["Emp Id"] || first["EmployeeID"] || "",
          name: first["Employee Name"] || first["Name"] || "",
          designation: first["Designation"] || "",
          department: first["Department"] || first["Section"] || "",
          joiningDate: first["JoiningDate"] || first["Joining Date"] || ""
        };

        const startDate = rows[0].Date || "";
        const endDate = rows[rows.length - 1].Date || "";
        cachedDateRange = `${startDate} - ${endDate}`;

        rebuildReport(); // auto-build after upload
      };
      reader.readAsArrayBuffer(file);
    }

    function rebuildReport() {
      if (!cachedRows) {
        alert("Please upload a file first.");
        return;
      }
      const version = document.getElementById("version").value;
      if (!version) {
        alert("Please select a version first!");
        return;
      }

      const reportsDiv = document.getElementById("reports");
      reportsDiv.innerHTML = "";

      switch (version) {
        case "ot2": generateReport(cachedRows, "OT capped at 2h/day", cachedEmp, cachedDateRange, 120); break;
        case "ot4": generateReport(cachedRows, "OT capped at 4h/day", cachedEmp, cachedDateRange, 240); break;
        case "noWeekend": generateReport(cachedRows, "Exclude Weekends", cachedEmp, cachedDateRange, null, true); break;
        case "actual": generateReport(cachedRows, "Actual Data (No Manipulation)", cachedEmp, cachedDateRange); break;
      }
    }

    function generateReport(rows, title, emp, dateRange, otCap = null, excludeWeekend = false) {
      let present=0, absent=0, leave=0, weekend=0, holiday=0, totalNight=0;
      let earlyExit=0, totalTiffin=0, lateMin=0, totalOT=0;
      let eduty=0, edutyMin=0;
      let cLeave=0, sLeave=0, eLeave=0;
      let displayedRows = 0;

      let tbodyHtml = "";

      rows.forEach(r => {
        const isWeekend = (r.Status === "Weekend" || r.Status === "Weekend, Present" || r.Status === "Weekend, 0.5 day Present");

        let OTmin = toMinutes(r["Overtime"]) || 0;
        if (otCap !== null && OTmin > otCap) OTmin = otCap;

        // For excludeWeekend version, still show weekends but don't count them as working days
        if (excludeWeekend && isWeekend) {
          // weekend++;
          tbodyHtml += `
            <tr>
              <td>${r.Date || ''}</td>
              <td>${r["Shift(s)"] || ''}</td>
              <td>${time(r["Check-in"]) || ''}</td>
              <td>${time(r["Check-out"]) || ''}</td>
              <td>${r["Late Entry"] || ''}</td>
              <td>${r["Early Exit"] || ''}</td>
              <td>${r.Status || ''}</td>
              <td>--:--</td>
              <td>${r["Shift Allowance"] || ''}</td>
              <td>${r["Total Shift Allowance"] || ''}</td>
              <td>${r.Remarks || ''}</td>
            </tr>`;
          // return; // skip counting OT, present, leave, etc.
        }else{
          tbodyHtml += `
          <tr>
            <td>${r.Date || ''}</td>
            <td>${r["Shift(s)"] || ''}</td>
            <td>${time(r["Check-in"]) || ''}</td>
            <td>${time(r["Check-out"]) || ''}</td>
            <td>${r["Late Entry"] || ''}</td>
            <td>${r["Early Exit"] || ''}</td>
            <td>${r.Status || ''}</td>
            <td>${toHHMM(OTmin)}</td>
            <td>${r["Shift Allowance"] || ''}</td>
            <td>${r["Total Shift Allowance"] || ''}</td>
            <td>${r.Remarks || ''}</td>
          </tr>`;
        }

        
        displayedRows++;

        // Count summaries
        switch (r.Status) {
          case "L": leave++; break;
          case "H": holiday++; break;
          case "Absent": absent++; break;
          case "Present": present++; break;
          case "Weekend": weekend++; break;
          case "C/L": case "CL": cLeave++; leave++; break;
          case "E/L": case "EL": eLeave++; leave++; break;
          case "S/L": case "SL": sLeave++; leave++; break;
          case "0.5 day Present, 0.5 day Absent": absent++; break;
        }

        switch (r["Shift Allowance"]) {
          case "N": case "Night": totalNight++; break;
        }

        if (r["Total Shift Allowance"] && r["Total Shift Allowance"].toString().trim() !== "") {
          const val = parseFloat(r["Total Shift Allowance"]);
          if (!Number.isNaN(val)) totalTiffin += val;
        }

        lateMin   += toMinutes(r["Late Entry"])  || 0;
        totalOT   += OTmin;
        earlyExit += toMinutes(r["Early Exit"])  || 0;
      });
            // <strong>Job Card Report — ${title}</strong>

      const box = `
        <div class="report-box">
          <div class="header">
            <strong style="font-size: 2em;">Good and Fast Packaging Co. Ltd.</strong><br>
            Plot-1425, Laskarchala, Haturiachal, Kaliakair, Gazipur, Bangladesh <br>
          </div>

          <table class="header-table">
            <tr>
              <th style="text-align: left; padding: 4px;">Date Range: ${dateRange}</th>
              <th style="text-align: right; padding: 4px;">Employee Id: ${emp.id || ""}</th>
            </tr>
            <tr>
              <th style="text-align: left; padding: 4px;">Name: ${emp.name || ""}</th>
              <th style="text-align: right; padding: 4px;">Designation: ${emp.designation || ""}</th>
            </tr>
            <tr>
              <th style="text-align: left; padding: 4px;">Section/Line: ${emp.department || ""}</th>
              <th style="text-align: right; padding: 4px;">Joining Date: ${emp.joiningDate || ""}</th>
            </tr>
          </table>

          <table class="middle-table">
            <thead>
              <tr>
                <th>Date</th>
                <th>Shift</th>
                <th>In Time</th>
                <th>Out Time</th>
                <th>Late</th>
                <th>E.Exit</th>
                <th>Status</th>
                <th>OT</th>
                <th>N.Status</th>
                <th>T.Status</th>
                <th>Remarks</th>
              </tr>
            </thead>
            <tbody>${tbodyHtml}</tbody>
          </table>

          <div class="summary">
            <div>Present: ${present}</div>
            <div>Holiday: ${holiday}</div>
            <div>Early Exit: ${toHHMM(earlyExit)}</div>
            <div>Total Night: ${totalNight}</div>

            <div>Leave: ${leave}</div>
            <div>C/L: ${cLeave}</div>
            <div>Late Min: ${toHHMM(lateMin)}</div>
            <div>Night Shift: ${totalNight}</div>

            <div>Absent: ${absent}</div>
            <div>S/L: ${sLeave}</div>
            <div>Total OT: ${toHHMM(totalOT)}</div>
            <div>Total Tiffin: ${Number.isFinite(totalTiffin) ? totalTiffin : 0}</div>

            <div>Weekend: ${weekend}</div>
            <div>E/L: ${eLeave}</div>
            <div>E.Duty Min: ${edutyMin}</div>
            <div>E.Duty: ${eduty}</div>

            <div>Total Day: ${displayedRows}</div>
          </div>

          <div class="footer">
            <div>______________________<br>Signature of Employee</div>
            <div>______________________<br>Authorized Signature</div>
          </div>
          <div style="text-align: center; margin-top: 6px;">
            __________________________________________________________________________________________________<br>
            Status Legend: P-Present, A-Absent, W-Weekly Holiday, H-Holiday, CL-Casual Leave, SL-Sick Leave, EL-Earn Leave
          </div>
        </div>`;
      document.getElementById("reports").innerHTML = box;
    }

    function time(val) {
      if (!val) return "--:--";
      if (typeof val === "string" && val.includes(" ")) return val.split(" ")[1] || "--:--";
      if (typeof val === "number") {
        const d = XLSX.SSF.parse_date_code(val);
        if (d) return String(d.H).padStart(2, "0") + ":" + String(d.M).padStart(2, "0");
      }
      try {
        const d = new Date(val);
        const h = d.getHours(), m = d.getMinutes();
        if (Number.isNaN(h) || Number.isNaN(m)) return "--:--";
        return String(h).padStart(2, "0") + ":" + String(m).padStart(2, "0");
      } catch { return "--:--"; }
    }

    function toMinutes(str) {
      if (!str) return 0;
      const parts = String(str).split(":");
      if (parts.length !== 2) return 0;
      const h = parseInt(parts[0], 10) || 0;
      const m = parseInt(parts[1], 10) || 0;
      return h * 60 + m;
    }

    function toHHMM(mins) {
      const total = Math.max(0, mins|0);
      const h = Math.floor(total / 60);
      const m = total % 60;
      return String(h).padStart(2,"0") + ":" + String(m).padStart(2,"0");
    }
  </script>
</body>
</html>
