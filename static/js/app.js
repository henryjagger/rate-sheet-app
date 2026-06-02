// Select / deselect all term checkboxes
function selectAllTerms(checked) {
  document.querySelectorAll('.term-check').forEach(cb => { cb.checked = checked; });
}

// Tab switching
function showTab(name) {
  document.querySelectorAll('.tab-panel').forEach(p => p.classList.remove('active'));
  document.querySelectorAll('.nav-tab').forEach(t => t.classList.remove('active'));
  const panel = document.getElementById('tab-' + name);
  if (panel) panel.classList.add('active');
  document.querySelectorAll('.nav-tab').forEach(t => {
    if (t.getAttribute('onclick') === "showTab('" + name + "')") {
      t.classList.add('active');
    }
  });
  history.replaceState(null, '', '?tab=' + name);
}

// Source toggle (master / formatted)
function toggleSource(val) {
  const master    = document.getElementById('upload-master');
  const formatted = document.getElementById('upload-formatted');
  if (!master || !formatted) return;
  if (val === 'master') {
    master.classList.remove('hidden');
    formatted.classList.add('hidden');
  } else {
    master.classList.add('hidden');
    formatted.classList.remove('hidden');
  }
}

// Copy results table to clipboard as rich HTML + TSV fallback
async function copyTable() {
  const table = document.getElementById('results-table');
  if (!table) return;

  const btn = document.getElementById('copy-btn');

  // Build Calibri 11pt HTML for Outlook/Excel paste
  const thStyle  = "border:1px solid #ccc;padding:6px 12px;background-color:#000000;font-family:Calibri,sans-serif;font-size:11pt;";
  const tdStyle  = "border:1px solid #ccc;padding:6px 12px;font-family:Calibri,sans-serif;font-size:11pt;vertical-align:middle;";
  const rateStyle = tdStyle + "color:#C00000;";

  let htmlOut = "<table style='border-collapse:collapse;'><thead><tr>";
  table.querySelectorAll('thead th').forEach(th => {
    htmlOut += `<th style='${thStyle}'><font color='#ffffff' face='Calibri'><b>${th.innerText}</b></font></th>`;
  });
  htmlOut += "</tr></thead><tbody>";

  table.querySelectorAll('tbody tr').forEach(tr => {
    htmlOut += "<tr>";
    tr.querySelectorAll('td').forEach(td => {
      const span  = td.getAttribute('rowspan') ? ` rowspan='${td.getAttribute('rowspan')}'` : '';
      const isRate   = td.classList.contains('cell-rate');
      const style = isRate ? rateStyle : tdStyle;
      const color = isRate ? " color='#C00000'" : "";
      htmlOut += `<td${span} style='${style}'><font face='Calibri'${color}>${td.innerHTML}</font></td>`;
    });
    htmlOut += "</tr>";
  });
  htmlOut += "</tbody></table>";

  // TSV fallback
  let tsv = "";
  table.querySelectorAll('thead th').forEach((th, i, arr) => {
    tsv += th.innerText + (i < arr.length - 1 ? "\t" : "\n");
  });
  table.querySelectorAll('tbody tr').forEach(tr => {
    const cells = tr.querySelectorAll('td');
    cells.forEach((td, i) => {
      tsv += td.innerText.trim() + (i < cells.length - 1 ? "\t" : "\n");
    });
  });

  try {
    await navigator.clipboard.write([
      new ClipboardItem({
        'text/html':  new Blob([htmlOut], { type: 'text/html' }),
        'text/plain': new Blob([tsv],     { type: 'text/plain' }),
      })
    ]);
  } catch (e) {
    await navigator.clipboard.writeText(tsv);
  }

  if (btn) {
    const orig = btn.textContent;
    btn.textContent = '✓ Copied!';
    setTimeout(() => { btn.textContent = orig; }, 2000);
  }
}
