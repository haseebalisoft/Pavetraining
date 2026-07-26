/**
 * Client-side CSV / Excel downloads (no extra packages).
 */

function escapeCsvCell(value: string): string {
  const needsQuotes =
    value.indexOf(",") >= 0 ||
    value.indexOf('"') >= 0 ||
    value.indexOf("\n") >= 0 ||
    value.indexOf("\r") >= 0;
  const escaped = value.replace(/"/g, '""');
  return needsQuotes ? '"' + escaped + '"' : escaped;
}

function triggerDownload(blob: Blob, filename: string): void {
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  a.style.display = "none";
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}

function stamp(): string {
  const d = new Date();
  const p = (n: number): string => (n < 10 ? "0" + n : String(n));
  return (
    d.getFullYear() +
    p(d.getMonth() + 1) +
    p(d.getDate()) +
    "-" +
    p(d.getHours()) +
    p(d.getMinutes())
  );
}

export function exportTableAsCsv(
  title: string,
  headers: string[],
  rows: string[][]
): void {
  const lines: string[] = [];
  lines.push(headers.map(escapeCsvCell).join(","));
  for (let i = 0; i < rows.length; i++) {
    lines.push(rows[i].map((c) => escapeCsvCell(c == null ? "" : String(c))).join(","));
  }
  // UTF-8 BOM so Excel opens accents correctly
  const csv = "\uFEFF" + lines.join("\r\n");
  const blob = new Blob([csv], { type: "text/csv;charset=utf-8;" });
  const safe = title.replace(/[^\w\-]+/g, "_").replace(/^_+|_+$/g, "") || "export";
  triggerDownload(blob, safe + "-" + stamp() + ".csv");
}

/**
 * Excel-compatible .xls via HTML table (opens directly in Microsoft Excel).
 */
export function exportTableAsExcel(
  title: string,
  headers: string[],
  rows: string[][]
): void {
  const escapeHtml = (s: string): string =>
    String(s)
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;");

  let html =
    '<html xmlns:o="urn:schemas-microsoft-com:office:office" ' +
    'xmlns:x="urn:schemas-microsoft-com:office:excel" ' +
    'xmlns="http://www.w3.org/TR/REC-html40">' +
    "<head><meta charset=\"UTF-8\"><!--[if gte mso 9]><xml>" +
    "<x:ExcelWorkbook><x:ExcelWorksheets><x:ExcelWorksheet>" +
    "<x:Name>Sheet1</x:Name>" +
    "<x:WorksheetOptions><x:DisplayGridlines/></x:WorksheetOptions>" +
    "</x:ExcelWorksheet></x:ExcelWorksheets></x:ExcelWorkbook>" +
    "</xml><![endif]--></head><body><table border=\"1\"><thead><tr>";

  for (let h = 0; h < headers.length; h++) {
    html += "<th>" + escapeHtml(headers[h]) + "</th>";
  }
  html += "</tr></thead><tbody>";

  for (let r = 0; r < rows.length; r++) {
    html += "<tr>";
    const row = rows[r];
    for (let c = 0; c < headers.length; c++) {
      html += "<td>" + escapeHtml(row[c] == null ? "" : String(row[c])) + "</td>";
    }
    html += "</tr>";
  }
  html += "</tbody></table></body></html>";

  const blob = new Blob([html], {
    type: "application/vnd.ms-excel;charset=utf-8;",
  });
  const safe = title.replace(/[^\w\-]+/g, "_").replace(/^_+|_+$/g, "") || "export";
  triggerDownload(blob, safe + "-" + stamp() + ".xls");
}
