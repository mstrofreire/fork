export function toCSV(rows: string[][]): string {
  const esc = (v: string) => {
    if (v == null) return "";
    if (/[",\n]/.test(v)) return '"' + v.replace(/"/g, '""') + '"';
    return v;
  };
  return rows.map(r => r.map(esc).join(",")).join("\n");
}

export function fromCSV(csv: string): string[][] {
  const result: string[][] = [];
  let row: string[] = [];
  let cur = "";
  let inQuotes = false;
  for (let i = 0; i < csv.length; i++) {
    const ch = csv[i];
    if (inQuotes) {
      if (ch === '"') {
        if (csv[i + 1] === '"') { cur += '"'; i++; }
        else { inQuotes = false; }
      } else {
        cur += ch;
      }
    } else {
      if (ch === '"') inQuotes = true;
      else if (ch === ',') { row.push(cur); cur = ""; }
      else if (ch === '\n' || ch === '\r') {
        if (ch === '\r' && csv[i + 1] === '\n') i++; // handle CRLF
        row.push(cur); cur = ""; result.push(row); row = [];
      } else { cur += ch; }
    }
  }
  row.push(cur);
  result.push(row);
  return result;
}