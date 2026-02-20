// ===== Helpers =====
export function cleanStr(v){
  return (v ?? "").toString().trim();
}

export function yes(v){
  return cleanStr(v).toLowerCase() === "yes";
}

export function excelSerialToDate(n){
  const epoch = new Date(Date.UTC(1899, 11, 30));
  return new Date(epoch.getTime() + n * 86400000);
}

export function parsePubDate(v){
  if (!v) return null;

  if (typeof v === "number" && isFinite(v)){
    const d = excelSerialToDate(v);
    return isNaN(d.getTime()) ? null : d;
  }

  const t = Date.parse(cleanStr(v));
  if (!Number.isNaN(t)) return new Date(t);

  return null;
}

export function startOfDay(d){
  const x = new Date(d);
  x.setHours(0,0,0,0);
  return x;
}

export function endOfDay(d){
  const x = new Date(d);
  x.setHours(23,59,59,999);
  return x;
}
