import { cleanStr, yes, parsePubDate } from "./utils.js";

export async function loadData(EXCEL_FILE, SHEET_NAME, from, to){

  const res = await fetch(`${EXCEL_FILE}?v=${Date.now()}`);
  if(!res.ok) throw new Error("Excel missing");

  const buf = await res.arrayBuffer();
  const wb = XLSX.read(buf,{type:"array"});
  const sheet = wb.Sheets[SHEET_NAME];
  if(!sheet) throw new Error("Sheet missing");

  const rows = XLSX.utils.sheet_to_json(sheet,{defval:""});

  const filtered = rows.filter(r=>{
    if(!from && !to) return true;
    const d = parsePubDate(r["pubdate"]);
    if(!d) return false;
    if(from && d < from) return false;
    if(to && d > to) return false;
    return true;
  });

  const byArea = {};

  for(const r of filtered){
    const area = cleanStr(r["Area"]);
    if(!area) continue;

    if(!byArea[area])
      byArea[area]={fiveStar:0,five:0,utd:0,ft:0};

    const ranking = cleanStr(r["Ranking - Updated"]);

    if(ranking==="5*") byArea[area].fiveStar++;
    if(ranking==="5")  byArea[area].five++;

    if(yes(r["UTD Journal"])) byArea[area].utd++;
    if(yes(r["Financial Times Journal"])) byArea[area].ft++;
  }

  return { rows, filtered, byArea };
}
