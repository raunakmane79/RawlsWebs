const charts = {};

const FALLBACK_COLOR = "#475467";

export const AREA_SOLID_COLORS = {
  MANAGEMENT:"#4169E1",
  MARKETING:"#CC0000",
  SUPPLYCHAIN:"#228B22",
  ISQS:"#00A3E0",
  FINANCE:"#003366",
  ECBE:"#FFBF00",
  ACCOUNTING:"#800020"
};

export function areaCategoryKey(area){
  const a = area.toLowerCase();

  if (a.includes("marketing") && a.includes("supply chain"))
    return "SUPPLYCHAIN";

  if (a.includes("management")) return "MANAGEMENT";
  if (a.includes("finance")) return "FINANCE";
  if (a.includes("account")) return "ACCOUNTING";
  if (a.includes("isqs")) return "ISQS";
  if (a.includes("supply")) return "SUPPLYCHAIN";
  if (a.includes("marketing")) return "MARKETING";

  return "";
}

function destroy(key){
  if(charts[key]){
    charts[key].destroy();
    delete charts[key];
  }
}

export function buildBar(canvasId, labels, values, key, labelToColor){
  const el = document.getElementById(canvasId);
  if(!el) return;

  destroy(key);

  charts[key] = new Chart(el,{
    type:"bar",
    data:{
      labels,
      datasets:[{
        data:values,
        backgroundColor: labels.map(l=>labelToColor[l]||FALLBACK_COLOR),
        borderRadius:10,
        borderWidth:0
      }]
    },
    options:{
      responsive:true,
      maintainAspectRatio:false,
      plugins:{ legend:{display:false}}
    }
  });
}
