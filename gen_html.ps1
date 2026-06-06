Set-ExecutionPolicy -Scope Process -ExecutionPolicy RemoteSigned -Force
Import-Module ImportExcel

$specDir = "C:\Users\MyBook PRO\Downloads\Spec"
$data = Import-Excel -Path "$specDir\SPEC_INDEX.xlsx"
$headers = $data[0].PSObject.Properties.Name | Where-Object { $_ -ne "Spec" }

$sw = New-Object System.IO.StreamWriter("$specDir\index.html", $false, [System.Text.Encoding]::UTF8)

# --- Clean name mapping ---
$fixMap = @{}
foreach ($h in $headers) {
    $c = $h
    $c = $c -replace '^MyBook L7V$', 'MyBook Pro L7V'
    $c = $c -replace '[()]', ''
    $c = $c -replace '\bPRO\b', 'Pro'
    $fixMap[$h] = $c
}
$orderedClean = $headers | ForEach-Object { $fixMap[$_] }

# --- JSON escape (String.Replace for literal replacement) ---
$esc_rdq = [char]8221  # right double quotation mark "
$esc_ldq = [char]8220  # left double quotation mark "
function esc($s) {
    $s = [string]$s
    $s = $s.Replace('\', '\\')
    $s = $s.Replace("`"", '\"')
    $s = $s.Replace("$esc_rdq", '\"')
    $s = $s.Replace("$esc_ldq", '\"')
    $s = $s.Replace("`n", '\n')
    $s = $s.Replace("`r", '')
    $s = $s.Replace("`t", '\t')
    return $s
}

# --- Build JSON manually (PS 5.1 ConvertTo-Json strips quotes) ---
$prodParts = @()
foreach ($h in $headers) {
    $clean = $fixMap[$h]
    $specParts = @()
    foreach ($row in $data) {
        $val = esc ($row.$h)
        $specParts += '"' + $row.Spec + '":"' + $val + '"'
    }
    $prodParts += '"' + $clean + '":{' + ($specParts -join ',') + '}'
}
$prodDataJson = '{' + ($prodParts -join ',') + '}'

$specList = $data.Spec
$specsJson = ($specList | ForEach-Object { '"' + $_ + '"' }) -join ','

$icons = @{}
$icons["Processor"] = "CPU"; $icons["LCD"] = "LCD"; $icons["Graphics"] = "GPU"
$icons["RAM"] = "RAM"; $icons["Storage"] = "STO"; $icons["OS"] = "OS"
$icons["Codename"] = "TAG"; $icons["Display"] = "SCR"; $icons["Camera"] = "CAM"
$icons["Weight"] = "WGT"; $icons["Connectivity"] = "NET"; $icons["I/O Ports"] = "I/O"
$icons["Power"] = "PWR"; $icons["Dimension"] = "DIM"; $icons["Speaker"] = "SPK"
$icons["Sound"] = "SND"; $icons["Security"] = "SEC"; $icons["BIOS"] = "BIOS"
$icons["NPU"] = "NPU"; $icons["Design"] = "DSG"

$cats = @{}
$cats["Processor"] = "Performance"; $cats["Graphics"] = "Performance"
$cats["RAM"] = "Performance"; $cats["Storage"] = "Performance"
$cats["NPU"] = "Performance"; $cats["LCD"] = "Display"
$cats["Display"] = "Display"; $cats["OS"] = "Software"
$cats["Codename"] = "Platform"; $cats["BIOS"] = "Platform"
$cats["Camera"] = "Features"; $cats["Connectivity"] = "Features"
$cats["I/O Ports"] = "Features"; $cats["Speaker"] = "Features"
$cats["Sound"] = "Features"; $cats["Security"] = "Features"
$cats["Weight"] = "Physical"; $cats["Power"] = "Physical"
$cats["Dimension"] = "Physical"; $cats["Design"] = "Physical"

$catColors = @{ "Performance" = "#5a67d8"; "Display" = "#38b2ac"; "Software" = "#d69e2e"; "Platform" = "#805ad5"; "Features" = "#d53f8c"; "Physical" = "#2b6cb0" }

$cleanNamesStr = '[' + (($orderedClean | ForEach-Object { '"' + $_ + '"' }) -join ',') + ']'

$iconsParts = ($icons.GetEnumerator() | ForEach-Object { '"' + $_.Key + '":"' + $_.Value + '"' }) -join ','
$iconsJson = '{' + $iconsParts + '}'
$catsParts = ($cats.GetEnumerator() | ForEach-Object { '"' + $_.Key + '":"' + $_.Value + '"' }) -join ','
$catsJson = '{' + $catsParts + '}'

$html = @"
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>MyBook Series - Product Comparison</title>
<style>
*{margin:0;padding:0;box-sizing:border-box}
body{font-family:'Inter','Segoe UI',system-ui,sans-serif;background:#f0f2f5;color:#1a202c;min-height:100vh}
.hdr{display:flex;justify-content:space-between;align-items:center;background:linear-gradient(135deg,#4f46e5,#7c3aed);color:#fff;padding:28px 32px}
.hdr h1{font-size:28px;font-weight:700;letter-spacing:-0.5px}
.hdr .sub{font-size:15px;opacity:0.85;margin-top:4px}
.hdr .cnt{font-size:13px;opacity:0.65;margin-top:8px}
.hdr .logo{height:52px;width:auto;border-radius:8px;filter:drop-shadow(0 2px 8px rgba(0,0,0,0.2))}
.main{max-width:1400px;margin:0 auto;padding:24px 32px}
.sel{display:flex;flex-wrap:wrap;gap:10px;margin-bottom:28px}
.sel .pc{border-radius:10px;padding:10px 18px;cursor:pointer;font-size:13px;font-weight:600;color:#fff;transition:all 0.2s;user-select:none;text-shadow:0 1px 2px rgba(0,0,0,0.15)}
.sel .pc:hover{transform:translateY(-2px);box-shadow:0 6px 20px rgba(0,0,0,0.2)}
.sel .pc.sel{box-shadow:0 0 0 3px #fff,0 6px 20px rgba(0,0,0,0.25);transform:translateY(-2px)}
.sel .pc.dim{opacity:0.3;pointer-events:none;filter:saturate(0.3)}
.info{text-align:center;padding:60px 20px;color:#718096;font-size:15px;line-height:1.7}
.info b{color:#4f46e5}
.tw{overflow-x:auto;border-radius:10px;box-shadow:0 4px 24px rgba(0,0,0,0.06);border:1px solid #e2e8f0}
table{width:100%;border-collapse:collapse;font-size:13px;background:#fff;min-width:600px}
thead th{position:sticky;top:0;z-index:10;background:#f8fafc;color:#64748b;padding:14px 16px;text-align:left;font-weight:600;font-size:11px;white-space:nowrap;border-bottom:2px solid #e2e8f0;text-transform:uppercase;letter-spacing:0.5px}
thead th:first-child{min-width:160px;color:#1a202c}
thead th.pc{text-align:center;color:#4f46e5;font-size:13px;text-transform:none;letter-spacing:0}
tbody tr.dr td{padding:12px 16px;border-bottom:1px solid #e2e8f0;vertical-align:top;line-height:1.5;font-size:12.5px}
tbody tr.dr:hover td{background:#fafbfc}
tbody tr.dr td:first-child{position:sticky;left:0;z-index:5;font-weight:500;color:#1a202c;background:#fafbfc;border-right:1px solid #e2e8f0;white-space:nowrap;font-size:12.5px}
tbody tr.dr td.c{text-align:center;vertical-align:middle;color:#475569}
tbody tr.dr td.m{color:#94a3b8;font-style:italic}
tbody tr.cr td{font-weight:700;font-size:11px;text-transform:uppercase;letter-spacing:0.8px;padding:10px 16px;border-bottom:1px solid #e2e8f0;background:#f8fafc;color:#4f46e5}
.sym{display:inline-block;width:24px;text-align:center;margin-right:6px;font-weight:700;font-size:11px;opacity:0.45;letter-spacing:0}
.ft{text-align:center;padding:32px 20px 24px;color:#94a3b8;font-size:13px;line-height:1.7;border-top:1px solid #e2e8f0;margin-top:32px}
.ft .fn{font-weight:600;color:#64748b}
.ta{display:none}
@media(max-width:768px){.hdr{padding:20px 16px}.hdr h1{font-size:22px}.main{padding:16px}.sel .pc{padding:8px 14px;font-size:12px}}
</style>
</head>
<body>
<div class="hdr">
  <div>
    <h1>MyBook Series Comparison</h1>
    <div class="sub">Select up to 4 products to compare side by side</div>
    <div class="cnt">$($headers.Count) models available</div>
  </div>
  <img class="logo" src="for index axioo.png" alt="Axioo">
</div>
<div class="sel" id="pclist"></div>
<div class="main">
  <div class="ta" id="tbl">
    <div class="tw">
      <table>
        <thead><tr id="thr"><th>Specification</th></tr></thead>
        <tbody id="tbd"></tbody>
      </table>
    </div>
  </div>
  <div class="info" id="info">
    <b>Click a product</b> above to start comparing specifications
  </div>
</div>
<script>
var DATA = $prodDataJson;
var SPC = [$specsJson];
var ICO = $iconsJson;
var CAT = $catsJson;
var NMS = $cleanNamesStr;

var PCOLS = ["#ef4444","#f97316","#eab308","#22c55e","#14b8a6","#6366f1","#a855f7","#ec4899","#0ea5e9","#84cc16","#f43f5e","#8b5cf6","#10b981","#d946ef","#f59e0b","#06b6d4"];

var sel = [];
function tc(n) {
  var i = sel.indexOf(n);
  i >= 0 ? sel.splice(i,1) : sel.length<4 && sel.push(n);
  r();
}
function r() {
  var L = document.getElementById("pclist");
  L.innerHTML = "";
  for (var k = 0; k < NMS.length; k++) {
    var n = NMS[k], d = document.createElement("div");
    d.className = "pc";
    var c = PCOLS[k % PCOLS.length];
    d.style.background = c;
    d.style.border = "2px solid " + c;
    if (sel.indexOf(n) >= 0) d.className += " sel";
    else if (sel.length >= 4) { d.className += " dim"; d.style.background = c; }
    d.textContent = n;
    d.onclick = function(v){ return function(){ tc(v); }; }(n);
    L.appendChild(d);
  }
  var info = document.getElementById("info");
  var tbl = document.getElementById("tbl");
  if (!sel.length) { info.style.display = ""; tbl.style.display = "none"; return; }
  info.style.display = "none"; tbl.style.display = "block";
  var thr = document.getElementById("thr");
  thr.innerHTML = "<th>Specification</th>";
  for (var j = 0; j < sel.length; j++) {
    var th = document.createElement("th");
    th.className = "pc"; th.textContent = sel[j];
    thr.appendChild(th);
  }
  var B = document.getElementById("tbd");
  B.innerHTML = "";
  var p = "";
  for (var s = 0; s < SPC.length; s++) {
    var sp = SPC[s], cat = CAT[sp] || "Other";
    if (cat !== p) {
      var cr = document.createElement("tr");
      cr.className = "cr";
      var cd = document.createElement("td");
      cd.colSpan = sel.length + 1; cd.textContent = cat;
      cr.appendChild(cd); B.appendChild(cr);
    }
    p = cat;
    var dr = document.createElement("tr");
    dr.className = "dr";
    var dd = document.createElement("td");
    dd.innerHTML = "<span class=\"sym\">" + (ICO[sp] || "") + "</span>" + sp;
    dr.appendChild(dd);
    for (var j = 0; j < sel.length; j++) {
      var v = DATA[sel[j]] ? DATA[sel[j]][sp] || "" : "";
      var td = document.createElement("td");
      td.className = v ? "c" : "c m";
      td.textContent = v || "---";
      dr.appendChild(td);
    }
    B.appendChild(dr);
  }
}
r();
</script>
<div class="ft">
  <div class="fn">&copy; Alward Jevon 2026</div>
  <div>Axioo B2G Product Command Center</div>
</div>
</body>
</html>
"@

$sw.Write($html)
$sw.Close()

Write-Output "Created: $specDir\index.html"
