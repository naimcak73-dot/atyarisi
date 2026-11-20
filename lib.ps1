# Ortak analiz + HTTP isleyici kutuphanesi (hem ana surec hem runspace'ler kullanir)
. (Join-Path $PSScriptRoot 'analiz.ps1')

function Resolve-Tarih([string]$s) {
    if ([string]::IsNullOrWhiteSpace($s)) { return (Get-Date).ToString('dd/MM/yyyy', $script:CIV) }
    $t = $s.Trim()
    $low = $t.ToLowerInvariant()
    if ($low -eq 'bugun' -or $low -eq 'bugün') { return (Get-Date).ToString('dd/MM/yyyy', $script:CIV) }
    if ($low -eq 'dun' -or $low -eq 'dün') { return (Get-Date).AddDays(-1).ToString('dd/MM/yyyy', $script:CIV) }
    if ($low -eq 'yarin' -or $low -eq 'yarın') { return (Get-Date).AddDays(1).ToString('dd/MM/yyyy', $script:CIV) }
    $d = [datetime]::MinValue
    if ([datetime]::TryParseExact($t, 'yyyy-MM-dd', $script:CIV, 'None', [ref]$d)) { return $d.ToString('dd/MM/yyyy', $script:CIV) }
    if ([datetime]::TryParseExact($t, 'dd/MM/yyyy', $script:CIV, 'None', [ref]$d)) { return $d.ToString('dd/MM/yyyy', $script:CIV) }
    if ([datetime]::TryParse($t, $script:CIV, 'None', [ref]$d)) { return $d.ToString('dd/MM/yyyy', $script:CIV) }
    return (Get-Date).ToString('dd/MM/yyyy', $script:CIV)
}

function To-JsonSafe($obj) { return (($obj | ConvertTo-Json -Depth 16 -Compress) -replace '</','<\/') }

# --- gun bazli payload onbellegi + TJK kapali isareti ---
function Get-PayloadCacheFile([string]$tarih, [string]$ek) {
    $d = [datetime]::MinValue
    if (-not [datetime]::TryParseExact($tarih, 'dd/MM/yyyy', $script:CIV, 'None', [ref]$d)) { return $null }
    $dir = Join-Path $script:Root 'veri'
    if (-not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
    if (-not $ek) { $ek = '' }
    return (Join-Path $dir ('gun-' + $d.ToString('yyyyMMdd', $script:CIV) + $ek + '.json'))
}
function Test-TjkDownMark {
    try {
        $f = Join-Path (Join-Path $script:Root 'veri') 'tjk-kapali.txt'
        if (Test-Path $f) { return (((Get-Date) - (Get-Item $f).LastWriteTime).TotalMinutes -lt 8) }
    } catch {}
    return $false
}
function Set-TjkDownMark([bool]$down) {
    try {
        $dir = Join-Path $script:Root 'veri'
        if (-not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
        $f = Join-Path $dir 'tjk-kapali.txt'
        if ($down) { [System.IO.File]::WriteAllText($f, (Get-Date).ToString('s')) }
        elseif (Test-Path $f) { Remove-Item $f -Force -ErrorAction SilentlyContinue }
    } catch {}
}
function Get-LatestPayloadCache {
    try {
        $dir = Join-Path $script:Root 'veri'
        $all = @(Get-ChildItem $dir -Filter 'gun-*.json' -ErrorAction SilentlyContinue)
        $f = @($all | Where-Object { $_.Name -notmatch '-yd\.json$' }) | Sort-Object Name -Descending | Select-Object -First 1
        if (-not $f) { $f = $all | Sort-Object Name -Descending | Select-Object -First 1 }
        if ($f) {
            $c = [System.IO.File]::ReadAllText($f.FullName, [System.Text.Encoding]::UTF8) | ConvertFrom-Json
            if ($c) {
                $c | Add-Member -NotePropertyName Uyari -NotePropertyValue ("İstenen tarih için veri yok (TJK). Son önbellek gösteriliyor: {0}." -f $c.Tarih) -Force
                return $c
            }
        }
    } catch {}
    return $null
}

function Read-PayloadCache([string]$cf, [string]$msg) {
    if (-not $cf) { return $null }
    try {
        if (-not (Test-Path $cf)) { return $null }
        $c = [System.IO.File]::ReadAllText($cf, [System.Text.Encoding]::UTF8) | ConvertFrom-Json
        if ($c) {
            if (-not $msg) { $msg = 'TJK şu an yanıt vermiyor; önbellekteki veri gösteriliyor.' }
            $c | Add-Member -NotePropertyName Bilgi -NotePropertyValue $msg -Force
        }
        return $c
    } catch { return $null }
}

function Build-Payload([string]$tarih, [bool]$hizli) {
    $script:Hizli = [bool]$hizli
    $cf = Get-PayloadCacheFile $tarih $(if ($script:Yabanci) { '-yd' } else { '' })
    # hizli yol + 5 dk'dan taze onbellek: aga hic cikmadan aninda ver
    if ($hizli -and $cf -and (Test-Path $cf)) {
        try {
            if (((Get-Date) - (Get-Item $cf).LastWriteTime).TotalMinutes -lt 5) {
                $c = Read-PayloadCache $cf 'Önbellekten gösteriliyor (son 5 dk).'
                if ($c) { return $c }
            }
        } catch {}
    }
    # TJK kapali isaretli: agi hic denemeden onbellegi ver (aninda)
    if (Test-TjkDownMark) {
        $c = Read-PayloadCache $cf
        if ($c) { return $c }
    }
    $g = Build-DayData $tarih $queryOptions $winsCache $histCache $statsCache
    if (-not $g) {
        Set-TjkDownMark $true
        $c = Read-PayloadCache $cf
        if ($c) { return $c }
        return $null
    }
    Set-TjkDownMark $false
    $d = [datetime]::MinValue
    try { [void][datetime]::TryParseExact($tarih, 'dd/MM/yyyy', $script:CIV, 'None', [ref]$d) } catch {}
    $iso = if ($d -ne [datetime]::MinValue) { $d.ToString('yyyy-MM-dd', $script:CIV) } else { '' }
    $out = [pscustomobject]@{
        __api__  = $false
        Uretim   = [DateTime]::UtcNow.AddHours(3).ToString('dd.MM.yyyy HH:mm')
        Donem    = $Donem
        Tarihler = @($g)
        Tarih    = $g.Tarih
        TarihISO = $iso
        Hava     = $g.Hava
        Sehirler = $g.Sehirler
    }
    # gercek veri geldiyse gun payload'ini onbellege yaz (TJK kapaliyken aninda acilsin)
    if ($cf -and @($g.Sehirler).Count -gt 0) {
        try { [System.IO.File]::WriteAllText($cf, (($out | ConvertTo-Json -Depth 16 -Compress) -replace '</','<\/'), (New-Object System.Text.UTF8Encoding($false))) } catch {}
    }
    return $out
}

function isoFromTRlieu([string]$s) {
    $d = [datetime]::MinValue
    if ([datetime]::TryParseExact($s, 'dd/MM/yyyy', $script:CIV, 'None', [ref]$d)) { return $d.ToString('yyyy-MM-dd', $script:CIV) }
    return $s
}

function Build-Shell([string]$tarih) {
    $tpl = [System.IO.File]::ReadAllText($tplPath, [System.Text.Encoding]::UTF8)
    $shell = [pscustomobject]@{
        __api__  = $true
        Uretim   = [DateTime]::UtcNow.AddHours(3).ToString('dd.MM.yyyy HH:mm')
        Donem    = $Donem
        Tarih    = $tarih
        TarihISO = (isoFromTRlieu $tarih)
        Tarihler = @()
        Sehirler = @()
    }
    return $tpl.Replace('__DATA__', (To-JsonSafe $shell))
}

function Build-Html([string]$tarih, [bool]$hizli) {
    $payload = Build-Payload $tarih $hizli
    if (-not $payload) {
        return "<!doctype html><meta charset='utf-8'><body style='font-family:system-ui;background:#0f1115;color:#e8ecf1;padding:24px'><h2>$tarih — veri alınamadı</h2><p>TJK şu an yanıt vermiyor olabilir. Birkaç dakika sonra tekrar deneyin.</p><a style='color:#4f8cff' href='/'>&larr; Geri</a></body>"
    }
    $tpl = [System.IO.File]::ReadAllText($tplPath, [System.Text.Encoding]::UTF8)
    return $tpl.Replace('__DATA__', (To-JsonSafe $payload))
}

function Get-StatsFile([string]$tip, [string]$grup, [int]$n, [string]$sehir) {
    $tip = (([string]$tip) -split '/')[0].Trim()
    $grup = ''
    if (-not $sehir) { $sehir = '-1' }
    if ($sehir -notmatch '^-?\d+$') { $sehir = '-1' }
    $key = (($tip + '|' + $grup + '|' + $n + '|' + $sehir) -replace '[^A-Za-z0-9]+', '_')
    if ($key.Length -gt 120) { $key = $key.Substring(0, 120) }
    return (Join-Path (Join-Path $script:Root 'veri') ('istatistik-' + $key + '.json'))
}

function Get-StatsCache([string]$file) {
    if (-not (Test-Path $file)) { return $null }
    try {
        $c = ([System.IO.File]::ReadAllText($file, [System.Text.Encoding]::UTF8) | ConvertFrom-Json)
        if ($c -and -not $c.Hata -and ($c.PSObject.Properties.Name -contains 'KadroSayisi') -and [int]$c.KosuSayisi -gt 0) { return $c }
    } catch {}
    return $null
}

function Test-StatsFresh([string]$file) {
    if (-not (Test-Path $file)) { return $false }
    if (-not (Get-StatsCache $file)) { return $false }
    $age = ((Get-Date) - (Get-Item $file).LastWriteTime).TotalHours
    return ($age -lt 24)
}

# Istek aninda uzun TJK taramasini engellemek icin istenen istatistigi kuyruga yazar.
function Add-StatsQueue([string]$tip, [int]$n, [string]$sehir) {
    try {
        $dir = Join-Path $script:Root 'veri'
        if (-not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
        if (-not $sehir) { $sehir = '-1' }
        if ($sehir -notmatch '^-?\d+$') { $sehir = '-1' }
        $line = ('{0}|{1}|{2}' -f (([string]$tip) -split '/')[0].Trim(), $n, $sehir)
        $q = Join-Path $dir 'kuyruk.txt'
        if (Test-Path $q) { try { if (Select-String -Path $q -SimpleMatch -Pattern $line -Quiet) { return } } catch {} }
        Add-Content -Path $q -Value $line -Encoding utf8
    } catch {}
}

function Build-StatsData([string]$tip, [string]$grup, [int]$n, [string]$sehir = '') {
    # "Şartlı 4/DHÖW" -> "Şartlı 4" (kutuphane anahtari tek olsun)
    $tip = (([string]$tip) -split '/')[0].Trim()
    # irk/kum-cim artik ayri boyut olarak hesaplaniyor; anahtar tip + adet + hipodrom
    $grup = ''
    if (-not $sehir) { $sehir = '-1' }
    if ($sehir -notmatch '^-?\d+$') { $sehir = '-1' }
    $script:StatsKosu = $n
    $dir = Join-Path $script:Root 'veri'
    if (-not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
    $key = (($tip + '|' + $grup + '|' + $n + '|' + $sehir) -replace '[^A-Za-z0-9]+', '_')
    if ($key.Length -gt 120) { $key = $key.Substring(0, 120) }
    $file = Join-Path $dir ('istatistik-' + $key + '.json')
    $cached = $null
    if (Test-Path $file) { try { $cached = ([System.IO.File]::ReadAllText($file, [System.Text.Encoding]::UTF8) | ConvertFrom-Json) } catch {} }
    $cachedValid = ($cached -and -not $cached.Hata -and ($cached.PSObject.Properties.Name -contains 'KadroSayisi') -and [int]$cached.KosuSayisi -gt 0)
    $fresh = $false
    if ($cachedValid -and (Test-Path $file)) { $age = ((Get-Date) - (Get-Item $file).LastWriteTime).TotalHours; if ($age -lt 24) { $fresh = $true } }
    if ($fresh -and $cached) { return $cached }
    if (-not $queryOptions) { if ($cachedValid) { return $cached }; return [pscustomobject]@{ Hata='TJK sorgu seçenekleri alınamadı'; Tip=$tip; Grup=$grup; KosuSayisi=0 } }
    $race = [pscustomobject]@{ Tip=$tip; Grup=$grup; SehirId=-1; Sehir=''; Pist=''; Mesafe=0 }
    try {
        $res = Build-RaceTypeStats $race $queryOptions $sehir
        if ($res -and -not $res.Hata -and [int]$res.KosuSayisi -gt 0) {
            try { [System.IO.File]::WriteAllText($file, ($res | ConvertTo-Json -Depth 12 -Compress), (New-Object System.Text.UTF8Encoding($false))) } catch {}
        }
        if ($res.Hata -and $cachedValid) { $cached | Add-Member -NotePropertyName Bayat -NotePropertyValue $true -Force; return $cached }
        return $res
    } catch {
        if ($cachedValid) { $cached | Add-Member -NotePropertyName Bayat -NotePropertyValue $true -Force; return $cached }
        return [pscustomobject]@{ Hata=$_.Exception.Message; Tip=$tip; Grup=$grup; KosuSayisi=0 }
    }
}

function Build-StatsShell([string]$tip, [string]$grup, [int]$n, [string]$sehir, [string]$sadi) {
    if (-not $sehir) { $sehir = '-1' }
    if ($sehir -notmatch '^-?\d+$') { $sehir = '-1' }
    if (-not $sadi) { $sadi = '' }
    $tipH = [System.Net.WebUtility]::HtmlEncode($tip)
    $grupH = [System.Net.WebUtility]::HtmlEncode($grup)
    $sadiH = [System.Net.WebUtility]::HtmlEncode($sadi)
    $tipJ = ($tip -replace '\\','\\') -replace "'","\'"
    $grupJ = ($grup -replace '\\','\\') -replace "'","\'"
    $sadiJ = ($sadi -replace '\\','\\') -replace "'","\'"
    $h = @'
<!doctype html><html lang="tr"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1"><title>__TIPH__ istatistigi</title>
<style>
:root{--bg:#0f1115;--card:#171a21;--line:#2a3140;--ink:#e8ecf1;--mut:#98a2b3;--acc:#4f8cff}
*{box-sizing:border-box}body{margin:0;font-family:system-ui,Segoe UI,Arial;background:var(--bg);color:var(--ink);padding:18px}
.wrap{max-width:900px;margin:0 auto}a{color:var(--acc)}h1{font-size:20px;margin:0 0 4px}h3{font-size:13px;color:var(--mut);margin:18px 0 8px;text-transform:uppercase;letter-spacing:.04em}
.kpi{display:flex;flex-wrap:wrap;gap:16px;background:var(--card);border:1px solid var(--line);border-radius:12px;padding:14px}
.kpi div{font-size:13px;color:var(--mut)}.kpi b{display:block;font-size:20px;color:var(--ink)}
.bars{display:flex;flex-direction:column;gap:6px}.bar{display:grid;grid-template-columns:80px 1fr 60px;gap:8px;align-items:center;font-size:13px}
.track{height:10px;background:var(--card);border:1px solid var(--line);border-radius:6px;overflow:hidden}.fill{display:block;height:100%;background:var(--acc)}
.num{text-align:right}.mut{color:var(--mut);font-size:12px}
.sizebtn{display:inline-block;padding:6px 11px;border:1px solid var(--line);border-radius:8px;text-decoration:none;margin:0 6px 10px 0;font-size:12.5px;color:var(--ink)}
.sizebtn.sel{background:var(--acc);color:#fff;border-color:var(--acc)}
.tbl{width:100%;border-collapse:collapse;font-size:12px}.tbl th,.tbl td{padding:6px 7px;border-bottom:1px solid var(--line);text-align:left;white-space:nowrap}.tbl th{color:var(--mut);font-size:11px;text-transform:uppercase}
</style></head><body><div class="wrap">
<p><a href="/">&larr; Uygulamaya dön</a></p>
<h1>__TIPH__ __GRUPH__ — genel istatistik</h1>
<div id="citiesel" style="margin-bottom:6px"></div>
<div id="sizes"></div>
<div id="content"><p class="mut">Yükleniyor… TJK geçmiş sonuçları çekiliyor.</p></div>
</div>
<script>
var P = { tip: '__TIPJ__', grup: '__GRUPJ__', n: __N__, sehir: '__SEHIRJ__', sadi: '__SADIJ__' };
function esc(s){ return String(s==null?'':s).replace(/[&<>"]/g, function(c){ return {'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]; }); }
function citiesel(){
  var e=document.getElementById('citiesel'); if(!e) return;
  var h='<a class="sizebtn'+(P.sehir==='-1'?' sel':'')+'" href="#" onclick="yukle(null,\'-1\');return false">Tüm Türkiye</a>';
  if(P.sadi){ h+='<a class="sizebtn'+(P.sehir!=='-1'?' sel':'')+'" href="#" onclick="yukle(null,\''+P.sehir+'\');return false">'+esc(P.sadi)+'</a>'; }
  e.innerHTML=h;
}
function sizes(cur){ var e=document.getElementById('sizes'), h=''; [20,50,100,200].forEach(function(k){ h+='<a class="sizebtn'+(k===cur?' sel':'')+'" href="#" onclick="yukle('+k+');return false">son '+k+'</a>'; }); e.innerHTML=h; }
var S=null, F={y:'',i:''};
async function yukle(k, sehir){
  if(k) P.n=k;
  if(sehir) P.sehir=sehir;
  citiesel(); sizes(P.n);
  var c=document.getElementById('content');
  c.innerHTML='<p class="mut">Yükleniyor… TJK geçmiş sonuçları çekiliyor (birkaç saniye).</p>';
  try{
    var r=await fetch('/api/istatistik?tip='+encodeURIComponent(P.tip)+'&grup='+encodeURIComponent(P.grup)+'&n='+P.n+'&sehir='+encodeURIComponent(P.sehir));
    if(!r.ok) throw new Error('HTTP '+r.status);
    S=await r.json();
    if(S.Hazirlaniyor){ c.innerHTML='<p class="mut">İstatistik hazırlanıyor… TJK geçmişi taranıyor. Sayfa birkaç dakika içinde kendiliğinden dolacak.</p>'; setTimeout(function(){ yukle(); }, 15000); return; }
    if(S.Hata){ c.innerHTML='<p class="mut">'+esc(S.Hata)+'</p><p class="mut">TJK yoğun olabilir. <a href="#" onclick="yukle();return false">Tekrar dene</a></p>'; return; }
    var h='';
    h+='<div id="filters" style="margin:2px 0 12px"></div>';
    h+='<div class="kpi" id="kpis"></div><p class="mut" id="kadro"></p>';
    h+='<h3>Kazananların HP dağılımı (yüksekten düşüğe)</h3><div class="bars" id="bands"></div><p class="mut" id="sknote"></p>';
    h+='<h3 id="thead">Kazananlar</h3>';
    h+='<div style="overflow:auto"><table class="tbl"><thead><tr><th>Tarih</th><th>Şehir</th><th>Pist</th><th>Irk</th><th>Kazanan</th><th>Yaş</th><th>Kilo</th><th>AGF</th><th>AGF sıra</th><th>HP</th><th>HP sıra</th><th>Favori</th><th>Alan</th></tr></thead><tbody id="tbody"></tbody></table></div>';
    c.innerHTML=h;
    F={y:'',i:''}; render();
  }catch(e){ c.innerHTML='<p class="mut">Veri alınamadı ('+esc(e.message)+'). <a href="#" onclick="yukle('+k+');return false">Tekrar dene</a></p>'; }
}
function num(x){ if(typeof x==='number') return x; var v=parseFloat(String(x==null?'':x).replace(',','.')); return isNaN(v)?null:v; }
function pct(a,b){ return b?Math.round(1000*a/b)/10:null; }
function avg(arr){ var v=arr.filter(function(x){return x!=null;}); if(!v.length) return null; var t=0; v.forEach(function(x){t+=x;}); return Math.round(10*t/v.length)/10; }
function rowsFilt(){ var A=(S&&S.Kosular)||[]; return A.filter(function(r){
  var p=String(r.Pist||'').toLowerCase();
  if(F.y==='kum' && p.indexOf('kum')<0) return false;
  if(F.y==='cim' && !(p.indexOf('çim')>=0 || p.indexOf('cim')>=0)) return false;
  if(F.y==='sentetik' && p.indexOf('sentetik')<0) return false;
  if(F.i && String(r.Irk||'')!==F.i) return false;
  return true;
}); }
function fbtn(kind,val,label){ var sel=(F[kind]===val); return '<a class="sizebtn'+(sel?' sel':'')+'" href="#" onclick="setF(\''+kind+'\',\''+val+'\');return false">'+label+'</a>'; }
function setF(kind,val){ F[kind]=val; render(); }
function render(){
  var rows=rowsFilt();
  document.getElementById('filters').innerHTML=
    fbtn('y','','Tümü')+fbtn('y','kum','Kum')+fbtn('y','cim','Çim')+fbtn('y','sentetik','Sentetik')+' &nbsp;|&nbsp; '+
    fbtn('i','','Hepsi')+fbtn('i','Arap','Arap')+fbtn('i','İngiliz','İngiliz');
  var n=rows.length;
  var fav=rows.filter(function(r){return !!r.Favori;}).length;
  var yuk=rows.filter(function(r){return !!r.EnYuksekHP;}).length;
  var dus=rows.filter(function(r){return !!r.EnDusukHP;}).length;
  var kilos=rows.map(function(r){return num(r.Kilo);}).filter(function(x){return x!=null;});
  var yaslar=rows.map(function(r){ var m=String(r.Yas||'').match(/(\d{1,2})/); return m?parseInt(m[1],10):null; }).filter(function(x){return x!=null;});
  document.getElementById('kpis').innerHTML=
    '<div>Koşu sayısı <b>'+n+'</b></div>'
    +'<div>Favori kazanma <b>'+(n?('%'+pct(fav,n)):'—')+'</b></div>'
    +'<div>Sürpriz <b>'+(n?('%'+pct(n-fav,n)):'—')+'</b></div>'
    +'<div>En yüksek HP kazandı <b>'+(n?('%'+pct(yuk,n)):'—')+'</b></div>'
    +'<div>En düşük HP kazandı <b>'+(n?('%'+pct(dus,n)):'—')+'</b></div>'
    +'<div>Ort. sıklet <b>'+(kilos.length?(avg(kilos)+' kg'):'—')+'</b></div>'
    +'<div>Ort. yaş <b>'+(yaslar.length?avg(yaslar):'—')+'</b></div>';
  document.getElementById('kadro').textContent='Tam kadro verisi: '+(S.KadroSayisi||0)+'/'+(S.KosuSayisi||0)+' koşu · seçili kırılımda '+n+' koşu';
  var hps=rows.map(function(r){return num(r.HP);}).filter(function(x){return x!=null;});
  var hb='';
  if(hps.length){ var mn=Math.min.apply(null,hps), mx=Math.max.apply(null,hps); var step=Math.max(5,Math.ceil((mx-mn)/6)); var b=Math.floor(mn/step)*step;
    while(b<=mx){ var lo=b, hi=b+step-1; var cnt=hps.filter(function(v){return v>=lo&&v<=hi;}).length; var oran=Math.round(1000*cnt/hps.length)/10;
      hb+='<div class="bar"><span>'+lo+'-'+hi+'</span><span class="track"><span class="fill" style="width:'+Math.round(oran)+'%"></span></span><span class="num">%'+oran+'</span></div>'; b+=step; } }
  else { hb='<p class="mut">Bu kırılımda HP verisi yok.</p>'; }
  document.getElementById('bands').innerHTML=hb;
  document.getElementById('sknote').textContent=kilos.length?('Kazanan sıklet aralığı: '+Math.min.apply(null,kilos)+'-'+Math.max.apply(null,kilos)+' kg'):'';
  document.getElementById('thead').textContent='Kazananlar ('+n+' koşu)';
  var tb=rows.map(function(r){ return '<tr><td>'+esc(r.Tarih)+'</td><td>'+esc(r.Sehir)+'</td><td>'+esc(r.Pist)+'</td><td>'+esc(r.Irk)+'</td><td>'+esc(r.At)+'</td><td>'+esc(r.Yas)+'</td><td>'+esc(r.Kilo)+'</td><td>'+(r.Agf!=null?('%'+r.Agf):'—')+'</td><td>'+(r.AgfRank!=null?r.AgfRank:'—')+'</td><td>'+esc(r.HP)+'</td><td>'+(r.HpSira!=null?r.HpSira:'—')+'</td><td>'+(r.Favori?'✔':'')+'</td><td>'+esc(r.Alan)+'</td></tr>'; });
  document.getElementById('tbody').innerHTML=tb.join('')||'<tr><td colspan="13" class="mut">Bu kırılımda koşu yok.</td></tr>';
}
yukle(P.n);
</script></body></html>
'@
    return $h.Replace('__TIPH__', $tipH).Replace('__GRUPH__', $grupH).Replace('__TIPJ__', $tipJ).Replace('__GRUPJ__', $grupJ).Replace('__N__', [string]$n).Replace('__SEHIRJ__', $sehir).Replace('__SADIJ__', $sadiJ)
}

function Build-StatsHtml([string]$tip, [string]$grup, [int]$n) {
    $script:StatsKosu = $n
    $race = [pscustomobject]@{ Tip=$tip; Grup=$grup; SehirId=-1; Sehir=''; Pist=''; Mesafe=0 }
    $s = $null
    if ($queryOptions) { $s = Build-RaceTypeStats $race $queryOptions }
    $tipH = [System.Net.WebUtility]::HtmlEncode($tip)
    $grupH = [System.Net.WebUtility]::HtmlEncode($grup)
    $parts = @()
    $parts += "<!doctype html><html lang='tr'><head><meta charset='utf-8'><meta name='viewport' content='width=device-width, initial-scale=1'><title>$tipH istatistigi</title><style>"
    $parts += ":root{--bg:#0f1115;--card:#171a21;--line:#2a3140;--ink:#e8ecf1;--mut:#98a2b3;--acc:#4f8cff}*{box-sizing:border-box}body{margin:0;font-family:system-ui,Segoe UI,Arial;background:var(--bg);color:var(--ink);padding:20px}.wrap{max-width:900px;margin:0 auto}a{color:var(--acc)}h1{font-size:20px;margin:0 0 4px}h3{font-size:13px;color:var(--mut);margin:18px 0 8px;text-transform:uppercase;letter-spacing:.04em}.kpi{display:flex;flex-wrap:wrap;gap:16px;background:var(--card);border:1px solid var(--line);border-radius:12px;padding:14px}.kpi div{font-size:13px;color:var(--mut)}.kpi b{display:block;font-size:20px;color:var(--ink)}.bars{display:flex;flex-direction:column;gap:6px}.bar{display:grid;grid-template-columns:80px 1fr 60px;gap:8px;align-items:center;font-size:13px}.track{height:10px;background:var(--card);border:1px solid var(--line);border-radius:6px;overflow:hidden}.fill{display:block;height:100%;background:var(--acc)}.num{text-align:right}.mut{color:var(--mut);font-size:12px}.sizebtn{display:inline-block;padding:5px 10px;border:1px solid var(--line);border-radius:8px;text-decoration:none;margin:0 6px 0 0;font-size:12.5px}.sizebtn.sel{background:var(--acc);color:#fff;border-color:var(--acc)}</style></head><body><div class='wrap'>"
    $parts += "<p><a href='/'>&larr; Uygulamaya dön</a></p><h1>$tipH $grupH — genel istatistik</h1><p class='mut'>TJK geçmiş sonuçlarından hesaplanır.</p>"
    $sizes = ''
    foreach ($k in @(20,50,100,200)) { $cls = if ($k -eq $n) { 'sel' } else { '' }; $sizes += "<a class='sizebtn $cls' href='/istatistik?tip=$([uri]::EscapeDataString($tip))&grup=$([uri]::EscapeDataString($grup))&n=$k'>son $k</a>" }
    $parts += "<div>$sizes</div>"
    if (-not $s) { $parts += "<p class='mut'>İstatistik servisi kullanılamıyor (TJK sorgu seçenekleri alınamadı).</p>" }
    elseif ($s.Hata) { $parts += "<p class='mut'>" + ([System.Net.WebUtility]::HtmlEncode($s.Hata)) + "</p><p class='mut'>TJK yoğun olabilir; birkaç dakika sonra tekrar deneyin.</p>" }
    else {
        $sk = $s.Siklet
        $skTxt = if ($sk) { ("{0} kg" -f $sk.Ortalama) } else { '—' }
        $skRange = if ($sk) { ("{0}-{1} kg" -f $sk.Min, $sk.Max) } else { '—' }
        $parts += "<div class='kpi'>"
        $parts += "<div>Koşu sayısı <b>$($s.KosuSayisi)</b></div>"
        $parts += "<div>Favori kazanma <b>%$($s.FavoriKazanma)</b></div>"
        $parts += "<div>Sürpriz <b>%$($s.SurprizOran)</b></div>"
        $parts += "<div>En yüksek HP kazandı <b>%$($s.EnYuksekHpKazanma)</b></div>"
        $parts += "<div>En düşük HP kazandı <b>%$($s.EnDusukHpKazanma)</b></div>"
        $parts += "<div>Ort. sıklet <b>$skTxt</b></div>"
        $parts += "</div><h3>Kazananların HP dağılımı (yüksekten düşüğe)</h3><div class='bars'>"
        if ($s.HpBantlari) { foreach ($b in $s.HpBantlari) { $w = [int]$b.Oran; $parts += "<div class='bar'><span>$($b.Bant)</span><span class='track'><span class='fill' style='width:$w%'></span></span><span class='num'>%$($b.Oran)</span></div>" } }
        $parts += "</div><p class='mut'>Kazanan sıklet aralığı: $skRange · Örnek: son $($s.KosuSayisi) koşu</p>"
    }
    $parts += "</div></body></html>"
    return ($parts -join "")
}

function Build-RaceAnalysis([string]$tarih, [string]$kosu, [string]$sehirId, [string]$ad) {
    $esc = [uri]::EscapeDataString($tarih)
    $sid = 0; [void][int]::TryParse($sehirId, [ref]$sid)
    $progUrl = "$script:Base/TR/YarisSever/Info/Page/GunlukYarisProgrami?QueryParameter_Tarih=$esc"
    $prog = Invoke-Tjk $progUrl -TimeoutSec 10
    if (-not $prog) { return $null }
    $cities = Get-Cities $prog
    $city = @($cities | Where-Object { [int]$_.SehirId -eq $sid }) | Select-Object -First 1
    if (-not $city) { return $null }
    $cityUrl = $city.Url -replace '[&?]Era=[^&]*', ''
    $html = Invoke-Tjk $cityUrl -TimeoutSec 15
    if (-not $html) { return $null }
    $races = Get-Races $html $sid $city.Ad
    $race = @($races | Where-Object { $_.KosuKodu -eq $kosu }) | Select-Object -First 1
    if (-not $race) { return $null }
    $winsMap = Get-JockeyWinsMap @($race) $sid $Donem $winsCache
    $jr = Rank-FromMap $race.Atlar $winsMap
    $surf = Build-Surfaces $race $histCache
    $h2h = Build-HeadToHead $race $histCache
    $combo = Build-JockeyHorseCombo $race $histCache
    $form = Build-FormScore $race
    $cons = Build-Consensus $race $surf $h2h $jr $form
    $stats = $null
    if ($queryOptions) { $stats = Build-StatsData ([string]$race.Tip) ([string]$race.Grup) 50 }
    $yorum = Build-Verdict $race $stats $cons
    $genis = Build-ExtendedVerdict $race $form $jr $surf $combo
    return [pscustomobject]@{
        KosuKodu  = $kosu
        JokeySira = $jr; Yuzey = $surf; IkiliGecme = $h2h; JokeyAt = $combo
        Form      = $form; Konsensus = $cons; Istatistik = $stats; Yorum = $yorum; GenisYorum = $genis
    }
}

function Send-File($stream, [string]$path, [string]$ctype) {
    if (-not (Test-Path $path)) { Send-Response $stream 404 'text/plain' 'bulunamadi'; return }
    $bytes = [System.IO.File]::ReadAllBytes($path)
    $header = "HTTP/1.1 200 OK`r`nContent-Type: $ctype`r`nContent-Length: $($bytes.Length)`r`nConnection: close`r`nCache-Control: public, max-age=86400`r`n`r`n"
    $hb = [System.Text.Encoding]::ASCII.GetBytes($header)
    $stream.Write($hb, 0, $hb.Length)
    $stream.Write($bytes, 0, $bytes.Length)
    $stream.Flush()
}

function Send-Response($stream, [int]$code, [string]$ctype, [string]$body) {
    $bytes = [System.Text.Encoding]::UTF8.GetBytes($body)
    $status = if ($code -eq 200) { 'OK' } elseif ($code -eq 404) { 'Not Found' } else { 'OK' }
    $header = "HTTP/1.1 $code $status`r`nContent-Type: $ctype; charset=utf-8`r`nContent-Length: $($bytes.Length)`r`nConnection: close`r`nCache-Control: no-store`r`n`r`n"
    $hb = [System.Text.Encoding]::ASCII.GetBytes($header)
    $stream.Write($hb, 0, $hb.Length)
    if ($bytes.Length -gt 0) { $stream.Write($bytes, 0, $bytes.Length) }
    $stream.Flush()
}

function Handle-Client($client, $config) {
    $stream = $null
    try {
        $script:queryOptions = $config.QueryOptions
        $script:Donem = $config.Donem
        $script:StatsYil = $config.StatsYil
        $script:StatsKosu = $config.StatsKosu
        $script:Sehir = $config.Sehir
        $script:VarsayilanTarih = $config.VarsayilanTarih
        $script:tplPath = $config.TplPath
        $script:winsCache = @{}; $script:histCache = @{}; $script:statsCache = @{}
        $script:Hizli = $false; $script:MaxRace = 0; $script:UseCacheOnly = $false
        $script:Yabanci = $false

        $stream = $client.GetStream()
        $reader = New-Object System.IO.StreamReader($stream, [System.Text.Encoding]::ASCII)
        $requestLine = $reader.ReadLine()
        if (-not $requestLine) { return }
        while ($true) { $ln = $reader.ReadLine(); if ($null -eq $ln -or $ln -eq '') { break } }

        $parts = $requestLine.Split(' ')
        $target = if ($parts.Count -ge 2) { $parts[1] } else { '/' }
        $path = $target; $qs = ''
        $qi = $target.IndexOf('?')
        if ($qi -ge 0) { $path = $target.Substring(0, $qi); $qs = $target.Substring($qi + 1) }

        $params = @{}
        if ($qs) {
            foreach ($kv in $qs.Split('&')) {
                if (-not $kv) { continue }
                $p = $kv.Split('=', 2)
                $k = [uri]::UnescapeDataString($p[0])
                $v = if ($p.Count -gt 1) { [uri]::UnescapeDataString($p[1]) } else { '' }
                $params[$k] = $v
            }
        }

        if ($path -eq '/favicon.ico') {
            Send-Response $stream 204 'text/plain' ''
        } elseif ($path -eq '/saglik') {
            Send-Response $stream 200 'text/plain' 'ok'
        } elseif ($path -eq '/manifest.webmanifest') {
            Send-File $stream (Join-Path $script:Root 'manifest.webmanifest') 'application/manifest+json'
        } elseif ($path -eq '/sw.js') {
            Send-File $stream (Join-Path $script:Root 'sw.js') 'application/javascript'
        } elseif ($path -match '^/icons/icon-\d+\.png$') {
            Send-File $stream (Join-Path (Join-Path $script:Root 'icons') (Split-Path $path -Leaf)) 'image/png'
        } elseif ($path -eq '/' -or $path -eq '/index.html') {
            $tarih = if ($params['tarih']) { Resolve-Tarih $params['tarih'] } elseif ($VarsayilanTarih) { Resolve-Tarih $VarsayilanTarih } else { Resolve-Tarih '' }
            Send-Response $stream 200 'text/html' (Build-Shell $tarih)
        } elseif ($path -eq '/api/gun') {
            $tarih = Resolve-Tarih $params['tarih']
            $hizli = ($params['hizli'] -eq '1')
            $script:Yabanci = ($params['yabanci'] -eq '1')
            $payload = Build-Payload $tarih $hizli
            if (-not $payload) {
                $d = [datetime]::MinValue
                $isToday = $false
                if ([datetime]::TryParseExact($tarih, 'dd/MM/yyyy', $script:CIV, 'None', [ref]$d)) { $isToday = ($d.Date -eq (Get-Date).Date) }
                if ($isToday) {
                    $y = (Get-Date).AddDays(-1).ToString('dd/MM/yyyy', $script:CIV)
                    $alt = Build-Payload $y $hizli
                    if ($alt) { $alt | Add-Member -NotePropertyName Uyari -NotePropertyValue ("Bugünün verisi alınamadı (TJK). Dün ($y) gösteriliyor.") -Force; $payload = $alt }
                }
            }
            if (-not $payload) { $payload = Get-LatestPayloadCache }
            if (-not $payload) { Send-Response $stream 503 'application/json' '{"error":"veri alinamadi"}' }
            else { Send-Response $stream 200 'application/json' (To-JsonSafe $payload) }
        } elseif ($path -eq '/api/analiz') {
            $tarih = Resolve-Tarih $params['tarih']
            $res = Build-RaceAnalysis $tarih $params['kosu'] $params['sehir'] $params['ad']
            if ($res) { Send-Response $stream 200 'application/json' (To-JsonSafe $res) }
            else { Send-Response $stream 404 'application/json' '{"error":"analiz yok"}' }
        } elseif ($path -eq '/istatistik') {
            $tip = $params['tip']; $grup = $params['grup']; $n = 100
            if ($params['n']) { [void][int]::TryParse($params['n'], [ref]$n) }
            Send-Response $stream 200 'text/html' (Build-StatsShell $tip $grup $n $params['sehir'] $params['sadi'])
        } elseif ($path -eq '/api/jokeyler') {
            $sid = $params['sehir']
            $yil = if ($params['yil']) { $params['yil'] } else { $Donem }
            if ($yil -eq '-1' -or $yil -notmatch '^\d{4}$') { $yil = $Donem }
            # 12 aylik pencere: eski yil dosyalari birikmesin (400 gunden eski jokey onbellegi silinir)
            try {
                $cd = Join-Path $script:Root 'cache'
                if (Test-Path $cd) {
                    Get-ChildItem $cd -Filter 'jokey-*' -File -ErrorAction SilentlyContinue |
                        Where-Object { ((Get-Date) - $_.LastWriteTime).TotalDays -gt 400 } |
                        Remove-Item -Force -ErrorAction SilentlyContinue
                }
            } catch {}
            $tab = @(Get-JockeyTable $sid $yil)
            $kullanilan = $yil
            if (-not $tab.Count -and $yil -ne '-1') { $tab = @(Get-JockeyTable $sid '-1'); $kullanilan = '-1' }
            Send-Response $stream 200 'application/json' (To-JsonSafe ([pscustomobject]@{ SehirId=$sid; Yil=$kullanilan; Jokeyler=$tab }))
        } elseif ($path -eq '/api/istatistik') {
            $tip = $params['tip']; $grup = $params['grup']; $n = 100
            if ($params['n']) { [void][int]::TryParse($params['n'], [ref]$n) }
            $uret = ($params['uret'] -eq '1')
            $dosya = Get-StatsFile $tip $grup $n $params['sehir']
            if (-not $uret -and -not (Test-StatsFresh $dosya)) {
                $bayat = Get-StatsCache $dosya
                if ($bayat) {
                    Add-StatsQueue $tip $n $params['sehir']
                    $bayat | Add-Member -NotePropertyName Bayat -NotePropertyValue $true -Force
                    Send-Response $stream 200 'application/json' (To-JsonSafe $bayat)
                } else {
                    Add-StatsQueue $tip $n $params['sehir']
                    $bekle = [pscustomobject]@{
                        Hazirlaniyor = $true
                        Tip          = (([string]$tip) -split '/')[0].Trim()
                        Grup         = $grup
                        KosuSayisi   = 0
                        Mesaj        = 'İstatistik hazırlanıyor; birkaç dakika içinde otomatik gelecek.'
                    }
                    Send-Response $stream 200 'application/json' (To-JsonSafe $bekle)
                }
            } else {
                Send-Response $stream 200 'application/json' (To-JsonSafe (Build-StatsData $tip $grup $n $params['sehir']))
            }
        } elseif ($path -eq '/api/turler') {
            $liste = @()
            if ($script:queryOptions -and $script:queryOptions.KosuCinsi) {
                $liste = @($script:queryOptions.KosuCinsi | Where-Object { $_.Ad -and [string]$_.Id -ne '-1' } | ForEach-Object {
                    [pscustomobject]@{
                        Ad    = [string]$_.Ad
                        Id    = [string]$_.Id
                        Hazir = (Test-StatsFresh (Get-StatsFile ([string]$_.Ad) '' 50 '-1'))
                    }
                })
            }
            Send-Response $stream 200 'application/json' (To-JsonSafe ([pscustomobject]@{ Turler = $liste }))
        } elseif ($path -eq '/api/sehirler') {
            $liste = @()
            if ($script:queryOptions -and ($script:queryOptions.PSObject.Properties.Name -contains 'Sehir') -and $script:queryOptions.Sehir) {
                $liste = @($script:queryOptions.Sehir | Where-Object { $_.Ad -and [string]$_.Id -ne '-1' } | ForEach-Object {
                    [pscustomobject]@{ Id = [string]$_.Id; Ad = [string]$_.Ad }
                })
            }
            Send-Response $stream 200 'application/json' (To-JsonSafe ([pscustomobject]@{ Sehirler = $liste }))
        } elseif ($path -eq '/turler') {
            Send-File $stream (Join-Path $script:Root 'turler.html') 'text/html; charset=utf-8'
        } elseif ($path -eq '/jokeyler') {
            Send-File $stream (Join-Path $script:Root 'jokeyler.html') 'text/html; charset=utf-8'
        } elseif ($path -eq '/tam') {
            $tarih = Resolve-Tarih $params['tarih']
            $hizli = ($params['hizli'] -eq '1')
            Send-Response $stream 200 'text/html' (Build-Html $tarih $hizli)
        } else {
            Send-Response $stream 404 'text/plain' 'bulunamadi'
        }
    } catch {
        try { if ($stream) { Send-Response $stream 500 'text/plain' 'sunucu hatasi' } } catch {}
    } finally {
        try { $client.Close() } catch {}
    }
}
