param(
    [string]$OutDir = '',
    [int]$MaxAnaliz = 12,
    [int]$MaxStats = 2,
    [int]$ButceDk = 30,
    [switch]$CacheOnly,
    [switch]$SadeceVeri
)
# Bulut yayini icin statik site uretir: public/ klasoru (Cloudflare Pages'e yuklenir).
$ErrorActionPreference = 'Continue'
$root = $PSScriptRoot
# Agir isler (istatistik sifirdan hesaplama) icin zaman butcesi: is akisinin
# 45 dk sinirina takilmamak icin varsayilan 30 dk. Butce bitince istatistik
# sifirdan hesaplanmaz, onbellek/kuyruk yoluna dusulur.
$script:Baslangic = Get-Date
$script:Butce = $script:Baslangic.AddMinutes($ButceDk)
function ButceVar { return ((Get-Date) -lt $script:Butce) }
function Gecen([string]$etiket) {
    Log ("[{0}] gecen {1} dk, kalan butce {2} dk" -f $etiket, [math]::Round(((Get-Date) - $script:Baslangic).TotalMinutes, 1), [math]::Round(($script:Butce - (Get-Date)).TotalMinutes, 1))
}
if (-not $OutDir) { $OutDir = Join-Path $root 'public' }
$script:Root = $root
. (Join-Path $root 'lib.ps1')
# analiz.ps1'in param varsayilanlari bunlari sifirlayabilir; dot-source sonrasi tekrar ayarla
$global:UseCacheOnly = [bool]$CacheOnly
$script:UseCacheOnly = [bool]$CacheOnly

$script:Donem = (Get-Date).Year.ToString()
$script:StatsYil = 5
$script:StatsKosu = 50
$script:Sehir = ''
$script:VarsayilanTarih = ''
$script:tplPath = Join-Path $root 'sablon.html'
$script:Yabanci = $false
$script:Hizli = $false
$script:MaxRace = 0
$script:winsCache = @{}; $script:histCache = @{}; $script:statsCache = @{}

$veri = Join-Path $OutDir 'veri'
if (-not (Test-Path $veri)) { New-Item -ItemType Directory -Path $veri -Force | Out-Null }
$yerelVeri = Join-Path $root 'veri'
if (-not (Test-Path $yerelVeri)) { New-Item -ItemType Directory -Path $yerelVeri -Force | Out-Null }

function YazJson([string]$path, $obj) {
    $json = ($obj | ConvertTo-Json -Depth 14 -Compress)
    [System.IO.File]::WriteAllText($path, $json, (New-Object System.Text.UTF8Encoding($false)))
}
function Log($m) { Write-Host ("[uret] " + $m) }

# --- sorgu secenekleri (tur listesi, sehir listesi) ---
$qHtml = Invoke-Tjk "$script:Base/TR/YarisSever/Query/Page/KosuSorgulama" -TimeoutSec 25 -Retries 2
$script:queryOptions = $null
if ($qHtml) {
    $script:queryOptions = [pscustomobject]@{
        KosuCinsi = Get-SelectOptions $qHtml 'QueryParameter_KosuCinsiId'
        Grup      = Get-SelectOptions $qHtml 'QueryParameter_GrupId'
        Irk       = Get-SelectOptions $qHtml 'QueryParameter_IrkId'
        Cinsiyet  = Get-SelectOptions $qHtml 'QueryParameter_Cinsiyet'
        Sehir     = Get-SelectOptions $qHtml 'QueryParameter_SehirId'
    }
    Log ("sorgu secenekleri alindi")
} else { Log "sorgu secenekleri ALINAMADI" }
$queryOptions = $script:queryOptions

# --- gun verisi ---
# GitHub sunucusu UTC saat kullanir; TJK Turkiye saatine (UTC+3) gore yayin yapar.
$bugun = [DateTime]::UtcNow.AddHours(3)
$tarih = $bugun.ToString('dd/MM/yyyy', $script:CIV)
$payload = Build-Payload $tarih $false
if (-not $payload) { Log "gun verisi alinamadi"; }
$gunDosya = $null
if ($payload) {
    $iso = $bugun.ToString('yyyyMMdd')
    $gunDosya = Join-Path $veri ('gun-' + $iso + '.json')
    YazJson $gunDosya $payload
    YazJson (Join-Path $veri 'gun.json') $payload
    $sr = 0; foreach ($d in @($payload.Tarihler)) { foreach ($c in @($d.Sehirler)) { $sr += @($c.Kosular).Count } }
    Log ("gun verisi yazildi: $sr kosu, $($payload.Tarihler.Count) tarih")
    Gecen 'gun verisi'
}

# --- turler / sehirler ---
if ($script:queryOptions) {
    $turler = @($script:queryOptions.KosuCinsi | Where-Object { $_.Ad -and [string]$_.Id -ne '-1' } | ForEach-Object {
        [pscustomobject]@{ Ad = [string]$_.Ad; Id = [string]$_.Id; Hazir = (Test-StatsFresh (Get-StatsFile ([string]$_.Ad) '' 50 '-1')) }
    })
    YazJson (Join-Path $veri 'turler.json') ([pscustomobject]@{ Turler = $turler })
    $sehirler = @($script:queryOptions.Sehir | Where-Object { $_.Ad -and [string]$_.Id -ne '-1' } | ForEach-Object {
        [pscustomobject]@{ Id = [string]$_.Id; Ad = [string]$_.Ad }
    })
    YazJson (Join-Path $veri 'sehirler.json') ([pscustomobject]@{ Sehirler = $sehirler })
    Log ("turler={0} sehirler={1}" -f $turler.Count, $sehirler.Count)
}

# --- jokey tablolari (gunun hipodromlari) ---
if ($payload) {
    $sidler = @{}
    foreach ($d in @($payload.Tarihler)) { foreach ($c in @($d.Sehirler)) { if ($c.SehirId) { $sidler[[string]$c.SehirId] = $true } } }
    foreach ($sid in $sidler.Keys) {
        $tab = @(Get-JockeyTable $sid $script:Donem)
        if (-not $tab.Count) { $tab = @(Get-JockeyTable $sid '-1') }
        YazJson (Join-Path $veri ('jokey-' + $sid + '-' + $script:Donem + '.json')) ([pscustomobject]@{ SehirId = $sid; Yil = $script:Donem; Jokeyler = $tab })
        Log ("jokey $sid -> " + $tab.Count)
    }
    # Tum Turkiye jokey tablosu (jokeyler sayfasindaki "Tum Turkiye" secimi icin)
    $tabTum = @(Get-JockeyTable '-1' $script:Donem)
    if (-not $tabTum.Count) { $tabTum = @(Get-JockeyTable '-1' '-1') }
    if ($tabTum.Count) {
        YazJson (Join-Path $veri ('jokey--1-' + $script:Donem + '.json')) ([pscustomobject]@{ SehirId = '-1'; Yil = $script:Donem; Jokeyler = $tabTum })
        Log ("jokey -1 -> " + $tabTum.Count)
    }
    Gecen 'jokey tablolari'
}

# --- kutuphane: istatistik dosyalarini kopyala ---
Get-ChildItem $yerelVeri -Filter 'istatistik-*.json' -ErrorAction SilentlyContinue | ForEach-Object {
    Copy-Item $_.FullName (Join-Path $veri $_.Name) -Force
}

# --- bugunun kosulari: analiz + eksik istatistik ---
$tipler = @{}
$analizSay = 0
if ($payload) {
    foreach ($d in @($payload.Tarihler)) {
        foreach ($c in @($d.Sehirler)) {
            foreach ($r in @($c.Kosular)) {
                $tip = (([string]$r.Tip) -split '/')[0].Trim()
                if ($tip) { $tipler[$tip] = $true }
                if ($analizSay -ge $MaxAnaliz) { continue }
                $kod = [string]$r.KosuKodu
                if (-not $kod) { continue }
                $hedef = Join-Path $veri ('analiz-' + ($kod -replace '[^A-Za-z0-9_-]', '_') + '.json')
                if (Test-Path $hedef) { continue }
                # istatistigi sifirdan hesaplamak dakikalar suruyor; butce dolduysa
                # onbellekten oku, eksikse kuyruga at (sayfa "Hazirlaniyor" gosterir)
                $res = Build-RaceAnalysis $tarih $kod ([string]$r.SehirId) ([string]$r.Sehir) -AllowStatsBuild:(ButceVar)
                if ($res) {
                    YazJson $hedef $res
                    Copy-Item $hedef (Join-Path $yerelVeri (Split-Path $hedef -Leaf)) -Force
                    $analizSay++
                }
            }
        }
    }
    Log ("analiz yazildi: " + $analizSay)
    Gecen 'analiz'
}

# eksik istatistikler (sinirli sayida, kutuphaneye de yazilir)
$statSay = 0
foreach ($tip in $tipler.Keys) {
    if ($statSay -ge $MaxStats) { break }
    if (-not (ButceVar)) { Log "butce doldu, istatistik hesaplama atlandi"; break }
    $dosya = Get-StatsFile $tip '' 50 '-1'
    if (Test-StatsFresh $dosya) { Copy-Item $dosya (Join-Path $veri (Split-Path $dosya -Leaf)) -Force; continue }
    $res = Build-StatsData $tip '' 50 '-1'
    if ($res -and -not $res.Hata -and [int]$res.KosuSayisi -gt 0) {
        $json = ($res | ConvertTo-Json -Depth 12 -Compress)
        [System.IO.File]::WriteAllText($dosya, $json, (New-Object System.Text.UTF8Encoding($false)))
        $leaf = Split-Path $dosya -Leaf
        YazJson (Join-Path $veri $leaf) $res
        $statSay++
        Log ("istatistik uretildi: $tip")
    }
    Gecen 'istatistik'
}

if ($SadeceVeri) { Log "sadece veri istendi, HTML atlandi"; return }

# --- statik sayfalar ---
foreach ($f in @('turler.html', 'jokeyler.html', 'sw.js', 'manifest.webmanifest')) {
    $kaynak = Join-Path $root $f
    if (Test-Path $kaynak) { Copy-Item $kaynak (Join-Path $OutDir $f) -Force }
    else { Log ("uyari: $f bulunamadi, atlandi") }
}

# Statik API koprusu: /api/... isteklerini veri/*.json dosyalarina cevirir
$shim = @'
<script>
(function(){var of=window.fetch;window.fetch=function(u,o){try{if(typeof u==='string'&&u.indexOf('/api/')===0){var q=new URL(u,location.href),p=q.pathname;
if(p==='/api/gun')u='veri/gun.json?t='+Date.now();
else if(p==='/api/turler')u='veri/turler.json';
else if(p==='/api/sehirler')u='veri/sehirler.json';
else if(p==='/api/jokeyler'){var s=q.searchParams.get('sehir')||'-1',y=q.searchParams.get('yil')||new Date().getFullYear();u='veri/jokey-'+s+'-'+y+'.json';}
else if(p==='/api/istatistik'){var t=(q.searchParams.get('tip')||'').split('/')[0].trim(),n=q.searchParams.get('n')||'50',se=q.searchParams.get('sehir')||'-1';u='veri/istatistik-'+(t+'||'+n+'|'+se).replace(/[^A-Za-z0-9]+/g,'_')+'.json';}
else if(p==='/api/analiz'){var k=(q.searchParams.get('kosu')||'').replace(/[^A-Za-z0-9_-]/g,'_');u='veri/analiz-'+k+'.json';}
}}catch(e){}return of(u,o);};})();
</script>
'@
function Add-Shim([string]$path) {
    if (-not (Test-Path $path)) { return }
    $t = [System.IO.File]::ReadAllText($path, [System.Text.Encoding]::UTF8)
    if ($t -match 'of=window.fetch') { return }
    $t = [regex]::Replace($t, '(<body[^>]*>)', ('$1' + $shim), 1)
    [System.IO.File]::WriteAllText($path, $t, (New-Object System.Text.UTF8Encoding($false)))
}
Add-Shim (Join-Path $OutDir 'turler.html')
Add-Shim (Join-Path $OutDir 'jokeyler.html')
if (Test-Path (Join-Path $root 'icons')) { Copy-Item (Join-Path $root 'icons') $OutDir -Recurse -Force }

$shell = Build-Html $tarih $false
[System.IO.File]::WriteAllText((Join-Path $OutDir 'index.html'), $shell, (New-Object System.Text.UTF8Encoding($false)))
[System.IO.File]::WriteAllText((Join-Path $OutDir '.nojekyll'), '', (New-Object System.Text.UTF8Encoding($false)))
Add-Shim (Join-Path $OutDir 'index.html')

# istatistik sayfalari (tur listesi icin HTML kabuk)
if ($script:queryOptions) {
    foreach ($t in @($script:queryOptions.KosuCinsi)) {
        $ad = [string]$t.Ad
        if (-not $ad -or [string]$t.Id -eq '-1') { continue }
        $html = Build-StatsShell $ad '' 50 '-1' ''
        $slug = (($ad + '||50|-1') -replace '[^A-Za-z0-9]+', '_')
        [System.IO.File]::WriteAllText((Join-Path $OutDir ('istatistik-' + $slug + '.html')), $html, (New-Object System.Text.UTF8Encoding($false)))
    }
}

# Pages Functions (API yonlendirici)
Copy-Item (Join-Path $root 'yayin-worker.js') (Join-Path $OutDir '_worker.js') -Force
Log ("hazir: " + $OutDir)
