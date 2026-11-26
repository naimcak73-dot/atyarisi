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
# hizli=$true: gun verisi ~4 sn (tam veri TJK'da dakikalarca suruyor ve is akisini
# 45 dk sinirina yaklastiriyordu). Jokey sirasi ve detaylar istemci tarafinda
# /api/jokeyler + analiz dosyalarindan tamamlanir.
$payload = Build-Payload $tarih $true
if (-not $payload) { Log "gun verisi alinamadi"; }
$gunDosya = $null
$gunIsoList = @()
$gunYazilan = @{}
if ($payload) {
    $iso = $bugun.ToString('yyyyMMdd')
    $gunDosya = Join-Path $veri ('gun-' + $iso + '.json')
    YazJson $gunDosya $payload
    YazJson (Join-Path $veri 'gun.json') $payload
    $gunYazilan['gun-' + $iso + '.json'] = $true
    $sr = 0; foreach ($d in @($payload.Tarihler)) { foreach ($c in @($d.Sehirler)) { $sr += @($c.Kosular).Count } }
    Log ("gun verisi yazildi: $sr kosu, $($payload.Tarihler.Count) tarih")
    Gecen 'gun verisi'
    $gunIsoList = @($iso)

    # Yarin'in programi: TJK aksamlari yayinliyor. Sayfada "sonraki gun"
    # calissin diye ayri dosya. Veri gercekten yarina ait degilse (TJK henuz
    # yayinlamadiysa) dosya yazilmaz; istemci bilgi bandi gosterir.
    if (ButceVar) {
        try {
            $yarin = $bugun.AddDays(1)
            $yarinTarih = $yarin.ToString('dd/MM/yyyy', $script:CIV)
            $pYarin = Build-Payload $yarinTarih $true
            $beklenen = $yarin.ToString('yyyy-MM-dd')
            if ($pYarin -and ([string]$pYarin.TarihISO) -eq $beklenen) {
                $yi = $yarin.ToString('yyyyMMdd')
                YazJson (Join-Path $veri ('gun-' + $yi + '.json')) $pYarin
                $gunYazilan['gun-' + $yi + '.json'] = $true
                $gunIsoList += $yi
                $sy2 = 0; foreach ($d in @($pYarin.Tarihler)) { foreach ($c in @($d.Sehirler)) { $sy2 += @($c.Kosular).Count } }
                Log ("yarin gun verisi yazildi: $sy2 kosu ($yarinTarih)")
            } else { Log ("yarin gun verisi yok ($yarinTarih) - TJK henuz yayinlamamis olabilir") }
        } catch { Log ("yarin gun verisi alinamadi: " + $_.Exception.Message) }
        Gecen 'yarin'
    }

    # Dun'un programi: sayfada "Onceki gun" calissin diye ayri dosya.
    if (ButceVar) {
        try {
            $dun = $bugun.AddDays(-1)
            $dunTarih = $dun.ToString('dd/MM/yyyy', $script:CIV)
            $pDun = Build-Payload $dunTarih $true
            if ($pDun) {
                $di = $dun.ToString('yyyyMMdd')
                YazJson (Join-Path $veri ('gun-' + $di + '.json')) $pDun
                $gunYazilan['gun-' + $di + '.json'] = $true
                $gunIsoList += $di
                $srd = 0; foreach ($d in @($pDun.Tarihler)) { foreach ($c in @($d.Sehirler)) { $srd += @($c.Kosular).Count } }
                Log ("dun gun verisi yazildi: $srd kosu ($dunTarih)")
            }
        } catch { Log ("dun gun verisi alinamadi: " + $_.Exception.Message) }
        Gecen 'dun'
    }

    # Yurt disi gun verisi AYRI dosyaya: sayfada "Yurt disi" acilinca istemci
    # /api/gun?...&yabanci=1 ister; statik kopru bu dosyaya yonlendirilir.
    if (ButceVar) {
        try {
            $script:Yabanci = $true
            $yd = Build-Payload $tarih $true
            if ($yd) {
                $yi = $bugun.ToString('yyyyMMdd')
                YazJson (Join-Path $veri ('gun-' + $yi + '-yd.json')) $yd
                YazJson (Join-Path $veri 'gun-yabanci.json') $yd
                $gunYazilan['gun-' + $yi + '-yd.json'] = $true
                $gunYazilan['gun-yabanci.json'] = $true
                $sy = 0; foreach ($d in @($yd.Tarihler)) { foreach ($c in @($d.Sehirler)) { $sy += @($c.Kosular).Count } }
                Log ("yabanci gun yazildi: $sy kosu ($yi)")
                # Not: yurt disi kosularin ayrintili analizi uretilmiyor (TJK detay cagrilari
                # dakikalarca suruyor). Sayfa, analiz dosyasi yoksa eldeki kosu verisiyle acar.
            }
        } catch { Log ("yabanci gun alinamadi: " + $_.Exception.Message) }
        finally { $script:Yabanci = $false }
        Gecen 'yabanci gun'
    }
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
    # Jokey sayfasındaki il düğmeleri bugünkü hipodromlarla sınırlı kalmasın:
    # sorgu seçeneklerindeki tüm iller için tablo üret (statik yayında da çalışsın).
    if ($script:queryOptions -and $script:queryOptions.Sehir) {
        foreach ($s in @($script:queryOptions.Sehir)) {
            if ($s.Id -and [string]$s.Id -ne '-1') { $sidler[[string]$s.Id] = $true }
        }
    }
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
# Onceki calismalardan kalan gun/analiz dosyalarini yayina tasi (gh-pages her
# sefer sifirdan kuruluyor; bu kopya olmazsa gecmis gunler siteden silinir).
Get-ChildItem $yerelVeri -Filter 'gun-*.json' -ErrorAction SilentlyContinue | ForEach-Object {
    # Bu calismada taze uretilen dosyayi depodaki eski kopyayla EZME.
    if (-not $gunYazilan.ContainsKey($_.Name)) {
        Copy-Item $_.FullName (Join-Path $veri $_.Name) -Force
    }
}
Get-ChildItem $yerelVeri -Filter 'analiz-*.json' -ErrorAction SilentlyContinue | ForEach-Object {
    Copy-Item $_.FullName (Join-Path $veri $_.Name) -Force
}
Get-ChildItem $yerelVeri -Filter 'jokey-*.json' -ErrorAction SilentlyContinue | ForEach-Object {
    Copy-Item $_.FullName (Join-Path $veri $_.Name) -Force
}
# Bu calismada uretilen gun dosyalarini depoya geri yaz (is akisi main'e commit eder).
Get-ChildItem $veri -Filter 'gun-*.json' -ErrorAction SilentlyContinue | ForEach-Object {
    Copy-Item $_.FullName (Join-Path $yerelVeri $_.Name) -Force
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
# SIRA: once bugun kosulan turler, sonra TUM temel turler (Turkiye geneli).
# Boylece yayindaki "istatistik yok" durumu gunler degil, birkac kosuda kapanir.
$statSay = 0
$adayTipler = New-Object System.Collections.ArrayList
foreach ($t in @($tipler.Keys)) { if ($t) { [void]$adayTipler.Add([string]$t) } }
if ($script:queryOptions -and $script:queryOptions.KosuCinsi) {
    foreach ($o in @($script:queryOptions.KosuCinsi)) {
        $ad = (([string]$o.Ad) -split '/')[0].Trim()
        if (-not $ad -or [string]$o.Id -eq '-1') { continue }
        # TJK'da karsiligi olmayan / pist adi olan kayitlar
        if ($ad -match '^(YurtDışıBoş|YurtDisiBos|Tapeta)$') { continue }
        if (-not $adayTipler.Contains($ad)) { [void]$adayTipler.Add($ad) }
    }
}
Log ("istatistik adaylari: {0} (bugun {1} + tum turler)" -f $adayTipler.Count, @($tipler.Keys).Count)
foreach ($tip in $adayTipler) {
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

# Jokey x hipodrom matrisi: 10 hipodromun yillik tablosu statik olarak yazilir.
# Yayinda bu dosyalar dogrudan okunur; tarayici TJK'ya gitmez.
try {
    $jYil = (Get-Date).Year.ToString()
    $sidler = @('3','5','2','1','4','10','6','7','8','9')
    $kokVeri = Join-Path $root 'veri'
    if (-not (Test-Path $kokVeri)) { New-Item -ItemType Directory -Path $kokVeri -Force | Out-Null }
    $metaF = Join-Path $kokVeri 'jokey-yillar.json'
    $bugun = (Get-Date).ToString('yyyy-MM-dd')
    $meta = $null
    if (Test-Path $metaF) { try { $meta = [IO.File]::ReadAllText($metaF, [Text.Encoding]::UTF8) | ConvertFrom-Json } catch {} }
    $mevcut = 0
    foreach ($sid in $sidler) { if (Test-Path (Join-Path $kokVeri ("jokey-tablo-$jYil-$sid.json"))) { $mevcut++ } }
    $tam = ($meta -and [string]$meta.Tarih -eq $bugun -and $mevcut -eq $sidler.Count)

    if ($tam) {
        # Bugün zaten üretilmiş: TJK'ya 10 istek daha atmadan depodan kopyala.
        foreach ($sid in $sidler) { Copy-Item (Join-Path $kokVeri ("jokey-tablo-$jYil-$sid.json")) (Join-Path $veri ("jokey-tablo-$jYil-$sid.json")) -Force }
        Copy-Item $metaF (Join-Path $veri 'jokey-yillar.json') -Force
        Log ("jokey matrisi bugun hazir, depodan kopyalandi (" + $mevcut + "/10)")
    } else {
        $yazilan = 0
        foreach ($sid in $sidler) {
            if (-not (ButceVar)) { Log "butce doldu, jokey matrisi kismi yayinlanacak"; break }
            $tab = @(Get-JockeyTable $sid $jYil)
            if (-not $tab.Count -and $ButceVar) { $tab = @(Get-JockeyTable $sid '-1') }
            if ($tab.Count) {
                $obj = [pscustomobject]@{ SehirId=$sid; Yil=$jYil; Jokeyler=@($tab) }
                $json = ($obj | ConvertTo-Json -Depth 8 -Compress)
                [System.IO.File]::WriteAllText((Join-Path $kokVeri ("jokey-tablo-$jYil-$sid.json")), $json, (New-Object System.Text.UTF8Encoding($false)))
                [System.IO.File]::WriteAllText((Join-Path $veri ("jokey-tablo-$jYil-$sid.json")), $json, (New-Object System.Text.UTF8Encoding($false)))
                $yazilan++
            }
            Gecen 'jokey-matris'
        }
        # Sayfa yıl listesini bu dosyadan kurar; Tarih, gunluk atlamayi saglar.
        $metaObj = [pscustomobject]@{ Yillar=@($jYil); Hipodrom=$yazilan; Tarih=$bugun }
        $mj = ($metaObj | ConvertTo-Json -Compress)
        [System.IO.File]::WriteAllText($metaF, $mj, (New-Object System.Text.UTF8Encoding($false)))
        [System.IO.File]::WriteAllText((Join-Path $veri 'jokey-yillar.json'), $mj, (New-Object System.Text.UTF8Encoding($false)))
        Log ("jokey matrisi yazildi: " + $jYil + " (" + $yazilan + "/10 hipodrom)")
    }
} catch { Log ("jokey matrisi atlandi: " + $_.Exception.Message) }

if ($SadeceVeri) { Log "sadece veri istendi, HTML atlandi"; return }

# --- statik sayfalar ---
foreach ($f in @('turler.html', 'jokeyler.html', 'jokey-matris.html', 'sw.js', 'manifest.webmanifest', 'logo-kucuk.webp', 'logo.webp', 'logo.png')) {
    $kaynak = Join-Path $root $f
    if (Test-Path $kaynak) { Copy-Item $kaynak (Join-Path $OutDir $f) -Force }
    else { Log ("uyari: $f bulunamadi, atlandi") }
}

# Statik API koprusu: /api/... isteklerini veri/*.json dosyalarina cevirir
$shim = @'
<script>
(function(){var GL=__GUNLER__;var GYD=__YGUNLER__;window.__GL__=GL;window.__GYD__=GYD;var of=window.fetch;window.fetch=function(u,o){try{if(typeof u==='string'&&u.indexOf('/api/')===0){var q=new URL(u,location.href),p=q.pathname;
if(p==='/api/gun'){var tt=q.searchParams.get('tarih')||'',yd=q.searchParams.get('yabanci')==='1',is='';if(/^\d{4}-\d{2}-\d{2}$/.test(tt))is=tt.replace(/-/g,'');else if(/^\d{2}\/\d{2}\/\d{4}$/.test(tt)){var aa=tt.split('/');is=aa[2]+aa[1]+aa[0];}
 if(yd){if(is&&GYD.indexOf(is)>=0)u='veri/gun-'+is+'-yd.json?t='+Date.now();else if(!is)u='veri/gun-yabanci.json?t='+Date.now();else u='veri/gun-yd-yok.json';}
 else if(is){var d='veri/gun-'+is+'.json?t='+Date.now(),fb='veri/gun.json?t='+Date.now();if(GL.indexOf(is)>=0)u=d;else return of(d,o).then(function(r){return r.ok?r:of(fb,o);});}
 else u='veri/gun.json?t='+Date.now();}
else if(p==='/api/turler')u='veri/turler.json';
else if(p==='/api/sehirler')u='veri/sehirler.json';
else if(p==='/api/jokeyler'){var s=q.searchParams.get('sehir')||'-1',y=q.searchParams.get('yil')||new Date().getFullYear();u='veri/jokey-'+s+'-'+y+'.json';}
else if(p==='/api/jokey-tablo'){var s2=q.searchParams.get('sehir')||'-1',y2=q.searchParams.get('yil')||new Date().getFullYear();u='veri/jokey-tablo-'+y2+'-'+s2+'.json';}
else if(p==='/api/istatistik'){var t=(q.searchParams.get('tip')||'').split('/')[0].trim(),g=q.searchParams.get('grup')||'',n=q.searchParams.get('n')||'50',se=q.searchParams.get('sehir')||'-1',me=q.searchParams.get('mesafe')||'',ci=q.searchParams.get('cins')||'';u='veri/istatistik-'+(t+'|'+g+'|'+n+'|'+se+'|'+me+'|'+ci).replace(/[^A-Za-z0-9]+/g,'_')+'.json';}
else if(p==='/api/analiz'){var k=(q.searchParams.get('kosu')||'').replace(/[^A-Za-z0-9_-]/g,'_');u='veri/analiz-'+k+'.json';}
}}catch(e){}return of(u,o);};})();
</script>
'@
# Yayinlanan gunlerin ISO listesi (gun dosyalari): ["20261123","20261122"]
$gunJson = '["' + (@($gunIsoList) -join '","') + '"]'
$gydList = @(Get-ChildItem $veri -Filter 'gun-*-yd.json' -ErrorAction SilentlyContinue | ForEach-Object {
    ($_.BaseName -replace '^gun-', '' -replace '-yd$', '')
})
$gydJson = '["' + (@($gydList) -join '","') + '"]'
$shim = $shim.Replace('__GUNLER__', $gunJson)
$shim = $shim.Replace('__YGUNLER__', $gydJson)

function Add-Shim([string]$path) {
    if (-not (Test-Path $path)) { return }
    $t = [System.IO.File]::ReadAllText($path, [System.Text.Encoding]::UTF8)
    if ($t -match 'of=window.fetch') { return }
    $t = [regex]::Replace($t, '(<body[^>]*>)', ('$1' + $shim), 1)
    [System.IO.File]::WriteAllText($path, $t, (New-Object System.Text.UTF8Encoding($false)))
}
Add-Shim (Join-Path $OutDir 'turler.html')
Add-Shim (Join-Path $OutDir 'jokeyler.html')
Add-Shim (Join-Path $OutDir 'jokey-matris.html')
if (Test-Path (Join-Path $root 'icons')) { Copy-Item (Join-Path $root 'icons') $OutDir -Recurse -Force }

# hizli=$true: gomulu ilk veri de hizli payload olsun. Tam veri (hizli=$false)
# TJK'da dakikalarca surup yayini 45 dk sinirina yaklastiriyordu; detaylar
# analiz dosyalari + istemci tarafindaki /api/jokeyler ile tamamlanir.
$shell = Build-Html $tarih $true
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
