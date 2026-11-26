param([int]$N = 50, [int]$Adet = 50)
# Kuyruktaki ve bugunun istatistik anahtarlarini arka planda uretir.
# Istek aninda TJK taramasi yapilmaz; sayfa "hazirlaniyor" der, bu script doldurur.

$ErrorActionPreference = 'Continue'
$root = $PSScriptRoot
$veri = Join-Path $root 'veri'
if (-not (Test-Path $veri)) { New-Item -ItemType Directory -Path $veri -Force | Out-Null }
$log = Join-Path $veri 'toplama.log'
function Log($m) { try { Add-Content -Path $log -Value ((Get-Date).ToString('yyyy-MM-dd HH:mm:ss') + ' ' + $m) -Encoding utf8 } catch {} }

function Get-Port {
    foreach ($pp in 8770..8795) {
        $c = Get-NetTCPConnection -LocalPort $pp -State Listen -ErrorAction SilentlyContinue
        if ($c) {
            $id = $c.OwningProcess | Select-Object -First 1
            $pr = Get-CimInstance Win32_Process -Filter "ProcessId=$id" -ErrorAction SilentlyContinue
            if ($pr -and $pr.CommandLine -match 'atyarisi') { return $pp }
        }
    }
    return $null
}
$port = Get-Port
if (-not $port) {
    Log 'doldur: sunucu yok, baslatiliyor'
    Start-Process powershell -ArgumentList @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', (Join-Path $root 'sunucu.ps1'), '-NoOpen') -WindowStyle Hidden | Out-Null
    Start-Sleep -Seconds 15
    $port = Get-Port
}
if (-not $port) { Log 'doldur: sunucu baslatilamadi'; return }
$Base = "http://127.0.0.1:$port"

# 1) Kuyruktakiler
$q = Join-Path $veri 'kuyruk.txt'
$isler = New-Object System.Collections.ArrayList
if (Test-Path $q) {
    foreach ($line in @(Get-Content $q -Encoding UTF8)) {
        $line = ([string]$line).Trim()
        if (-not $line) { continue }
        $p = $line.Split('|')
        if ($p.Count -lt 3) { continue }
        [void]$isler.Add([pscustomobject]@{
            Tip = $p[0]; N = [int]$p[1]; Sehir = $p[2]
            Grup = $(if ($p.Count -gt 3) { $p[3] } else { '' })
            Mesafe = $(if ($p.Count -gt 4) { $p[4] } else { '' })
            Cins = $(if ($p.Count -gt 5) { $p[5] } else { '' })
            Kaynak = 'kuyruk'
        })
    }
}

# 2) Bugunun kosu tipleri x hipodromlari (n=50) - onceden isit
$qAdet = $isler.Count   # kuyruk (kullanici istekleri) en onde kalsin
$tipler = @{}
try {
    $gunFile = Join-Path $veri ('gun-' + (Get-Date).ToString('yyyyMMdd') + '.json')
    # Bugunun gun verisi henuz olusmadiysa bir kez cek (tip/kategori matrisi bunu kullanir)
    if (-not (Test-Path $gunFile)) {
        try { Invoke-WebRequest -Uri "$Base/api/gun?hizli=1" -UseBasicParsing -TimeoutSec 120 | Out-Null } catch {}
    }
    if (Test-Path $gunFile) {
        $g = [System.IO.File]::ReadAllText($gunFile, [System.Text.Encoding]::UTF8) | ConvertFrom-Json
        $tipler = @{}
        $sehirler = @{}
        foreach ($d in @($g.Tarihler)) {
            foreach ($c in @($d.Sehirler)) {
                if ($c.SehirId) { $sehirler[[string]$c.SehirId] = $true }
                foreach ($r in @($c.Kosular)) {
                    $t = (([string]$r.Tip) -split '/')[0].Trim()
                    if ($t) { $tipler[$t] = $true }
                }
            }
        }
        $sehirList = @('-1') + @($sehirler.Keys)
        foreach ($t in $tipler.Keys) {
            foreach ($s in $sehirList) {
                $line = ('{0}|{1}|{2}|||' -f $t, $N, $s)
                $zaten = @($isler | Where-Object { ('{0}|{1}|{2}|{3}|{4}|{5}' -f $_.Tip, $_.N, $_.Sehir, $_.Grup, $_.Mesafe, $_.Cins) -eq $line }).Count -gt 0
                if (-not $zaten) { [void]$isler.Add([pscustomobject]@{ Tip = $t; N = $N; Sehir = $s; Grup = ''; Mesafe = ''; Cins = ''; Kaynak = 'isit' }) }
            }
        }
    }
} catch { Log ('doldur: gun verisi okunamadi ' + $_.Exception.Message) }

# 3) Tum kosu tipleri x kategoriler (Turkiye geneli, n=50)
# Kategori secilince dakikalarca beklememek icin onden hesaplanir.
try {
    $katlar = @()
    try {
        $kj = (Invoke-WebRequest -Uri "$Base/api/kategoriler" -UseBasicParsing -TimeoutSec 20).Content | ConvertFrom-Json
        $katlar = @(($kj.Grup) | ForEach-Object { [string]$_.Ad } | Where-Object { $_ })
    } catch { Log ('doldur: kategoriler alinamadi ' + $_.Exception.Message) }
    # Tip listesi TJK sorgu seceneklerinden; olmazsa bugunun tipleri kullanilir.
    $tipListe = @()
    try {
        $tj = (Invoke-WebRequest -Uri "$Base/api/turler" -UseBasicParsing -TimeoutSec 20).Content | ConvertFrom-Json
        $tipListe = @(($tj.Turler) | ForEach-Object { ([string]$_.Ad -split '/')[0].Trim() } |
            Where-Object { $_ -and $_ -notmatch '^(Handikap|Şartlı|Grup|YurtDışıBoş|Amatör|Amatör Binici|KADIN AMATÖR BİNİCİ|KADIN BİNİCİ)$' } | Select-Object -Unique)
    } catch {}
    if ($tipListe.Count -eq 0 -and $tipler.Count -gt 0) { $tipListe = @($tipler.Keys) }
    # Yaygin tipler once: kademeli dolumda en cok kullanilanlar hazir olsun
    if ($tipListe.Count -gt 0) {
        $tipListe = @($tipListe | Sort-Object @{Expression={ if ($_ -match '^(Handikap|Kısa Vade|KV|Şartlı|Maiden|G ?[123]$|G3 Handikap)') { 0 } else { 1 } }}, @{Expression={ $_ }})
    }
    # Once turlarin kendisi (Turkiye geneli, kategorisiz) hazir olsun.
    if ($tipListe.Count -gt 0) {
        $temelIsler = New-Object System.Collections.ArrayList
        foreach ($t in $tipListe) {
            $line = ('{0}|{1}|{2}|||' -f $t, $N, '-1')
            $zaten = @($isler | Where-Object { ('{0}|{1}|{2}|{3}|{4}|{5}' -f $_.Tip, $_.N, $_.Sehir, $_.Grup, $_.Mesafe, $_.Cins) -eq $line }).Count -gt 0
            if (-not $zaten) { [void]$temelIsler.Add([pscustomobject]@{ Tip = $t; N = $N; Sehir = '-1'; Grup = ''; Mesafe = ''; Cins = ''; Kaynak = 'temel' }) }
        }
        if ($temelIsler.Count -gt 0) { $isler.InsertRange([math]::Min($qAdet, $isler.Count), $temelIsler) }
        Log ('doldur: temel tur isleri {0}' -f $temelIsler.Count)
    }
    if ($tipListe.Count -gt 0 -and $katlar.Count -gt 0) {
        $katIsler = New-Object System.Collections.ArrayList
        foreach ($t in $tipListe) {
            foreach ($k in $katlar) {
                $line = ('{0}|{1}|{2}|{3}||' -f $t, $N, '-1', $k)
                $zaten = @($isler | Where-Object { ('{0}|{1}|{2}|{3}|{4}|{5}' -f $_.Tip, $_.N, $_.Sehir, $_.Grup, $_.Mesafe, $_.Cins) -eq $line }).Count -gt 0
                if (-not $zaten) { [void]$katIsler.Add([pscustomobject]@{ Tip = $t; N = $N; Sehir = '-1'; Grup = $k; Mesafe = ''; Cins = ''; Kaynak = 'kategori' }) }
            }
        }
        # Kategori isleri kuyrugun hemen ardindan calissin (onbellek amaci oncelikli)
        if ($katIsler.Count -gt 0) { $isler.InsertRange([math]::Min($qAdet, $isler.Count), $katIsler) }
        Log ('doldur: kategori isleri {0} ({1} tip x {2} kategori)' -f $katIsler.Count, $tipListe.Count, $katlar.Count)
    }
} catch { Log ('doldur: kategori isleri ' + $_.Exception.Message) }

if (-not $isler.Count) { Log 'doldur: is yok'; return }
Log ('doldur: {0} is' -f $isler.Count)

# Once tamamlanmis/gecerli olanlari atla
$basarili = 0; $atlanan = 0; $hata = 0; $sayac = 0
foreach ($is in $isler) {
    $sayac++
    if ($sayac -gt $Adet) { break }
    $isim = ('{0}|{1}|{2}|{3}|{4}|{5}' -f $is.Tip, $is.N, $is.Sehir, $is.Grup, $is.Mesafe, $is.Cins)
    $u = "$Base/api/istatistik?tip=$([uri]::EscapeDataString($is.Tip))&grup=$([uri]::EscapeDataString($is.Grup))&mesafe=$([uri]::EscapeDataString($is.Mesafe))&cins=$([uri]::EscapeDataString($is.Cins))&n=$($is.N)&sehir=$($is.Sehir)&uret=1"
    try {
        $r = Invoke-WebRequest -Uri $u -UseBasicParsing -TimeoutSec 900
        $o = $r.Content | ConvertFrom-Json
        if ($o -and -not $o.Hata -and [int]$o.KosuSayisi -gt 0) {
            $basarili++
            Log ("doldur OK: {0} / sehir {1} -> {2} kosu" -f $is.Tip, $is.Sehir, $o.KosuSayisi)
            if (Test-Path $q) {
                $kalan = @(Get-Content $q -Encoding UTF8 | Where-Object { $_.Trim() -ne $isim })
                Set-Content -Path $q -Value $kalan -Encoding utf8
            }
        } else {
            $atlanan++
            Log ("doldur ..: {0} / sehir {1} -> {2}" -f $is.Tip, $is.Sehir, $o.Hata)
            # TJK'da karsiligi olmayan tip: kuyrukta sonsuz tekrar etmesin
            if ([string]$o.Hata -match 'ayrıştırılamadı|ayristirilamadi|bulunamadı|bulunamadi') {
                if (Test-Path $q) {
                    $kalan = @(Get-Content $q -Encoding UTF8 | Where-Object { $_.Trim() -ne $isim })
                    Set-Content -Path $q -Value $kalan -Encoding utf8
                }
            }
        }
    } catch { $hata++; Log ("doldur X: {0} / sehir {1} -> {2}" -f $is.Tip, $is.Sehir, $_.Exception.Message) }
    Start-Sleep -Milliseconds 400
}
Log ("doldur bitti: ok={0} atlanan={1} hata={2}" -f $basarili, $atlanan, $hata)
