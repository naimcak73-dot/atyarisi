param([int]$IntervalMin = 1440)
# Sunucu icinde arka planda calisir: istatistik kutuphanesini doldurur/tazeler.
# - Kuyrukta is varsa 3 dakikada bir birkac isi isler (kullanici bekletilmez).
# - Tam tarama gunde bir yapilir (TJK'yi bos yere yormamak icin).
$ErrorActionPreference = 'Continue'
$root = $PSScriptRoot
$veri = Join-Path $root 'veri'
$kuyruk = Join-Path $veri 'kuyruk.txt'

function Test-ServerUp {
    foreach ($pp in 8770..8795) {
        $c = Get-NetTCPConnection -LocalPort $pp -State Listen -ErrorAction SilentlyContinue
        if ($c) { return $true }
    }
    return $false
}
function Get-KuyrukSayisi {
    try {
        if (-not (Test-Path $kuyruk)) { return 0 }
        return @(Get-Content $kuyruk -Encoding UTF8 | Where-Object { $_.Trim() }).Count
    } catch { return 0 }
}

$sonTam = [datetime]::MinValue
while ($true) {
    if (-not (Test-ServerUp)) { break }
    try {
        $tamVakti = (((Get-Date) - $sonTam).TotalHours -ge 24)
        if ($tamVakti) {
            try { & (Join-Path $root 'istatistik-topla.ps1') -N 50 } catch {}
            try { & (Join-Path $root 'istatistik-doldur.ps1') -Adet 300 } catch {}
            $sonTam = Get-Date
        } elseif ((Get-KuyrukSayisi) -gt 0) {
            try { & (Join-Path $root 'istatistik-doldur.ps1') -Adet 12 } catch {}
        }
    } catch {}
    Start-Sleep -Seconds 180
}
