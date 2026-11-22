param(
    [string]$Tarih = (Get-Date).ToString('dd/MM/yyyy', $script:CIV),
    [string]$Sehir = "",
    [switch]$Hizli,
    [switch]$UseCacheOnly,
    [switch]$NoOpen,
    [int]$StatsYil = 5,
    [int]$MaxRace = 0,
    [string]$Donem = (Get-Date).Year.ToString(),
    [int]$Aralik = 0,
    [int]$StatsKosu = 100,
    [switch]$Yabanci
)

$ErrorActionPreference = 'Stop'
$script:CIV = [System.Globalization.CultureInfo]::InvariantCulture
try { [Console]::OutputEncoding = [System.Text.Encoding]::UTF8; $OutputEncoding = [System.Text.Encoding]::UTF8 } catch {}

$script:Root      = $PSScriptRoot
$script:CacheBase = Join-Path $script:Root 'cache'
$script:RaporDir  = Join-Path $script:Root 'rapor'
$script:Base      = 'https://www.tjk.org'
$script:Headers   = @{
    'User-Agent'      = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36'
    'Accept'          = 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8'
    'Accept-Language' = 'tr-TR,tr;q=0.9,en;q=0.8'
}
$script:LastFetch = [datetime]::MinValue
$script:FailStreak = @{}
$script:Blocked = @{}

foreach ($d in @($script:CacheBase, $script:RaporDir)) { if (-not (Test-Path $d)) { New-Item -ItemType Directory -Path $d -Force | Out-Null } }

# ---------- yardımcılar ----------
function Html-Decode([string]$s) { if ($null -eq $s) { return '' }; try { return [System.Net.WebUtility]::HtmlDecode($s) } catch { return $s } }
function Url-Decode([string]$s) { if ([string]::IsNullOrWhiteSpace($s)) { return '' }; try { return [uri]::UnescapeDataString($s) } catch { return $s } }
function Strip-Tags([string]$s) {
    if ($null -eq $s) { return '' }
    $t = [regex]::Replace($s, '(?s)<script.*?</script>|<style.*?</style>', ' ')
    $t = [regex]::Replace($t, '(?s)<[^>]+>', ' ')
    $t = Html-Decode $t
    $t = [regex]::Replace($t, '\s+', ' ')
    return $t.Trim()
}
function M1([string]$text, [string]$pattern) {
    $m = [regex]::Match($text, $pattern, [System.Text.RegularExpressions.RegexOptions]::Singleline)
    if ($m.Success) { return $m.Groups[1].Value } else { return $null }
}
function Parse-Num([string]$s) {
    if ([string]::IsNullOrWhiteSpace($s)) { return $null }
    $t = ($s.Trim() -replace '\s','' -replace '%','')
    $t = ($t -replace '[^\d,\.]','')
    if ($t -eq '' -or $t -eq ',' -or $t -eq '.') { return $null }
    if (($t -match ',') -and ($t -match '\.')) { $t = ($t -replace '\.','') -replace ',','.' }
    elseif ($t -match ',') { $t = $t -replace ',','.' }
    $d = 0.0
    if ([double]::TryParse($t, [ref]$d)) { return $d } else { return $null }
}
function First-Num([string]$s) {
    if ([string]::IsNullOrWhiteSpace($s)) { return $null }
    $m = [regex]::Match($s, '\d{1,3}(?:[.,]\d+)?')
    if ($m.Success) { return (Parse-Num $m.Value) }
    return $null
}
# Tam sayi hucreleri (kosu sayisi, birincilik vb.) — 1.183 gibi binlik ayraclarini bozmadan okur
function First-Int([string]$s) {
    if ([string]::IsNullOrWhiteSpace($s)) { return $null }
    $t = ($s -replace '[^\d]', '')
    if (-not $t) { return $null }
    $v = 0
    if ([int]::TryParse($t, [ref]$v)) { return $v }
    return $null
}
function Sha1([string]$s) {
    $sha = [System.Security.Cryptography.SHA1]::Create()
    $bytes = [System.Text.Encoding]::UTF8.GetBytes($s)
    $hash = $sha.ComputeHash($bytes)
    ($hash | ForEach-Object { $_.ToString('x2') }) -join ''
}

function Invoke-Tjk {
    param([string]$Url, [int]$TimeoutSec = 10, [int]$Retries = 1, [string]$Cat = $null, [int]$TtlMin = 720)
    $dir = $script:CacheBase
    if (-not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
    $file = Join-Path $dir ((Sha1 $Url) + '.html')
    $stale = $null
    if (Test-Path $file) {
        $cached = [System.IO.File]::ReadAllText($file, [System.Text.Encoding]::UTF8)
        $age = ((Get-Date) - (Get-Item $file).LastWriteTime).TotalMinutes
        if ($age -lt $TtlMin) { return $cached }
        $stale = $cached
    }
    if ($UseCacheOnly) { if ($stale) { return $stale } else { return $null } }
    if (-not $Cat) { try { $Cat = ([uri]$Url).AbsolutePath } catch { $Cat = $Url } }
    if ($script:Blocked.ContainsKey($Cat) -and $script:Blocked[$Cat]) { return $null }
    for ($i = 1; $i -le $Retries; $i++) {
        try {
            $since = ((Get-Date) - $script:LastFetch).TotalMilliseconds
            if ($since -lt 400) { Start-Sleep -Milliseconds ([int](400 - $since)) }
            $resp = Invoke-WebRequest -Uri $Url -UseBasicParsing -TimeoutSec $TimeoutSec -Headers $script:Headers
            $script:LastFetch = Get-Date
            $bytes = $resp.RawContentStream.ToArray()
            $text = [System.Text.Encoding]::UTF8.GetString($bytes)
            [System.IO.File]::WriteAllText($file, $text, (New-Object System.Text.UTF8Encoding($false)))
            $script:FailStreak[$Cat] = 0
            return $text
        } catch {
            Write-Host ("   ! istek hatasi ({0}/{1}) [{2}]: {3}" -f $i, $Retries, $Cat, $_.Exception.Message) -ForegroundColor DarkYellow
            Start-Sleep -Milliseconds ([int]([math]::Min(3000, 600 * [math]::Pow(2, $i))))
        }
    }
    $script:Blocked[$Cat] = $true
    Write-Host ("   ! {0} yanit vermiyor - bu calistirmada atlaniyor." -f $Cat) -ForegroundColor DarkYellow
    if ($stale) { return $stale }
    return $null
}

# ---------- ayrıştırıcılar ----------
function Get-Cities([string]$html) {
    $list = @()
    $ms = [regex]::Matches($html, '<a id="([^"]*)" data-sehir-id="(\d+)" href="([^"]+)"[^>]*>([^<]+)</a>')
    foreach ($m in $ms) {
        $href = Html-Decode $m.Groups[3].Value
        $name = (Strip-Tags $m.Groups[4].Value)
        if ($href -notmatch 'SehirId=') { continue }
        if (-not $href.StartsWith('http')) { $href = $script:Base + $href }
        $list += [pscustomobject]@{
            SehirId = [int]$m.Groups[2].Value
            Ad      = $name
            Url     = $href
            Yabanci = ($name -match '(?i)(Avustralya|ABD|Kanada|Arjantin|Malezya|Guney Afrika|Güney Afrika|İngiltere|Birleşik Krallık|Birlesik Krallik|Krallık|Krallik|Fransa|Irlanda|İrlanda|Almanya|Brezilya|Sili|Uruguay|Italya|İtalya|Japonya|Katar|Suudi|Bae|Dubai|Hong Kong|Singapur|Yeni Zelanda|Peru|Venezuela|Panama|Meksika|Danimarka|Isvec|İsveç|Norvec|Norveç|Belcika|Belçika|Hollanda|Ispanya|İspanya|Rusya|Polonya|Cek|Çek|Avusturya|Isvicre|İsviçre)')
        }
    }
    return $list
}

# Gunluk sonuc sayfasindaki sehir baglantilari (program sayfasindan farkli isaretleme)
function Get-ResultCities([string]$html) {
    $list = @()
    if (-not $html) { return $list }
    $ms = [regex]::Matches($html, '<a[^>]+href="([^"]*Info/Sehir/GunlukYarisSonuclari\?[^"]*)"[^>]*>([^<]*)</a>')
    foreach ($m in $ms) {
        $href = Html-Decode $m.Groups[1].Value
        $sm = [regex]::Match($href, '[?&]SehirId=(\d+)')
        if (-not $sm.Success) { continue }
        $ad = (Strip-Tags $m.Groups[2].Value)
        $am = [regex]::Match($href, 'SehirAdi=([^&]+)')
        if ($am.Success) { try { $ad = [uri]::UnescapeDataString($am.Groups[1].Value) } catch {} }
        if (-not $ad) { continue }
        $list += [pscustomobject]@{ SehirId = [int]$sm.Groups[1].Value; Ad = $ad }
    }
    return $list
}

function Get-RaceWeather([string]$html) {
    $m = M1 $html '<span class="raceWeatherBrown">\s*([^<]+?)\s*</span>'
    if ($m) { return (Strip-Tags $m) } else { return '' }
}

function Get-Races([string]$html, [int]$sehirId, [string]$sehirAdi) {
    $races = @()
    # yurt disi kosularda kod negatif olur (or. -2722284)
    $panes = [regex]::Matches($html, '<div id="(-?\d+)" sehir="([^"]*)">')
    for ($p = 0; $p -lt $panes.Count; $p++) {
        $kosuKodu = $panes[$p].Groups[1].Value
        if ($kosuKodu -notmatch '^-?\d+$') { continue }
        $start = $panes[$p].Index
        $end = if ($p + 1 -lt $panes.Count) { $panes[$p + 1].Index } else { $html.Length }
        $chunk = $html.Substring($start, $end - $start)
        if ($chunk -notmatch ('id="kosubilgisi-' + $kosuKodu + '"')) { continue }

        $no = ''; $saat = ''
        $h = [regex]::Match($chunk, '<h3 class="race-no">\s*<a[^>]*>\s*(\d+)\.\s*Ko[şs]u\s*([\d\.:]*)')
        if ($h.Success) { $no = $h.Groups[1].Value; $saat = $h.Groups[2].Value.Trim() }
        if (-not $saat) {
            $h2 = [regex]::Match($chunk, '(\d+)\.\s*Ko[şs]u\s+([0-2]?\d[\.:][0-5]\d)')
            if ($h2.Success) { $no = $h2.Groups[1].Value; $saat = $h2.Groups[2].Value }
        }

        $cfgBlock = M1 $chunk '<h3 class="race-config">(.*?)</h3>'
        if (-not $cfgBlock) { $cfgBlock = '' }
        $cfgText = Strip-Tags $cfgBlock
        $tip = M1 $cfgBlock '<a[^>]*>\s*([^<]+?)\s*</a>'
        if (-not $tip) { $tip = ($cfgText -split ',')[0].Trim() }
        $tip = Strip-Tags $tip

        $pist = ''
        if ($cfgText -match '(?i)Çim') { $pist = 'Çim' } elseif ($cfgText -match '(?i)Kum') { $pist = 'Kum' }
        $mesafe = ''
        foreach ($nm in [regex]::Matches($cfgText, '\b(\d{3,4})\b')) {
            $v = [int]$nm.Groups[1].Value
            if ($v -ge 800 -and $v -le 4000) { $mesafe = $v; break }
        }
        $grup = ''
        # "3 ve Yukarı İngilizler - Dişi - Handikap 15" gibi hem virgül hem tire ayraçlı olabilir
        $gm = [regex]::Match($cfgText, '(?i)(\d+\s*(?:ya[şs]\s*)?(?:ve\s+yukar[ıi]|ya[şs]l[ıi])\s*(?:ingilizler|araplar|ingiliz|arap)[^,\-]*)')
        if ($gm.Success) {
            $grup = $gm.Groups[1].Value.Trim()
            # TJK secenegi "3 ve Yukarı ..." seklinde; "3 Yaş ve Yukarı ..." yazimini esitle
            $grup = [regex]::Replace($grup, '(?i)^(\d+)\s+ya[şs]\s+', '$1 ')
        }
        else {
            $gm2 = [regex]::Match($cfgText, '(?i),\s*(.+?)\s*(?:,|$)')
            if ($gm2.Success) { $grup = $gm2.Groups[1].Value.Trim() }
        }
        $cinsAd = ''; $cinsId = ''
        if ($cfgText -match '(?i)\bdi[şs]i') { $cinsAd = 'Dişi'; $cinsId = '2' }
        elseif ($cfgText -match '(?i)\berkek') { $cinsAd = 'Erkek'; $cinsId = '1' }

        # atlar
        $horses = @()
        $rows = [regex]::Matches($chunk, '(?s)<tr class="(?:odd|even)">(.*?)</tr>')
        foreach ($row in $rows) {
            $r = $row.Groups[1].Value
            $atId = M1 $r 'QueryParameter_AtId=(\d+)'
            if ($atId) {
                $atAdi = M1 $r ('AtId=' + $atId + '[^>]*>\s*([^<]+?)\s*<')
            } else {
                # yurt disi atlar TJK veritabaninda yok: AtId baglantisi olmaz, hucreden oku
                $atAdi = M1 $r 'GunlukYarisProgrami-AtAdi[^>]*>(.*?)</td>'
                if ($atAdi) { $atAdi = [regex]::Replace($atAdi, '(?s)<sup.*?</sup>', ' ') }
                $atAdi = Strip-Tags $atAdi
                if (-not $atAdi) { continue }
            }
            $jokeyId = M1 $r 'QueryParameter_JokeyId=(\d+)'
            $jokeyShort = M1 $r ('JokeyId=' + [regex]::Escape($jokeyId) + '[^>]*>([^<]+)')
            if (-not $jokeyShort) { $jokeyShort = M1 $r 'JokeAdi[^>]*>.*?JokeyId=\d+[^>]*>([^<]+)' }
            $jokeyFull = M1 $r 'JokeyIstatistikleri[^>]*title="([^"]+)"'
            if (-not $jokeyFull) { $jokeyFull = M1 $r 'QueryParameter_JokeyAdi=([^"&]+)' ; if ($jokeyFull) { $jokeyFull = Html-Decode $jokeyFull } }
            $agfBlock = M1 $r 'AGFORAN[^>]*>(.*?)</td>'
            $agfRank = $null; $agfPct = $null
            if ($agfBlock) {
                $rk = M1 $agfBlock '\((\d+)\)'; if ($rk) { $agfRank = [int]$rk }
                $pc = M1 $agfBlock '%([\d\.,]+)'; if ($pc) { $agfPct = Parse-Num $pc }
            }
            $formRaw = M1 $r 'Son6Yaris[^>]*>(.*?)</td>'
            $form = if ($formRaw) { (Strip-Tags $formRaw) -replace '\s','' } else { '' }
            $drc = M1 $r 'GunlukYarisProgrami-DERECE[^>]*>(?:(?!</td>).)*?<span[^>]*>\s*([\d\.:]+)\s*</span>'
            if (-not $drc) { $drc = Strip-Tags (M1 $r 'GunlukYarisProgrami-DERECE[^>]*>(.*?)</td>') }

            $yasTxt = Strip-Tags (M1 $r 'GunlukYarisProgrami-Yas[^>]*>\s*([^<]*)<')
            $yasNum = $null
            if ($yasTxt -match '(\d{1,2})') { $yasNum = [int]$Matches[1] }
            $horses += [pscustomobject]@{
                Sira     = [int](Parse-Num (M1 $r 'GunlukYarisProgrami-SiraId[^>]*>\s*([^<]*)<'))
                AtId     = $atId
                At       = (Strip-Tags $atAdi)
                Yas      = $yasTxt
                YasRakam = $yasNum
                Kilo     = Parse-Num (M1 $r 'GunlukYarisProgrami-Kilo[^>]*>\s*([^<]*)<')
                KiloText = (Strip-Tags (M1 $r 'GunlukYarisProgrami-Kilo[^>]*>\s*([^<]*)<'))
                Baba     = (Url-Decode (M1 $r 'QueryParameter_BabaAdi=([^"&]+)'))
                Anne     = (Url-Decode (M1 $r 'QueryParameter_AnneAdi=([^"&]+)'))
                KisrakBaba = (Url-Decode (M1 $r 'QueryParameter_KisrakBabaAdi=([^"&]+)'))
                JokeyId  = $jokeyId
                Jokey    = (Strip-Tags $jokeyShort)
                JokeyTam = (Strip-Tags $jokeyFull)
                St       = (Strip-Tags (M1 $r 'GunlukYarisProgrami-StartId[^>]*>\s*([^<]*)<'))
                HP       = Parse-Num (M1 $r 'GunlukYarisProgrami-Hc[^>]*>\s*([^<]*)<')
                AgfRank  = $agfRank
                AgfPct   = $agfPct
                Gny      = (Strip-Tags (M1 $r 'GunlukYarisProgrami-Gny[^>]*>\s*([^<]*)<'))
                Form     = $form
                Derece   = $drc
            }
        }
        if ($horses.Count -eq 0) { continue }

        $races += [pscustomobject]@{
            KosuKodu = $kosuKodu
            SehirId  = $sehirId
            Sehir    = $sehirAdi
            Cinsiyet = $cinsAd
            CinsiyetId = $cinsId
            KosuNo   = $no
            Saat     = $saat
            Tip      = $tip
            Grup     = $grup
            Mesafe   = $mesafe
            Pist     = $pist
            Atlar    = $horses
        }
    }
    return $races
}

function Get-JockeyWins([string]$html) {
    $m = [regex]::Match($html, 'JokeyIstatistikleri-derece1[^>]*>\s*(?:<[^>]+>\s*)*?([\d]+)')
    if (-not $m.Success) { $m = [regex]::Match($html, 'derece1[^>]*>\s*<a[^>]*>\s*([\d]+)') }
    if ($m.Success) { return [int]$m.Groups[1].Value } else { return $null }
}

# Atın geçmiş yarışları (AtKosuBilgileri) -> satırlar
function Get-HorseHistory([string]$html) {
    $out = @()
    $rows = [regex]::Matches($html, '(?s)<tr class="(?:odd|even)">(.*?)</tr>')
    foreach ($row in $rows) {
        $r = $row.Groups[1].Value
        $tds = [regex]::Matches($r, '(?s)<td[^>]*>(.*?)</td>')
        if ($tds.Count -lt 17) { continue }
        $kk = M1 $tds[0].Groups[1].Value '#(\d+)'
        $pistCell = Strip-Tags $tds[3].Groups[1].Value
        $pist = ''
        if ($pistCell -match '^\s*Ç' -or $pistCell -match '(?i)Çim') { $pist = 'Çim' }
        elseif ($pistCell -match '^\s*K' -or $pistCell -match '(?i)Kum') { $pist = 'Kum' }
        $sira = Parse-Num (Strip-Tags $tds[4].Groups[1].Value)
        $out += [pscustomobject]@{
            KosuKodu = $kk
            Tarih    = Strip-Tags $tds[0].Groups[1].Value
            Sehir    = Strip-Tags $tds[1].Groups[1].Value
            Mesafe   = Parse-Num (Strip-Tags $tds[2].Groups[1].Value)
            Pist     = $pist
            PistHam  = $pistCell
            Sira     = $sira
            Derece   = Strip-Tags $tds[5].Groups[1].Value
            Kilo     = Parse-Num (Strip-Tags $tds[6].Groups[1].Value)
            Jokey    = Strip-Tags $tds[8].Groups[1].Value
            Gny      = Strip-Tags $tds[10].Groups[1].Value
            HP       = Parse-Num (Strip-Tags $tds[16].Groups[1].Value)
        }
    }
    return $out
}

# Genel HTML tablo -> nesne listesi
function ConvertFrom-HtmlTable([string]$html) {
    $result = @()
    $tables = [regex]::Matches($html, '(?s)<table[^>]*>(.*?)</table>')
    foreach ($t in $tables) {
        $body = $t.Groups[1].Value
        $ths = @()
        foreach ($th in [regex]::Matches($body, '(?s)<th[^>]*>(.*?)</th>')) { $ths += (Strip-Tags $th.Groups[1].Value) }
        foreach ($tr in [regex]::Matches($body, '(?s)<tr[^>]*>(.*?)</tr>')) {
            $tds = [regex]::Matches($tr.Groups[1].Value, '(?s)<td[^>]*>(.*?)</td>')
            if ($tds.Count -eq 0) { continue }
            $o = @{}
            for ($i = 0; $i -lt $tds.Count; $i++) {
                $key = if ($ths.Count -gt $i) { $ths[$i] } else { "c$i" }
                $o[$key] = (Strip-Tags $tds[$i].Groups[1].Value)
            }
            $result += [pscustomobject]$o
        }
    }
    return $result
}

function Get-SelectOptions([string]$html, [string]$selectId) {
    $m = [regex]::Match($html, '(?s)<select[^>]*id="' + [regex]::Escape($selectId) + '".*?</select>')
    $opts = @()
    if ($m.Success) {
        foreach ($o in [regex]::Matches($m.Value, '<option[^>]*value="([^"]*)"[^>]*>([^<]*)</option>')) {
            $opts += [pscustomobject]@{ Id = $o.Groups[1].Value; Ad = (Strip-Tags $o.Groups[2].Value) }
        }
    }
    return $opts
}

# ---------- analizler ----------
# Hipodromun TUM jokeylerinin basari tablosu — tek istek (kosu, 1.-5., yuzdeler)
function Get-JockeyTable([string]$sehirId, [string]$yil) {
    $list = @()
    if (-not $sehirId) { return $list }
    if (-not $yil) { $yil = '-1' }
    $u = "$script:Base/TR/map/Query/Data/JokeyIstatistikleri?QueryParameter_YIL=$yil&QueryParameter_SehirId=$sehirId"
    $html = Invoke-Tjk $u -TimeoutSec 25 -Cat ('jokey-' + $sehirId + '-' + $yil)
    if (-not $html) { return $list }
    $parts = [regex]::Split($html, '(?i)<tr')
    foreach ($rh in $parts) {
        if ($rh -notmatch 'sorgu-JokeyIstatistikleri-JokeyAdi') { continue }
        $cell = { param($cls)
            $m = [regex]::Match($rh, '(?s)<td class="sorgu-JokeyIstatistikleri-' + $cls + '"[^>]*>(.*?)</td>')
            if ($m.Success) { return (Strip-Tags $m.Groups[1].Value) } else { return '' }
        }
        $ad = & $cell 'JokeyAdi'
        if (-not $ad) { continue }
        $jid = ''
        $jm = [regex]::Match($rh, 'QueryParameter_JokeyId=(\d+)')
        if ($jm.Success) { $jid = $jm.Groups[1].Value }
        $kosu = First-Int (& $cell 'KosuAdet')
        if (-not $kosu -or $kosu -le 0) { continue }
        $d1 = First-Int (& $cell 'derece1'); $d2 = First-Int (& $cell 'derece2'); $d3 = First-Int (& $cell 'derece3')
        $d4 = First-Int (& $cell 'derece4'); $d5 = First-Int (& $cell 'derece5')
        # tutarsiz satir (birincilik > kosu ya da yuzde > 100) = bozuk parse, atla
        if ($null -ne $d1 -and $d1 -gt $kosu) { continue }
        $p1 = First-Int (& $cell 'derece1yuzde')
        if ($null -ne $p1 -and $p1 -gt 100) { $p1 = $null }
        if ($null -eq $p1 -and $kosu -and $null -ne $d1 -and $kosu -gt 0) { $p1 = [int][math]::Round(100 * $d1 / $kosu) }
        $ilk5 = $null
        if ($kosu -and $kosu -gt 0) {
            $top = 0; foreach ($x in @($d1, $d2, $d3, $d4, $d5)) { if ($null -ne $x) { $top += $x } }
            $ilk5 = [int][math]::Round(100 * $top / $kosu)
        }
        $list += [pscustomobject]@{
            JokeyId = $jid
            Jokey   = $ad
            Kosu    = $kosu
            D1      = $d1
            D2      = $d2
            D3      = $d3
            D4      = $d4
            D5      = $d5
            P1      = $p1
            Ilk5    = $ilk5
        }
    }
    return @($list | Sort-Object @{Expression={ if ($null -ne $_.D1) { $_.D1 } else { -1 } }; Descending=$true }, @{Expression={ if ($null -ne $_.Kosu) { $_.Kosu } else { -1 } }; Descending=$true })
}

function Get-JockeyWinsMap($races, $sehirId, $donem, $winsCache) {
    if ($null -eq $winsCache) { $winsCache = @{} }
    $map = @{}
    $tkey = "tablo|$sehirId|$donem"
    if ($winsCache.ContainsKey($tkey)) { $tablo = @($winsCache[$tkey]) }
    else {
        $tablo = @()
        if (-not $UseCacheOnly) { $tablo = @(Get-JockeyTable $sehirId $donem) }
        $winsCache[$tkey] = $tablo
    }
    # JokeyId -> tablo satiri (Kosu, D1, P1, Ilk5 ...) — panelde gecmis/başari gosterimi icin
    foreach ($t in $tablo) { if ($t.JokeyId) { $map[$t.JokeyId] = $t } }
    return $map
}
function Jockey-Row($h, $winsMap, [bool]$withAt) {
    $t = $null; if ($h.JokeyId -and $winsMap -and $winsMap.ContainsKey($h.JokeyId)) { $t = $winsMap[$h.JokeyId] }
    $o = [ordered]@{
        Jokey = $h.Jokey; JokeyTam = $h.JokeyTam; JokeyId = $h.JokeyId
        Birincilik = $(if ($t) { $t.D1 } else { $null })
        Kosu       = $(if ($t) { $t.Kosu } else { $null })
        Galibiyet  = $(if ($t) { $t.P1 } else { $null })
        Ilk5       = $(if ($t) { $t.Ilk5 } else { $null })
    }
    if ($withAt) { $o['At'] = $h.At }
    return [pscustomobject]$o
}
function Rank-FromMap($horses, $winsMap) {
    $seen = @{}; $arr = @()
    foreach ($h in $horses) {
        if (-not $h.JokeyId -or [string]$h.JokeyId -eq '0' -or $seen.ContainsKey($h.JokeyId)) { continue }
        $seen[$h.JokeyId] = $true
        $arr += (Jockey-Row $h $winsMap $true)
    }
    return @($arr | Sort-Object @{Expression={ if ($_.Birincilik -ne $null) { $_.Birincilik } else { -1 } }; Descending=$true }, Jokey)
}
function City-JockeyRank($races, $winsMap) {
    $seen = @{}; $arr = @()
    foreach ($race in $races) {
        foreach ($h in $race.Atlar) {
            if (-not $h.JokeyId -or [string]$h.JokeyId -eq '0' -or $seen.ContainsKey($h.JokeyId)) { continue }
            $seen[$h.JokeyId] = $true
            $arr += (Jockey-Row $h $winsMap $false)
        }
    }
    return @($arr | Sort-Object @{Expression={ if ($_.Birincilik -ne $null) { $_.Birincilik } else { -1 } }; Descending=$true }, Jokey)
}

function Build-Surfaces($race, $histCache) {
    $res = @{}
    foreach ($h in $race.Atlar) {
        if (-not $h.AtId) { continue }
        if ($histCache.ContainsKey($h.AtId)) { $hist = $histCache[$h.AtId] }
        else {
            $hist = @()
            if (-not $UseCacheOnly) {
                $u = "$script:Base/TR/YarisSever/Query/ConnectedPage/AtKosuBilgileri?1=1&QueryParameter_AtId=$($h.AtId)&Era=today"
                $html = Invoke-Tjk $u -TimeoutSec 10
                if ($html) { $hist = Get-HorseHistory $html }
            }
            $histCache[$h.AtId] = $hist
        }
        $kum = @($hist | Where-Object { $_.Pist -eq 'Kum' })
        $cim = @($hist | Where-Object { $_.Pist -eq 'Çim' })
        $kW = @($kum | Where-Object { $_.Sira -eq 1 }).Count
        $cW = @($cim | Where-Object { $_.Sira -eq 1 }).Count
        $k3 = @($kum | Where-Object { $_.Sira -ge 1 -and $_.Sira -le 3 }).Count
        $c3 = @($cim | Where-Object { $_.Sira -ge 1 -and $_.Sira -le 3 }).Count
        $verdict = 'YOK'
        if ($kum.Count -ge 2 -and $cim.Count -ge 2) {
            $kr = 0.6 * ($kW / [double]$kum.Count) + 0.4 * ($k3 / [double]$kum.Count)
            $cr = 0.6 * ($cW / [double]$cim.Count) + 0.4 * ($c3 / [double]$cim.Count)
            if ([math]::Abs($kr - $cr) -lt 0.08) { $verdict = 'DENGELI' } elseif ($kr -gt $cr) { $verdict = 'KUM' } else { $verdict = 'CIM' }
        } elseif ($kum.Count -ge 2) { $verdict = 'KUM' } elseif ($cim.Count -ge 2) { $verdict = 'CIM' }
        $res[$h.AtId] = [pscustomobject]@{
            KumKosu=$kum.Count; KumBirinci=$kW; KumIlkUc=$k3
            CimKosu=$cim.Count; CimBirinci=$cW; CimIlkUc=$c3
            Karar=$verdict
        }
    }
    return $res
}

function Build-Distance($race, $histCache) {
    # Kosunun mesafesine gore gecmis: tam mesafe + en yakin (+-100 m).
    $res = @{}
    $target = 0
    try { $target = [int]$race.Mesafe } catch {}
    if ($target -le 0) { return $res }
    foreach ($h in $race.Atlar) {
        if (-not $h.AtId) { continue }
        $hist = if ($histCache.ContainsKey($h.AtId)) { $histCache[$h.AtId] } else { @() }
        $runs = @($hist | Where-Object { $_.Mesafe -gt 0 -and $null -ne $_.Sira })
        $exact = @($runs | Where-Object { [int]$_.Mesafe -eq $target })
        $near = @($runs | Where-Object { [int]$_.Mesafe -ne $target -and [math]::Abs([int]$_.Mesafe - $target) -le 100 })
        $nearDist = $null
        if ($near.Count -gt 0) { $nearDist = ($near | Sort-Object { [math]::Abs([int]$_.Mesafe - $target) } | Select-Object -First 1).Mesafe }
        $res[$h.AtId] = [pscustomobject]@{
            Mesafe        = $target
            Kosu          = $exact.Count
            Birinci       = @($exact | Where-Object { [int]$_.Sira -eq 1 }).Count
            IlkUc         = @($exact | Where-Object { [int]$_.Sira -ge 1 -and [int]$_.Sira -le 3 }).Count
            YakinMesafe   = $nearDist
            YakinKosu     = $near.Count
            YakinBirinci  = @($near | Where-Object { [int]$_.Sira -eq 1 }).Count
            IlkKez        = ($runs.Count -gt 0 -and $exact.Count -eq 0 -and $near.Count -eq 0)
            Gecmis        = $runs.Count
        }
    }
    return $res
}

# ---- Baba (sire) karnesi: kendi gecmisleri zayif atlarda gercek veri, kotali ----
$script:SireFile    = Join-Path $script:Root 'veri\baba.json'
$script:SireMax     = 300      # dosyada tutulacak en fazla baba
$script:SireMinKosu = 30       # gostermek icin en az kosu
$script:SireMinAt   = 3        # en az farkli at
function Dist-Band([int]$m) {
    if ($m -le 1200) { return '<=1200' }
    elseif ($m -le 1600) { return '1201-1600' }
    elseif ($m -le 2000) { return '1601-2000' }
    else { return '2001+' }
}
function Get-SireStore {
    $store = @{}
    if (-not (Test-Path $script:SireFile)) { return $store }
    try { $j = [IO.File]::ReadAllText($script:SireFile, [Text.Encoding]::UTF8) | ConvertFrom-Json } catch { return $store }
    if (-not $j -or -not $j.babalar) { return $store }
    foreach ($bp in $j.babalar.PSObject.Properties) {
        $sire = @{ Son = ''; Atlar = @{} }
        if ($bp.Value.Son) { $sire.Son = [string]$bp.Value.Son }
        if ($bp.Value.Atlar) {
            foreach ($ap in $bp.Value.Atlar.PSObject.Properties) {
                $a = $ap.Value; $m = @{}
                if ($a.Mesafe) { foreach ($mp in $a.Mesafe.PSObject.Properties) { $m[$mp.Name] = @{ K = [int]$mp.Value.K; B = [int]$mp.Value.B } } }
                $sire.Atlar[$ap.Name] = @{ Kosu = [int]$a.Kosu; Birinci = [int]$a.Birinci; KumK = [int]$a.KumK; KumB = [int]$a.KumB; CimK = [int]$a.CimK; CimB = [int]$a.CimB; Mesafe = $m }
            }
        }
        $store[$bp.Name] = $sire
    }
    return $store
}
function Save-SireStore($store) {
    if (-not $store -or $store.Count -eq 0) { return }
    $disk = Get-SireStore
    foreach ($k in $store.Keys) {
        if (-not $disk.ContainsKey($k)) { $disk[$k] = $store[$k]; continue }
        foreach ($aid in $store[$k].Atlar.Keys) { $disk[$k].Atlar[$aid] = $store[$k].Atlar[$aid] }
        if ($store[$k].Son) { $disk[$k].Son = $store[$k].Son }
    }
    # kota: eski (365 gun) ve en az kosulu olanlari at, en fazla 300 baba
    $temiz = @{}
    foreach ($k in $disk.Keys) {
        $s = $disk[$k]; $tot = 0; foreach ($aid in $s.Atlar.Keys) { $tot += [int]$s.Atlar[$aid].Kosu }
        $eski = $false
        try { if ($s.Son) { $d = [datetime]::ParseExact($s.Son, 'dd.MM.yyyy', $script:CIV); if (((Get-Date) - $d).TotalDays -gt 365) { $eski = $true } } } catch {}
        if (-not $eski -and $tot -gt 0) { $temiz[$k] = @{ Toplam = $tot; S = $s } }
    }
    $secili = @($temiz.Keys | Sort-Object @{ Expression = { $temiz[$_].Toplam }; Descending = $true } | Select-Object -First $script:SireMax)
    $babalar = [ordered]@{}
    foreach ($k in $secili) {
        $s = $temiz[$k].S
        $atl = [ordered]@{}
        foreach ($aid in $s.Atlar.Keys) {
            $a = $s.Atlar[$aid]
            $mes = [ordered]@{}
            foreach ($mb in $a.Mesafe.Keys) { $mes[$mb] = @{ K = [int]$a.Mesafe[$mb].K; B = [int]$a.Mesafe[$mb].B } }
            $atl[$aid] = @{ Kosu = [int]$a.Kosu; Birinci = [int]$a.Birinci; KumK = [int]$a.KumK; KumB = [int]$a.KumB; CimK = [int]$a.CimK; CimB = [int]$a.CimB; Mesafe = $mes }
        }
        $babalar[$k] = @{ Son = $s.Son; Atlar = $atl }
    }
    $obj = [pscustomobject]@{ guncelleme = (Get-Date).ToString('s'); babalar = $babalar }
    try {
        $dir = Split-Path $script:SireFile -Parent
        if (-not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
        [IO.File]::WriteAllText($script:SireFile, ($obj | ConvertTo-Json -Depth 10 -Compress), (New-Object Text.UTF8Encoding($false)))
    } catch {}
}
function Update-SireStore($race, $histCache, $store) {
    foreach ($h in @($race.Atlar)) {
        if (-not $h.AtId -or -not $h.Baba) { continue }
        if (-not $histCache.ContainsKey($h.AtId)) { continue }
        $hist = @($histCache[$h.AtId])
        $kosu = 0; $bir = 0; $kk = 0; $kb = 0; $ck = 0; $cb = 0; $mes = @{}; $son = ''
        foreach ($r in $hist) {
            if ($null -eq $r.Sira -or [int]$r.Mesafe -le 0) { continue }
            $kosu++
            $w = ([int]$r.Sira -eq 1)
            if ($w) { $bir++ }
            if ($r.Pist -eq 'Kum') { $kk++; if ($w) { $kb++ } }
            elseif ($r.Pist -eq 'Çim') { $ck++; if ($w) { $cb++ } }
            $b = Dist-Band ([int]$r.Mesafe)
            if (-not $mes.ContainsKey($b)) { $mes[$b] = @{ K = 0; B = 0 } }
            $mes[$b].K++
            if ($w) { $mes[$b].B++ }
            if ($r.Tarih -and $r.Tarih -gt $son) { $son = $r.Tarih }
        }
        if ($kosu -eq 0) { continue }
        $key = (Strip-Tags ([string]$h.Baba)).Trim()
        if (-not $key) { continue }
        if (-not $store.ContainsKey($key)) { $store[$key] = @{ Son = ''; Atlar = @{} } }
        $store[$key].Atlar[[string]$h.AtId] = @{ Kosu = $kosu; Birinci = $bir; KumK = $kk; KumB = $kb; CimK = $ck; CimB = $cb; Mesafe = $mes }
        if ($son) { $store[$key].Son = $son }
    }
}
function Build-SireHints($race, $histCache, $store) {
    $out = @{}
    $hedef = [int]$race.Mesafe
    $band = Dist-Band $hedef
    $say = 0
    foreach ($h in @($race.Atlar)) {
        if ($say -ge 3) { break }
        if (-not $h.AtId -or -not $h.Baba) { continue }
        $hist = @()
        if ($histCache.ContainsKey($h.AtId)) { $hist = @($histCache[$h.AtId]) }
        $ownSurf = @($hist | Where-Object { $_.Pist -eq 'Kum' -or $_.Pist -eq 'Çim' }).Count
        $ownDist = @($hist | Where-Object { [int]$_.Mesafe -eq $hedef }).Count
        if ($ownSurf -ge 2 -and $ownDist -ge 2) { continue }   # kendi gecmisi yeterliyse gosterme
        $key = (Strip-Tags ([string]$h.Baba)).Trim()
        if (-not $store.ContainsKey($key)) { continue }
        $s = $store[$key]
        $tot = 0; $bir = 0; $kk = 0; $kb = 0; $ck = 0; $cb = 0; $bK = 0; $bB = 0
        foreach ($aid in $s.Atlar.Keys) {
            $a = $s.Atlar[$aid]
            $tot += [int]$a.Kosu; $bir += [int]$a.Birinci
            $kk += [int]$a.KumK; $kb += [int]$a.KumB; $ck += [int]$a.CimK; $cb += [int]$a.CimB
            if ($a.Mesafe.ContainsKey($band)) { $bK += [int]$a.Mesafe[$band].K; $bB += [int]$a.Mesafe[$band].B }
        }
        if ($tot -lt $script:SireMinKosu -or $s.Atlar.Count -lt $script:SireMinAt) { continue }
        $out[[string]$h.AtId] = [pscustomobject]@{
            Baba     = $key
            Kosu     = $tot
            At       = $s.Atlar.Count
            KumPct   = $(if (($kk + $ck) -gt 0) { [int][math]::Round(100 * $kk / ($kk + $ck)) } else { $null })
            CimPct   = $(if (($kk + $ck) -gt 0) { [int][math]::Round(100 * $ck / ($kk + $ck)) } else { $null })
            Band     = $band
            BandKosu = $bK
            BandPct  = $(if ($bK -gt 0) { [int][math]::Round(100 * $bB / $bK) } else { $null })
        }
        $say++
    }
    return $out
}

function Build-HeadToHead($race, $histCache) {
    $byAt = @{}
    foreach ($h in $race.Atlar) {
        if (-not $h.AtId) { continue }
        $hist = if ($histCache.ContainsKey($h.AtId)) { $histCache[$h.AtId] } else { @() }
        $map = @{}
        foreach ($r in $hist) { if ($r.KosuKodu -and $null -ne $r.Sira) { $map[$r.KosuKodu] = $r.Sira } }
        $byAt[$h.AtId] = $map
    }
    $stats = @{}
    foreach ($h in $race.Atlar) { if ($h.AtId) { $stats[$h.AtId] = [pscustomobject]@{ At=$h.At; Gecme=0; Gecilme=0; Ortak=0 } } }
    $atl = @($race.Atlar)
    for ($i = 0; $i -lt $atl.Count; $i++) {
        for ($j = $i + 1; $j -lt $atl.Count; $j++) {
            $a = $atl[$i]; $b = $atl[$j]
            if (-not $a.AtId -or -not $b.AtId) { continue }
            $ma = $byAt[$a.AtId]; $mb = $byAt[$b.AtId]
            foreach ($k in $ma.Keys) {
                if ($mb.ContainsKey($k)) {
                    $stats[$a.AtId].Ortak++; $stats[$b.AtId].Ortak++
                    if ($ma[$k] -lt $mb[$k]) { $stats[$a.AtId].Gecme++; $stats[$b.AtId].Gecilme++ }
                    elseif ($mb[$k] -lt $ma[$k]) { $stats[$b.AtId].Gecme++; $stats[$a.AtId].Gecilme++ }
                }
            }
        }
    }
    $list = @()
    foreach ($k in $stats.Keys) {
        $s = $stats[$k]
        $list += [pscustomobject]@{ AtId=$k; At=$s.At; Gecme=$s.Gecme; Gecilme=$s.Gecilme; Net=($s.Gecme - $s.Gecilme); Ortak=$s.Ortak }
    }
    $list = @($list | Sort-Object @{Expression='Net';Descending=$true}, @{Expression='Gecme';Descending=$true})
    $best = $null
    if ($list.Count -gt 0 -and $list[0].Gecme -gt 0) { $best = $list[0].At }
    return [pscustomobject]@{ Liste=$list; EnCokGecen=$best }
}

function Build-JockeyHorseCombo($race, $histCache) {
    $res = @()
    foreach ($h in $race.Atlar) {
        if (-not $h.AtId) { continue }
        if ($null -eq $histCache -or -not $histCache.ContainsKey($h.AtId)) { continue }
        $hist = $histCache[$h.AtId]
        if (-not $h.Jokey) { continue }
        $same = @($hist | Where-Object { $_.Jokey -and ($_.Jokey -eq $h.Jokey) })
        $wins = @($same | Where-Object { $_.Sira -eq 1 }).Count
        if ($same.Count -gt 0) { $res += [pscustomobject]@{ At=$h.At; Jokey=$h.Jokey; Kosu=$same.Count; Birinci=$wins } }
    }
    return $res
}

function Build-FormScore($race) {
    # Son 6 yarış: 1 en iyi. Düşük ortalama = iyi form. Puan 0-100.
    $res = @()
    foreach ($h in $race.Atlar) {
        $f = ($h.Form -replace '[^\d]','')
        if ($f.Length -eq 0) { $res += [pscustomobject]@{ At=$h.At; Form=$f; Puan=$null; Ort=$null }; continue }
        $nums = @()
        foreach ($ch in $f.ToCharArray()) { $n = [int]::Parse([string]$ch); if ($n -eq 0) { $n = 10 }; if ($n -ge 1 -and $n -le 10) { $nums += $n } }
        if ($nums.Count -eq 0) { $res += [pscustomobject]@{ At=$h.At; Form=$f; Puan=$null; Ort=$null }; continue }
        $avg = ($nums | Measure-Object -Average).Average
        $puan = [math]::Round([math]::Max(0, [math]::Min(100, (10 - $avg) * 12)), 0)
        $res += [pscustomobject]@{ At=$h.At; Form=$f; Puan=$puan; Ort=[math]::Round($avg,2) }
    }
    return $res
}

function Build-Consensus($race, $surf, $h2h, $jockeyRank, $form) {
    $res = @()
    $maxWins = 1
    foreach ($j in $jockeyRank) { if ($j.Birincilik -ne $null -and $j.Birincilik -gt $maxWins) { $maxWins = $j.Birincilik } }
    $maxHP = 1
    foreach ($h in $race.Atlar) { if ($h.HP -ne $null -and $h.HP -gt $maxHP) { $maxHP = $h.HP } }
    foreach ($h in $race.Atlar) {
        $parts = [ordered]@{}
        $score = 0.0
        if ($h.HP -ne $null) { $s = 25 * ($h.HP / $maxHP); $parts['HP'] = [math]::Round($s,0); $score += $s }
        $fp = ($form | Where-Object { $_.At -eq $h.At } | Select-Object -First 1)
        if ($fp -and $fp.Puan -ne $null) { $parts['Form'] = $fp.Puan * 0.20; $score += $fp.Puan * 0.20 }
        $sf = if ($h.AtId) { $surf[$h.AtId] } else { $null }
        if ($sf) {
            $match = (($race.Pist -eq 'Kum' -and $sf.Karar -eq 'KUM') -or ($race.Pist -eq 'Çim' -and $sf.Karar -eq 'CIM'))
            if ($match) { $parts['Pist'] = 15; $score += 15 } elseif ($sf.Karar -eq 'DENGELI') { $parts['Pist'] = 7; $score += 7 }
        }
        $jr = ($jockeyRank | Where-Object { $_.JokeyId -eq $h.JokeyId } | Select-Object -First 1)
        if ($jr -and $jr.Birincilik -ne $null) { $js = 15 * ($jr.Birincilik / [math]::Max(1,$maxWins)); $parts['Jokey'] = [math]::Round($js,0); $score += $js }
        $h2 = ($h2h.Liste | Where-Object { $_.AtId -eq $h.AtId } | Select-Object -First 1)
        if ($h2 -and $h2.Net -gt 0) { $hs = [math]::Min(10, $h2.Net * 2); $parts['Gecme'] = $hs; $score += $hs }
        $res += [pscustomobject]@{ At=$h.At; Puan=[math]::Round($score,0); Bilesenler=$parts }
    }
    return @($res | Sort-Object @{Expression='Puan';Descending=$true})
}

function Build-ProgramVerdict($race, $form) {
    if (-not $race.Atlar -or $race.Atlar.Count -eq 0) { return $null }
    $fav = $null
    foreach ($h in $race.Atlar) { if ($h.AgfRank -eq 1) { $fav = $h; break } }
    if (-not $fav) {
        $best = $null
        foreach ($h in $race.Atlar) { $g = Parse-Num $h.Gny; if ($null -ne $g -and ($null -eq $best -or $g -lt $best)) { $best = $g; $fav = $h } }
    }
    $topHP = $null
    foreach ($h in $race.Atlar) { if ($null -ne $h.HP -and ($null -eq $topHP -or $h.HP -gt $topHP.HP)) { $topHP = $h } }
    $formTop = $null
    if ($form) { $formTop = @($form | Where-Object { $_.Puan -ne $null } | Sort-Object Puan -Descending | Select-Object -First 1)[0] }
    $hpOrder = @($race.Atlar | Where-Object { $null -ne $_.HP } | Sort-Object HP -Descending)
    $favHpRank = $null
    if ($fav) { $i = 0; foreach ($h in $hpOrder) { $i++; if ($h.At -eq $fav.At) { $favHpRank = $i; break } } }
    $uyum = 'belirsiz'
    if ($fav -and $favHpRank) {
        if ($favHpRank -eq 1) { $uyum = 'guclu-favori' }
        elseif ($favHpRank -le 3) { $uyum = 'orta-favori' }
        else { $uyum = 'zayif-favori' }
    }
    $p = @()
    if ($fav) { $p += ("Favori: {0} (AGF sırası {1})." -f $fav.At, $(if ($fav.AgfRank) { $fav.AgfRank } else { '—' })) }
    if ($topHP) { $p += ("En yüksek HP: {0} ({1})." -f $topHP.At, $topHP.HP) }
    if ($formTop) { $p += ("Form lideri: {0} (puan {1})." -f $formTop.At, $formTop.Puan) }
    if ($fav -and $favHpRank) { $p += ("Favori, HP sıralamasında {0}. sırada." -f $favHpRank) }
    $ages = @()
    foreach ($h in $race.Atlar) { if ($null -ne $h.YasRakam) { $ages += [int]$h.YasRakam } }
    if ($ages.Count -gt 0) {
        $avg = [math]::Round(($ages | Measure-Object -Average).Average, 1)
        $mn = ($ages | Measure-Object -Minimum).Minimum; $mx = ($ages | Measure-Object -Maximum).Maximum
        $p += ("Yaş: ort. {0} ({1}-{2})." -f $avg, $mn, $mx)
        if ($fav -and $null -ne $fav.YasRakam) { $p += ("Favori {0} yaşında." -f $fav.YasRakam) }
        if ($topHP -and $null -ne $topHP.YasRakam) { $p += ("En yüksek HP {0} yaşında." -f $topHP.YasRakam) }
        if ($formTop) { $fh = $race.Atlar | Where-Object { $_.At -eq $formTop.At } | Select-Object -First 1; if ($fh -and $null -ne $fh.YasRakam) { $p += ("Form lideri {0} yaşında." -f $fh.YasRakam) } }
    }
    if ($race.Grup -match '(\d+)\s*ve\s*Yukarı') { $p += ("Grup yaş kuralı: {0} ve Yukarı." -f $Matches[1]) }
    switch ($uyum) {
        'guclu-favori' { $p += 'Favori hem en yüksek HP' + "'" + 'te hem güçlü: favori mantığı (ganyan/plase) daha tutarlı.' }
        'orta-favori'  { $p += 'Favori HP' + "'" + 'de ilk 3' + "'" + 'te: favori + alternatif birlikte düşünülebilir.' }
        'zayif-favori' { $p += 'Favori HP' + "'" + 'de geride: sürpriz/alternatif atlara (yüksek HP + iyi form) bakmak mantıklı.' }
        default        { $p += 'Favori/HP sinyalleri net değil; dikkatli.' }
    }
    return [pscustomobject]@{
        Uyum = $uyum
        Metin = ($p -join ' ')
        Favori = $(if ($fav) { $fav.At } else { $null })
        EnYuksekHP = $(if ($topHP) { $topHP.At } else { $null })
        FormLideri = $(if ($formTop) { $formTop.At } else { $null })
    }
}

function Build-ExtendedVerdict($race, $form, $jockeyRank, $surf, $combo) {
    $base = Build-ProgramVerdict $race $form
    if (-not $base) { return $null }
    $p = @($base.Metin)
    $maxW = 0
    foreach ($j in $jockeyRank) { if ($j.Birincilik -ne $null -and [int]$j.Birincilik -gt $maxW) { $maxW = [int]$j.Birincilik } }
    if ($base.Favori -and $jockeyRank) {
        $favH = $race.Atlar | Where-Object { $_.At -eq $base.Favori } | Select-Object -First 1
        if ($favH) {
            $fj = $jockeyRank | Where-Object { $_.JokeyId -eq $favH.JokeyId } | Select-Object -First 1
            if ($fj -and $fj.Birincilik -ne $null) {
                if ($maxW -gt 0 -and [int]$fj.Birincilik -eq $maxW) { $p += ("Favorinin jokeyi {0} bu hipodromda dönemin en çok kazananı ({1})." -f $fj.Jokey, $fj.Birincilik) }
                elseif ([int]$fj.Birincilik -gt 0) { $p += ("Favorinin jokeyi {0} bu hipodromda {1} birincilik almış." -f $fj.Jokey, $fj.Birincilik) }
            }
        }
    }
    $surfaceName = if ($race.Pist -eq 'Kum') { 'KUM' } elseif ($race.Pist -eq 'Çim') { 'CIM' } else { '' }
    if ($surfaceName -and $surf) {
        $targets = @()
        if ($base.Favori) { $targets += $base.Favori }
        if ($base.EnYuksekHP -and $base.EnYuksekHP -ne $base.Favori) { $targets += $base.EnYuksekHP }
        foreach ($tn in $targets) {
            $th = $race.Atlar | Where-Object { $_.At -eq $tn } | Select-Object -First 1
            if (-not $th) { continue }
            $s = if ($th.AtId) { $surf[$th.AtId] } else { $null }
            if (-not $s) { continue }
            if ($s.Karar -eq $surfaceName) { $p += ("{0} bu pistte uygun (kum {1}/{2}, çim {3}/{4} tabela)." -f $tn, $s.KumIlkUc, $s.KumKosu, $s.CimIlkUc, $s.CimKosu) }
            elseif ($s.Karar -eq 'DENGELI') { $p += ("{0} iki pistte dengeli." -f $tn) }
            elseif ($s.Karar -ne 'YOK') { $p += ("{0} bu pistte zayıf olabilir (kum {1}/{2}, çim {3}/{4})." -f $tn, $s.KumIlkUc, $s.KumKosu, $s.CimIlkUc, $s.CimKosu) }
        }
    }
    if ($combo) { foreach ($c in $combo) { if ($c.Kosu -ge 2) { $p += ("{0}–{1} daha önce {2} kez birlikte, {3} birincilik." -f $c.At, $c.Jokey, $c.Kosu, $c.Birinci) } } }
    return [pscustomobject]@{ Uyum = $base.Uyum; Metin = ($p -join ' '); Favori = $base.Favori; EnYuksekHP = $base.EnYuksekHP; FormLideri = $base.FormLideri }
}

function Build-Verdict($race, $stats, $consensus) {
    if (-not $stats -or $stats.Hata) {
        return [pscustomobject]@{
            Seviye = 'yok'; Baslik = 'İstatistik yok'
            Metin = 'Bu koşu tipi için geçmiş istatistik alınamadı (TJK sorgu servisi). Yorum yapılamıyor.'
            SurprizAdayi = $null
        }
    }
    $sur = $stats.SurprizOran; $fav = $stats.FavoriKazanma
    $hy = $stats.EnYuksekHpKazanma; $dy = $stats.EnDusukHpKazanma
    $seviye = 'orta'; $baslik = 'Dengeli'
    if ($null -ne $sur -and $sur -ge 58) { $seviye = 'yuksek'; $baslik = 'Sürprize açık' }
    elseif ($null -ne $fav -and $fav -ge 55) { $seviye = 'dusuk'; $baslik = 'Favori eğilimli' }

    $p = @()
    if ($null -ne $fav -or $null -ne $sur) { $p += ("Son {0} koşuda favori kazanma %{1}, sürpriz %{2}." -f $stats.KosuSayisi, $fav, $sur) }
    if ($null -ne $hy) { $p += ("En yüksek HP kazandı %{0}; en düşük HP %{1}." -f $hy, $dy) }
    if ($seviye -eq 'yuksek') { $p += 'Bu tip sürprize açık: favoriye tek başına güvenme; yüksek HP + pist uyumu + formu iyi atları öne çıkar.' }
    elseif ($seviye -eq 'dusuk') { $p += 'Favori eğilimli: favoriyi HP ve form ile teyit etmek makul.' }
    elseif ($seviye -eq 'orta') { $p += 'Denge: HP sırası ve form sinyallerini birlikte değerlendir.' }

    # Sürpriz adayı: konsensüs lideri ama az oynanan (AGF sırası yüksek)
    $aday = $null
    if ($consensus -and $race.Atlar) {
        $top = @($consensus)[0]
        if ($top) {
            $h = $race.Atlar | Where-Object { $_.At -eq $top.At } | Select-Object -First 1
            if ($h -and $h.AgfRank -and [int]$h.AgfRank -ge 4) { $aday = $top.At }
        }
    }
    if ($aday) { $p += ("Sürpriz adayı: {0} (az oynanan ama sinyalleri güçlü)." -f $aday) }
    if ([int]$stats.KosuSayisi -lt 20) { $p += ("Örnek küçük (N={0}); ihtiyatlı yorumla." -f $stats.KosuSayisi) }

    return [pscustomobject]@{ Seviye = $seviye; Baslik = $baslik; Metin = ($p -join ' '); SurprizAdayi = $aday }
}

function Clean-Name([string]$s) {
    if (-not $s) { return '' }
    $t = $s.ToUpperInvariant()
    $t = $t -replace 'İ','I' -replace 'Ş','S' -replace 'Ğ','G' -replace 'Ü','U' -replace 'Ö','O' -replace 'Ç','C'
    $t = $t -replace '[^A-Z0-9]+',' '
    return $t.Trim()
}
function Find-OptionId($options, [string]$text) {
    $want = Clean-Name $text
    if (-not $want) { return $null }
    foreach ($o in $options) {
        $oAd = Clean-Name $o.Ad
        if ($oAd -eq $want) { return $o.Id }
    }
    # en uzun eslesen secenegi sec: "HANDIKAP 16 H2" -> "HANDIKAP 16" (duz "HANDIKAP" degil)
    $best = $null; $bestLen = 0
    foreach ($o in $options) {
        $oAd = Clean-Name $o.Ad
        if (-not $oAd) { continue }
        if ($want -like "*$oAd*" -or $oAd -like "*$want*") {
            if ($oAd.Length -gt $bestLen) { $best = $o.Id; $bestLen = $oAd.Length }
        }
    }
    if ($best) { return $best }
    # ilk kelimeler
    $w1 = ($want -split ' ')[0]
    foreach ($o in $options) { if ((Clean-Name $o.Ad) -like "$w1*") { return $o.Id } }
    return $null
}

# KosuSorgulama veri tablosu: kosu basina kazanan satirlari
function Get-QueryRows([string]$html) {
    $list = @()
    if (-not $html) { return $list }
    $parts = [regex]::Split($html, '(?i)<tr')
    foreach ($rh in $parts) {
        if ($rh -notmatch 'sorgu-KosuSorgulama-KosuSirasi') { continue }
        $cell = { param($cls)
            $m = [regex]::Match($rh, '(?s)<td class="sorgu-KosuSorgulama-' + $cls + '"[^>]*>(.*?)</td>')
            if ($m.Success) { return (Strip-Tags $m.Groups[1].Value) } else { return '' }
        }
        $km = [regex]::Match($rh, '#(\d{4,})')
        $list += [pscustomobject]@{
            Tarih    = (& $cell 'Tarih')
            Sehir    = (& $cell 'Sehir')
            Kosu     = (& $cell 'KosuSirasi')
            Grup     = (& $cell 'KosuGrubuAdi')
            Cins     = (& $cell 'KosuCinsiAdi')
            Mesafe   = (& $cell 'Mesafe')
            Pist     = (& $cell 'PistAdi')
            Kilo     = (& $cell 'Kilo')
            At       = (& $cell 'BirinciAtAdi')
            Yas      = (& $cell 'BirinciAtAdiYas')
            Hp       = (& $cell 'HandikapPuani')
            Derece   = (& $cell 'BirinciAtDerece')
            KosuKodu = $(if ($km.Success) { $km.Groups[1].Value } else { '' })
        }
    }
    return $list
}

# Gunluk sehir sonuclari: kosu basina tam kadro (bitis sirasi, HP, kilo, AGF)
function Get-ResultBlocks([string]$html) {
    $out = @()
    if (-not $html) { return $out }
    $ms = [regex]::Matches($html, '(?s)<div id="kosubilgisi-(\d+)"[^>]*>(.*?)(?=<div id="kosubilgisi-|\z)')
    foreach ($m in $ms) {
        $kodu = $m.Groups[1].Value
        $block = $m.Groups[2].Value
        $atlar = @()
        $rows = [regex]::Split($block, '(?i)<tr')
        foreach ($rh in $rows) {
            if ($rh -notmatch 'GunlukYarisSonuclari-SONUCNO') { continue }
            $cell = { param($cls)
                $x = [regex]::Match($rh, '(?s)<td class="gunluk-GunlukYarisSonuclari-' + $cls + '"[^>]*>(.*?)</td>')
                if ($x.Success) { return $x.Groups[1].Value } else { return '' }
            }
            $atHtml = & $cell 'AtAdi3'
            $am = [regex]::Match($atHtml, 'AtId=\d+[^>]*>\s*([^<]+)')
            $at = if ($am.Success) { Html-Decode $am.Groups[1].Value.Trim() } else { Strip-Tags $atHtml }
            $at = ($at -replace '\s*\(\d+\)\s*$','').Trim()
            $agfHtml = & $cell 'AGFORAN'
            $agf = $null; $agfRank = $null
            $tm = [regex]::Match($agfHtml, 'title="%([0-9.,]+)\((\d+)\)"')
            if ($tm.Success) { $agf = Parse-Num $tm.Groups[1].Value; $agfRank = [int]$tm.Groups[2].Value }
            if (-not $agfRank) { $vm = [regex]::Match($agfHtml, '\((\d+)\)'); if ($vm.Success) { $agfRank = [int]$vm.Groups[1].Value } }
            $atlar += [pscustomobject]@{
                Sira    = (Parse-Num (Strip-Tags (& $cell 'SONUCNO')))
                At      = $at
                Yas     = (Strip-Tags (& $cell 'Yas'))
                Kilo    = (First-Num (Strip-Tags (& $cell 'Kilo')))
                Hp      = (Parse-Num (Strip-Tags (& $cell 'Hc')))
                Gny     = (Parse-Num (Strip-Tags (& $cell 'Gny')))
                Agf     = $agf
                AgfRank = $agfRank
            }
        }
        if ($atlar.Count -gt 0) { $out += [pscustomobject]@{ KosuKodu = $kodu; Atlar = $atlar } }
    }
    return $out
}

function Build-RaceTypeStats($race, $queryOptions, [string]$sehirId = '-1') {
    if (-not $queryOptions) { return $null }
    $tipTemiz = (([string]$race.Tip) -split '/')[0].Trim()
    $cinsId = Find-OptionId $queryOptions.KosuCinsi $tipTemiz
    $grupId = Find-OptionId $queryOptions.Grup $race.Grup
    if (-not $cinsId) { return [pscustomobject]@{ Hata="Koşu tipi eşlenemedi: $($race.Tip)"; Tip=$race.Tip; Grup=$race.Grup } }

    # ırk ve cinsiyet (koşunun kimliğine göre)
    $irkId = '-1'
    if ($queryOptions.PSObject.Properties.Name -contains 'Irk' -and $queryOptions.Irk) {
        $x = Find-OptionId $queryOptions.Irk ($race.Grup + ' ' + $race.Tip); if ($x) { $irkId = $x }
    }
    $cinsiyetId = '-1'
    if ($race.PSObject.Properties.Name -contains 'CinsiyetId' -and [string]$race.CinsiyetId) { $cinsiyetId = [string]$race.CinsiyetId }
    elseif ($queryOptions.PSObject.Properties.Name -contains 'Cinsiyet' -and $queryOptions.Cinsiyet) {
        if ($race.Tip -match '(?i)Dişi') { $x = Find-OptionId $queryOptions.Cinsiyet 'Dişi'; if ($x) { $cinsiyetId = $x } }
        elseif ($race.Tip -match '(?i)Erkek') { $x = Find-OptionId $queryOptions.Cinsiyet 'Erkek'; if ($x) { $cinsiyetId = $x } }
    }
    # istatistik panelinden cinsiyet secildiyse onu kullan (Tumu/Disi/Erkek)
    if ($race.PSObject.Properties.Name -contains 'CinsiyetId' -and [string]$race.CinsiyetId) { $cinsiyetId = [string]$race.CinsiyetId }

    $end = [datetime]::Now
    $start = $end.AddYears(-[math]::Abs($StatsYil))
    if ($StatsKosu -gt 200) { $start = $end.AddYears(-10) }
    if (-not $grupId) { $grupId = '-1' }
    if (-not $irkId) { $irkId = '-1' }
    if (-not $cinsiyetId) { $cinsiyetId = '-1' }
    $dateQs = "QueryParameter_Tarih_Start=" + $start.ToString('dd/MM/yyyy', $script:CIV) + "&QueryParameter_Tarih_End=" + $end.ToString('dd/MM/yyyy', $script:CIV)
    if (-not $sehirId) { $sehirId = '-1' }
    if ($sehirId -notmatch '^-?\d+$') { $sehirId = '-1' }
    # Kategori (yaş + ırk) birebir sorgulanır; veri yoksa geniş sorguya DÜŞÜLMEZ.
    # Aksi halde "Handikap 15" gibi tiplerde 3 yaşlı / 4+ yaşlı koşular karışır.
    $set = "QueryParameter_SehirId=$sehirId&QueryParameter_IrkId=$irkId&QueryParameter_KosuCinsiId=$cinsId&QueryParameter_GrupId=$grupId&QueryParameter_Cinsiyet=$cinsiyetId"
    $u = "$script:Base/TR/YarisSever/Query/Data/KosuSorgulama?$dateQs&$set"
    $html = Invoke-Tjk $u -TimeoutSec 25 -Cat ('sorgu-' + $cinsId + '-' + $grupId + '-' + $irkId + '-' + $cinsiyetId)
    $rows = @()
    if ($html) { $rows = @(ConvertFrom-HtmlTable $html) }
    if (-not $html) { return [pscustomobject]@{ Hata="Veri alınamadı (TJK yoğun)"; Tip=$race.Tip; Grup=$race.Grup } }
    if ($rows.Count -eq 0) { return [pscustomobject]@{ Hata="Sonuç bulunamadı"; Tip=$race.Tip; Grup=$race.Grup; CinsId=$cinsId; GrupId=$grupId } }

    # kosu satirlari (kazananlar) - tarihe gore yeniden eskiye
    $qrows = @(Get-QueryRows $html)
    if ($qrows.Count -eq 0) { return [pscustomobject]@{ Hata='Sonuçlar ayrıştırılamadı'; Tip=$race.Tip; Grup=$race.Grup; CinsId=$cinsId } }
    $qrows = @($qrows | Sort-Object {
        $dd=[datetime]::MinValue; $tt=[string]$_.Tarih
        if (-not [datetime]::TryParseExact($tt,'dd.MM.yyyy',$script:CIV,'None',[ref]$dd)) {
            [void][datetime]::TryParseExact($tt,'dd/MM/yyyy',$script:CIV,'None',[ref]$dd)
        }
        if ($dd -eq [datetime]::MinValue) { [datetime]::MinValue } else { $dd }
    } -Descending)
    # mesafe bandi secildiyse once suz, sonra ornek sayisini kes
    if ($race.PSObject.Properties.Name -contains 'MesafeBand' -and $race.MesafeBand) {
        $qrows = @($qrows | Where-Object { (Dist-Band ([int](Parse-Num $_.Mesafe))) -eq [string]$race.MesafeBand })
    }
    if ($StatsKosu -gt 0) { $qrows = @($qrows | Select-Object -First $StatsKosu) }

    # tam kadro: gun+sehir sonuc sayfalarindan (1 sayfa = o gunun tum kosulari)
    # Ozel suzme (mesafe/cinsiyet) isteklerinde bu agir adim atlanir; ozet yine cikar.
    $skipField = ($race.PSObject.Properties.Name -contains 'HizliField' -and $race.HizliField)
    $need = @{}; $dayCache = @{}; $resCache = @{}; $fieldMap = @{}
    if (-not $skipField) {
    foreach ($r in $qrows) {
        if (-not $r.KosuKodu) { continue }
        $k = $r.Tarih + '|' + $r.Sehir
        if (-not $need.ContainsKey($k)) { $need[$k] = @() }
        $need[$k] += $r
    }
    foreach ($k in @($need.Keys)) {
        $pp = $k -split '\|', 2
        $tarih = $pp[0]; $sehir = $pp[1]
        $esc = [uri]::EscapeDataString(($tarih -replace '\.','/'))
        if (-not $dayCache.ContainsKey($tarih)) {
            $dh = Invoke-Tjk ("$script:Base/TR/YarisSever/Info/Page/GunlukYarisSonuclari?QueryParameter_Tarih=$esc") -TimeoutSec 25 -Cat ('sonuc-' + $tarih)
            $cs = @()
            if ($dh) { $cs = @(Get-ResultCities $dh) }
            $dayCache[$tarih] = $cs
        }
        $city = @($dayCache[$tarih] | Where-Object { $_.Ad -eq $sehir }) | Select-Object -First 1
        if (-not $city) { continue }
        $rk = $tarih + '|' + $city.SehirId
        if (-not $resCache.ContainsKey($rk)) {
            $ru = "$script:Base/TR/YarisSever/Info/Sehir/GunlukYarisSonuclari?SehirId=$($city.SehirId)&QueryParameter_Tarih=$esc&SehirAdi=" + [uri]::EscapeDataString($sehir)
            $rh = Invoke-Tjk $ru -TimeoutSec 30 -Cat ('sonuc-' + $rk)
            $map = @{}
            if ($rh) { foreach ($b in @(Get-ResultBlocks $rh)) { $map[$b.KosuKodu] = $b.Atlar } }
            $resCache[$rk] = $map
        }
        foreach ($r in $need[$k]) {
            if ($r.KosuKodu -and $resCache[$rk].ContainsKey($r.KosuKodu)) { $fieldMap[$r.KosuKodu] = $resCache[$rk][$r.KosuKodu] }
        }
    }
    }

    $raceList = @()
    $favoriOlan = 0; $enYuksekHpOlan = 0; $enDusukHpOlan = 0; $kosuSayisi = 0; $kadroVar = 0
    $hps = @(); $sik = @(); $yaslar = @(); $fieldHps = @(); $fieldSik = @()
    foreach ($r in $qrows) {
        $kosuSayisi++
        $field = @()
        if ($r.KosuKodu -and $fieldMap.ContainsKey($r.KosuKodu)) { $field = @($fieldMap[$r.KosuKodu]) }
        if ($field.Count -gt 0) {
            $kadroVar++
            $fieldHps += @($field | ForEach-Object { Parse-Num $_.Hp } | Where-Object { $null -ne $_ })
            $fieldSik += @($field | ForEach-Object { Parse-Num $_.Kilo } | Where-Object { $null -ne $_ })
        }
        $isFav = $false; $hpRank = $null; $alan = $field.Count; $wAgf = $null; $wAgfRank = $null
        $isMax = $false; $isMin = $false
        $irkAd = ''
        if ([string]$r.Grup -match '(?i)arap') { $irkAd = 'Arap' }
        elseif ([string]$r.Grup -match '(?i)ingiliz') { $irkAd = 'İngiliz' }
        $wHP = Parse-Num $r.Hp
        $wKilo = First-Num $r.Kilo
        $wYas = [string]$r.Yas
        if ($field.Count -gt 0) {
            $fav = @($field | Where-Object { $_.AgfRank -eq 1 }) | Select-Object -First 1
            if (-not $fav) { $fav = @($field | Where-Object { $null -ne $_.Agf } | Sort-Object Agf | Select-Object -First 1) }
            $winH = @($field | Where-Object { $_.Sira -eq 1 }) | Select-Object -First 1
            if ($winH) {
                $isFav = [bool]($fav -and $winH.At -eq $fav.At)
                if ($null -ne $winH.Hp) { $wHP = $winH.Hp }
                if ($null -ne $winH.Kilo) { $wKilo = $winH.Kilo }
                if ($winH.Yas) { $wYas = $winH.Yas }
                $wAgf = $winH.Agf; $wAgfRank = $winH.AgfRank
            }
            $vals = @($field | Where-Object { $null -ne $_.Hp } | ForEach-Object { $_.Hp })
            if ($vals.Count -gt 0 -and $null -ne $wHP) {
                $mx = ($vals | Measure-Object -Maximum).Maximum
                $mn = ($vals | Measure-Object -Minimum).Minimum
                if ($wHP -eq $mx) { $enYuksekHpOlan++; $isMax = $true }
                if ($wHP -eq $mn) { $enDusukHpOlan++; $isMin = $true }
                $hpRank = 1 + @($vals | Where-Object { $_ -gt $wHP }).Count
            }
        }
        if ($isFav) { $favoriOlan++ }
        if ($null -ne $wHP) { $hps += $wHP }
        if ($null -ne $wKilo) { $sik += $wKilo }
        if ($wYas -match '(\d{1,2})') { $yaslar += [int]$Matches[1] }
        $raceList += [pscustomobject]@{
            Tarih  = [string]$r.Tarih
            Sehir  = [string]$r.Sehir
            KosuNo = [string]$r.Kosu
            Mesafe = [int](Parse-Num $r.Mesafe)
            Pist   = [string]$r.Pist
            At     = [string]$r.At
            Yas    = $wYas
            Kilo   = $wKilo
            Gny    = ''
            Agf    = $wAgf
            AgfRank = $wAgfRank
            HP     = $wHP
            HpSira = $hpRank
            Favori = $isFav
            EnYuksekHP = $isMax
            EnDusukHP  = $isMin
            Irk    = $irkAd
            Alan   = $alan
        }
    }

    # HP bantları
    $bands = @()
    if ($hps.Count -gt 0) {
        $min = ($hps | Measure-Object -Minimum).Minimum; $max = ($hps | Measure-Object -Maximum).Maximum
        $step = [math]::Max(5, [math]::Ceiling(($max - $min) / 6.0))
        $b = [math]::Floor($min / $step) * $step
        while ($b -le $max) {
            $lo = $b; $hi = $b + $step - 1
            $cnt = @($hps | Where-Object { $_ -ge $lo -and $_ -le $hi }).Count
            $bands += [pscustomobject]@{ Bant="$lo-$hi"; Kazanan=$cnt; Oran=[math]::Round(100.0*$cnt/$hps.Count,1) }
            $b += $step
        }
        $bands = @($bands | Sort-Object @{Expression={ [int](($_.Bant -split '-')[0]) };Descending=$true})
    }
    # Girenlerin (tum kadro) ayni bantlardaki dagilimi: kazanan orani tek
    # basina yaniltir; karsilastirma icin taban oran sart.
    $bandsGiren = @()
    if ($bands.Count -gt 0 -and $fieldHps.Count -gt 0) {
        foreach ($b in $bands) {
            $pp = $b.Bant -split '-'
            $lo = [double]$pp[0]; $hi = [double]$pp[1]
            $c = @($fieldHps | Where-Object { $_ -ge $lo -and $_ -le $hi }).Count
            $bandsGiren += [pscustomobject]@{ Bant=$b.Bant; Giren=$c; Oran=[math]::Round(100.0*$c/$fieldHps.Count,1) }
        }
    }
    $sikletOzet = $null
    if ($sik.Count -gt 0) {
        $sikletOzet = [pscustomobject]@{
            Ortalama=[math]::Round(($sik | Measure-Object -Average).Average,1)
            Min=($sik | Measure-Object -Minimum).Minimum
            Max=($sik | Measure-Object -Maximum).Maximum
        }
    }
    $sikletGiren = $null
    if ($fieldSik.Count -gt 0) {
        $sikletGiren = [pscustomobject]@{
            Ortalama=[math]::Round(($fieldSik | Measure-Object -Average).Average,1)
            Min=($fieldSik | Measure-Object -Minimum).Minimum
            Max=($fieldSik | Measure-Object -Maximum).Maximum
        }
    }
    return [pscustomobject]@{
        Tip=$race.Tip; Grup=$race.Grup; MesafeBand=$(if ($race.PSObject.Properties.Name -contains 'MesafeBand') { [string]$race.MesafeBand } else { '' }); CinsId=$cinsId; GrupId=$grupId; IrkId=$irkId; CinsiyetId=$cinsiyetId; SehirId=$sehirId
        Yil=$StatsYil; KosuSayisi=$kosuSayisi; HedefKosu=$StatsKosu
        KadroSayisi=$kadroVar; KadroOran= if ($kosuSayisi -gt 0) { [math]::Round(100.0*$kadroVar/$kosuSayisi,0) } else { 0 }
        HpBantlari=$bands
        HpBantlariGiren=$bandsGiren
        Siklet=$sikletOzet
        SikletGiren=$sikletGiren
        FavoriKazanma= if ($kosuSayisi -gt 0) { [math]::Round(100.0*$favoriOlan/$kosuSayisi,1) } else { $null }
        SurprizOran= if ($kosuSayisi -gt 0) { [math]::Round(100.0*($kosuSayisi-$favoriOlan)/$kosuSayisi,1) } else { $null }
        EnYuksekHpKazanma= if ($kosuSayisi -gt 0) { [math]::Round(100.0*$enYuksekHpOlan/$kosuSayisi,1) } else { $null }
        EnDusukHpKazanma= if ($kosuSayisi -gt 0) { [math]::Round(100.0*$enDusukHpOlan/$kosuSayisi,1) } else { $null }
        OrtYas= if ($yaslar.Count -gt 0) { [math]::Round(($yaslar | Measure-Object -Average).Average,1) } else { $null }
        Kosular=$raceList
        Hata=$null
    }
}

# ---------- ana akış ----------
function Build-DayData([string]$TarihStr, $queryOptions, $winsCache, $histCache, $statsCache) {
    $isPast = $false
    try { $isPast = ([datetime]::ParseExact($TarihStr,'dd/MM/yyyy',$script:CIV).Date -lt (Get-Date).Date) } catch {}
    $esc = [uri]::EscapeDataString($TarihStr)
    $urls = @()
    $urls += "$script:Base/TR/YarisSever/Info/Page/GunlukYarisProgrami?QueryParameter_Tarih=$esc"
    if ($isPast) {
        $urls += "$script:Base/TR/YarisSever/Info/Data/GunlukYarisProgrami?QueryParameter_Tarih=$esc&Era=past"
        $urls += "$script:Base/TR/YarisSever/Info/Page/GunlukYarisProgrami?QueryParameter_Tarih=$esc&Era=today"
    }
    $progHtml = $null; $cities = @()
    foreach ($u in $urls) {
        $tmp = Invoke-Tjk $u
        if (-not $tmp) { continue }
        if (-not $progHtml) { $progHtml = $tmp }
        $cs = @(Get-Cities $tmp)
        if ($cs.Count -gt 0) { $progHtml = $tmp; $cities = $cs; break }
    }
    if (-not $progHtml) { Write-Host ("Program sayfası alınamadı: {0}" -f $TarihStr) -ForegroundColor Red; return $null }
    if ($Sehir) {
        $want = @($Sehir -split ',' | ForEach-Object { $_.Trim() })
        $cities = @($cities | Where-Object { $n=$_.Ad; ($want | Where-Object { $n -like "*$_*" }) })
    }
    if (-not $Yabanci) { $cities = @($cities | Where-Object { -not $_.Yabanci -and $_.Ad -notmatch '^Karma' -and $_.Ad -notmatch '\(YD' }) }
    Write-Host ("Şehir sayısı: {0}" -f $cities.Count)

    $sireStore = Get-SireStore
    $cityOut = @()
    foreach ($city in $cities) {
        Write-Host ("-> {0}" -f $city.Ad) -ForegroundColor Green
        $cityUrl = $city.Url -replace '[&?]Era=[^&]*', ''
        $html = Invoke-Tjk $cityUrl -TimeoutSec 15
        # onbellekte Era'li URL ile duran sayfalar icin geri donus
        if (-not $html -and $city.Url -ne $cityUrl) { $html = Invoke-Tjk $city.Url -TimeoutSec 15 }
        if (-not $html) { continue }
        $races = Get-Races $html $city.SehirId $city.Ad
        $winsMap = @{}; $cityJockey = @()
        if (-not $Hizli) {
            $winsMap = Get-JockeyWinsMap $races $city.SehirId $Donem $winsCache
            $cityJockey = City-JockeyRank $races $winsMap
        }
        $raceOut = @(); $ri = 0
        foreach ($race in $races) {
            $ri++
            if ($MaxRace -gt 0 -and $ri -gt $MaxRace) { break }
            Write-Host ("   {0}. Koşu {1} [{2}] {3} at" -f $race.KosuNo, $race.Saat, $race.Tip, $race.Atlar.Count)
            $jockeyRank = @()
            if (-not $Hizli) { $jockeyRank = Rank-FromMap $race.Atlar $winsMap }
            $surf = @{}; $h2h = $null; $combo = @(); $form = @(); $dist = @{}; $sireHints = @{}
            if (-not $Hizli) {
                $surf = Build-Surfaces $race $histCache
                $h2h = Build-HeadToHead $race $histCache
                $combo = Build-JockeyHorseCombo $race $histCache
                $dist = Build-Distance $race $histCache
                Update-SireStore $race $histCache $sireStore
                $sireHints = Build-SireHints $race $histCache $sireStore
            }
            $form = Build-FormScore $race
            $consensus = Build-Consensus $race $surf $h2h $jockeyRank $form
            $progYorum = Build-ProgramVerdict $race $form
            $genis = $null
            if (-not $Hizli) { $genis = Build-ExtendedVerdict $race $form $jockeyRank $surf $combo }

            # koşu tipi istatistiği (önbellekli)
            $statKey = Clean-Name ($race.Tip + '|' + $race.Grup)
            $stats = $null
            if ($queryOptions -and -not $Hizli) {
                if ($statsCache.ContainsKey($statKey)) { $stats = $statsCache[$statKey]; if ($stats -and $stats.Tip) { $stats = $stats.PSObject.Copy(); $stats | Add-Member -NotePropertyName Tip -NotePropertyValue $race.Tip -Force } }
                else {
                    if (Get-Command Build-StatsData -ErrorAction SilentlyContinue) { $stats = Build-StatsData ([string]$race.Tip) ([string]$race.Grup) 50 }
                    else { $stats = Build-RaceTypeStats $race $queryOptions }
                    $statsCache[$statKey] = $stats
                }
            }
            $yorum = Build-Verdict $race $stats $consensus

            $raceOut += [pscustomobject]@{
                KosuKodu=$race.KosuKodu; KosuNo=$race.KosuNo; Saat=$race.Saat; Sehir=$race.Sehir; SehirId=$race.SehirId; Tip=$race.Tip; Grup=$race.Grup; Cinsiyet=$race.Cinsiyet; CinsiyetId=$race.CinsiyetId
                Mesafe=$race.Mesafe; Pist=$race.Pist
                Atlar=$race.Atlar
                JokeySira=$jockeyRank
                Yuzey=$surf
                MesafeKayit=$dist
                BabaKarnesi=$sireHints
                IkiliGecme=$h2h
                JokeyAt=$combo
                Form=$form
                Konsensus=$consensus
                Istatistik=$stats
                Yorum=$yorum
                ProgramYorum=$progYorum
                GenisYorum=$genis
            }
        }
        $cityOut += [pscustomobject]@{ SehirId=$city.SehirId; Ad=$city.Ad; Yabanci=$city.Yabanci; JokeySira=$cityJockey; Kosular=$raceOut }
    }

    if (-not $Hizli) { Save-SireStore $sireStore }
    return [pscustomobject]@{ Tarih = $TarihStr; Hava = (Get-RaceWeather $progHtml); Sehirler = $cityOut }
}

function Invoke-Main {
    Write-Host "=== At Yarışı Analiz — TJK ===" -ForegroundColor Cyan
    Write-Host "Tarih: $Tarih"

    # koşu sorgu seçenekleri (istatistik için)
    $queryOptions = $null
    if (-not $Hizli) {
        $qHtml = Invoke-Tjk "$script:Base/TR/YarisSever/Query/Page/KosuSorgulama"
        if ($qHtml) {
            $queryOptions = [pscustomobject]@{
                KosuCinsi = Get-SelectOptions $qHtml 'QueryParameter_KosuCinsiId'
                Grup      = Get-SelectOptions $qHtml 'QueryParameter_GrupId'
                Irk       = Get-SelectOptions $qHtml 'QueryParameter_IrkId'
                Cinsiyet  = Get-SelectOptions $qHtml 'QueryParameter_Cinsiyet'
            }
        }
    }

    $winsCache = @{}; $histCache = @{}; $statsCache = @{}
    $dates = @()
    try {
        $base = [datetime]::ParseExact($Tarih, 'dd/MM/yyyy', $script:CIV)
        if ($Aralik -gt 0) { for ($k = -$Aralik; $k -le $Aralik; $k++) { $dates += $base.AddDays($k).ToString('dd/MM/yyyy', $script:CIV) } }
        else { $dates = @($Tarih) }
    } catch { $dates = @($Tarih) }

    $gunler = @()
    foreach ($d in $dates) {
        Write-Host ("=== {0} ===" -f $d) -ForegroundColor Cyan
        $g = Build-DayData $d $queryOptions $winsCache $histCache $statsCache
        if ($g) { $gunler += $g }
    }
    if ($gunler.Count -eq 0) { Write-Host "Hicbir tarih icin veri alinamadi." -ForegroundColor Red; return }

    $payload = [pscustomobject]@{
        Uretim = (Get-Date).ToString('dd.MM.yyyy HH:mm')
        Donem = $Donem
        Tarihler = $gunler
        Tarih = $gunler[0].Tarih
        Hava = $gunler[0].Hava
        Sehirler = $gunler[0].Sehirler
    }
    $json = ($payload | ConvertTo-Json -Depth 12 -Compress) -replace '</','<\/'

    $tplPath = Join-Path $script:Root 'sablon.html'
    if (-not (Test-Path $tplPath)) { Write-Host "sablon.html bulunamadi!" -ForegroundColor Red; return }
    $tpl = [System.IO.File]::ReadAllText($tplPath, [System.Text.Encoding]::UTF8)
    $htmlOut = $tpl.Replace('__DATA__', $json).Replace('__TARIH__', $Tarih)

    $outName = 'gunluk-' + ((Get-Date).ToString('yyyyMMdd')) + '.html'
    $outPath = Join-Path $script:RaporDir $outName
    [System.IO.File]::WriteAllText($outPath, $htmlOut, (New-Object System.Text.UTF8Encoding($false)))
    Write-Host ("Rapor: {0}" -f $outPath) -ForegroundColor Cyan
    if (-not $NoOpen) { Start-Process $outPath }
}

if ($MyInvocation.InvocationName -ne '.') { Invoke-Main }
