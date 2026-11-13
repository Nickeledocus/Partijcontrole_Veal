<#

    Partijcontrole Veal
    -------------------
    Controleert 1 UBN (Uniek Bedrijfsnummer) met 1 tot 5 AM91-nummers bij SKV:
      1. het SKV-kwaliteitscertificaat (SKV API, met bearer-token);
      2. de VKI (Voedselketeninformatie) via de SKV InfoKalf-webservice (SOAP).

    Werkt met Windows PowerShell 5.1 (standaard aanwezig op Windows 10/11).
    Start via "Start Partijcontrole Veal.bat" (dubbelklikken).

    Instellingen (adressen, gebruikersnamen, wachtwoorden) staan in appsettings.json
    in dezelfde map als dit script.

    Testmodus (alleen de functies laden, geen scherm):
        $env:PARTIJCONTROLE_TEST = '1'; . .\Partijcontrole-Veal.ps1
    (de oude naam $env:SKV_TEST = '1' werkt ook nog)
#>

$ErrorActionPreference = 'Stop'
$script:ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$script:ConfigPath = Join-Path $script:ScriptDir 'appsettings.json'
$script:AppNaam = 'Partijcontrole Veal'

# ---------------------------------------------------------------------------
#  Algemene instellingen voor internetverkeer
# ---------------------------------------------------------------------------
Add-Type -AssemblyName System.Net.Http
try {
    [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
} catch { }
try {
    # Zorgt dat een eventuele bedrijfsproxy je Windows-login gebruikt
    if ([System.Net.WebRequest]::DefaultWebProxy) {
        [System.Net.WebRequest]::DefaultWebProxy.Credentials = [System.Net.CredentialCache]::DefaultCredentials
    }
} catch { }

# Wordt tijdens wachten aangeroepen zodat het scherm niet "bevriest"
$script:UiPump = $null

$script:Token = $null
$script:TokenExpires = [datetime]::MinValue
$script:HttpClient = $null
$script:VkiHttpClient = $null
$script:Config = $null
$script:UserAgent = 'Partijcontrole-Veal/2.0'

# VKI-instellingen (worden gevuld door Read-Config)
$script:VkiStandaardUrl = 'https://infokalf.skv.info/service/VKIService/V2/VKIwebservice.svc'
$script:VkiSoapAction = '"http://tempuri.org/IVKIwebservice/GetVKIdieren"'
$script:Vki = [pscustomobject]@{ Actief = $false; Reden = 'uit'; Url = ''; Username = ''; Password = ''; TimeoutSeconden = 60; Waarschuwing = '' }
# Wordt $true na ERR0002 (inlog geweigerd). Daarna geen VKI-aanroepen meer tot het programma opnieuw start,
# want SKV blokkeert het account 10 minuten na 3 foute pogingen.
$script:VkiInlogGeweigerd = $false

# ---------------------------------------------------------------------------
#  Hulpfuncties
# ---------------------------------------------------------------------------
function Get-ErrorText($err) {
    $ex = $err
    if ($err -is [System.Management.Automation.ErrorRecord]) { $ex = $err.Exception }
    $msgs = New-Object System.Collections.Generic.List[string]
    while ($ex) {
        if ($ex -isnot [System.Management.Automation.MethodInvocationException] -and
            $ex -isnot [System.AggregateException]) {
            if (-not $msgs.Contains($ex.Message)) { $msgs.Add($ex.Message) }
        }
        $ex = $ex.InnerException
    }
    if ($msgs.Count -eq 0) { return [string]$err }
    return ($msgs -join ' -> ')
}

# Kijkt of een fout (of een van de onderliggende fouten) een time-out is
function Test-IsTimeout($err) {
    $ex = $err
    if ($err -is [System.Management.Automation.ErrorRecord]) { $ex = $err.Exception }
    while ($ex) {
        if ($ex -is [System.Threading.Tasks.TaskCanceledException] -or $ex -is [System.TimeoutException]) { return $true }
        $ex = $ex.InnerException
    }
    return $false
}

function Wait-Task($Task) {
    while (-not $Task.IsCompleted) {
        if ($script:UiPump) { & $script:UiPump }
        Start-Sleep -Milliseconds 30
    }
    return $Task.GetAwaiter().GetResult()
}

# Eerste gevulde waarde van een van de opgegeven velden (voor oude/nieuwe schrijfwijzen)
function Get-EersteWaarde($Obj, [string[]]$Velden) {
    foreach ($v in $Velden) {
        if ($Obj.PSObject.Properties[$v] -and $null -ne $Obj.$v -and -not [string]::IsNullOrWhiteSpace([string]$Obj.$v)) {
            return [string]$Obj.$v
        }
    }
    return ''
}

function ConvertTo-JaNee($Waarde, [bool]$Standaard) {
    if ($null -eq $Waarde) { return $Standaard }
    if ($Waarde -is [bool]) { return $Waarde }
    $t = ([string]$Waarde).Trim().ToLowerInvariant()
    if ($t -in @('false', 'nee', 'no', '0', 'uit', 'off')) { return $false }
    if ($t -in @('true', 'ja', 'yes', '1', 'aan', 'on')) { return $true }
    return $Standaard
}

# Leest de sectie "Vki" uit de configuratie. Gooit GEEN fout: bij problemen staat VKI uit
# (met een reden), zodat de certificaatcontrole gewoon blijft werken.
function Get-VkiInstellingen($Cfg) {
    $vki = [pscustomobject]@{ Actief = $false; Reden = 'uit'; Url = ''; Username = ''; Password = ''; TimeoutSeconden = 60; Waarschuwing = '' }
    if (-not $Cfg -or -not $Cfg.PSObject.Properties['Vki'] -or -not $Cfg.Vki) {
        $vki.Reden = 'uit (geen Vki-sectie in appsettings.json)'
        return $vki
    }
    $s = $Cfg.Vki
    $aan = $true
    if ($s.PSObject.Properties['Ingeschakeld']) { $aan = ConvertTo-JaNee $s.Ingeschakeld $true }
    if (-not $aan) {
        $vki.Reden = 'uit (Ingeschakeld = false)'
        return $vki
    }

    # Nieuwe sleutels, met de sleutels van Test-VkiCredentials.ps1 als terugval
    $vki.Url = Get-EersteWaarde $s @('Url', 'UrlProductie')
    if (-not $vki.Url) { $vki.Url = $script:VkiStandaardUrl }
    $vki.Username = Get-EersteWaarde $s @('Username', 'Gebruikersnaam')
    $vki.Password = Get-EersteWaarde $s @('Password', 'Wachtwoord')
    $t = 0
    if ([int]::TryParse((Get-EersteWaarde $s @('TimeoutSeconden', 'TimeoutSeconds')), [ref]$t) -and $t -gt 0) { $vki.TimeoutSeconden = $t }

    $voorbeeldWaarden = @('Gebruikersnaam', 'Wachtwoord', 'Username', 'Password')
    $leeg = { param($w) ([string]::IsNullOrWhiteSpace($w) -or $w -match '^VUL-HIER' -or $voorbeeldWaarden -contains $w) }
    if ((& $leeg $vki.Username) -or (& $leeg $vki.Password)) {
        $vki.Reden = 'uit (VKI-inloggegevens niet ingevuld)'
        $vki.Waarschuwing = "De VKI-controle staat uit: vul in appsettings.json bij `"Vki`" je VKI-gebruikersnaam en -wachtwoord in.`n$($script:ConfigPath)`n`nDe certificaatcontrole werkt gewoon."
        return $vki
    }
    if ($vki.Url -notmatch '^https://') {
        $vki.Reden = 'uit (Vki.Url is geen https-adres)'
        $vki.Waarschuwing = "De VKI-controle staat uit: Vki.Url in appsettings.json moet beginnen met https://"
        return $vki
    }
    $vki.Actief = $true
    $vki.Reden = ''
    return $vki
}

function Read-Config {
    if (-not (Test-Path $script:ConfigPath)) {
        $voorbeeld = Join-Path $script:ScriptDir 'appsettings.example.json'
        if (Test-Path $voorbeeld) {
            Copy-Item $voorbeeld $script:ConfigPath
            throw "appsettings.json is zojuist aangemaakt (kopie van appsettings.example.json).`nVul daarin de URL's, je gebruikersnamen en je wachtwoorden in:`n$($script:ConfigPath)`n`n(Rechtermuisknop > Openen met > Kladblok)"
        }
        throw "Configuratiebestand niet gevonden: $($script:ConfigPath)"
    }
    $raw = [IO.File]::ReadAllText($script:ConfigPath, [Text.Encoding]::UTF8)
    try { $cfg = $raw | ConvertFrom-Json }
    catch { throw "appsettings.json is geen geldige JSON. Controleer komma's en aanhalingstekens.`n$(Get-ErrorText $_)" }

    foreach ($veld in @('TokenUrl', 'CertificaatUrl', 'Username', 'Password')) {
        if (-not $cfg.$veld -or [string]::IsNullOrWhiteSpace([string]$cfg.$veld)) {
            throw "In appsettings.json ontbreekt de waarde voor '$veld'."
        }
    }
    $voorbeeldWaarden = @('Gebruikersnaam', 'Wachtwoord', 'VUL-HIER-JE-GEBRUIKERSNAAM-IN', 'VUL-HIER-JE-WACHTWOORD-IN')
    if ([string]$cfg.TokenUrl -notmatch '^https?://' -or [string]$cfg.CertificaatUrl -notmatch '^https?://') {
        throw "Vul in appsettings.json bij TokenUrl en CertificaatUrl een volledig adres in (beginnend met https://):`n$($script:ConfigPath)"
    }
    if ($voorbeeldWaarden -contains [string]$cfg.Username -or $voorbeeldWaarden -contains [string]$cfg.Password) {
        throw "Vul eerst je SKV-gebruikersnaam en -wachtwoord in, in appsettings.json:`n$($script:ConfigPath)`n`n(Rechtermuisknop > Openen met > Kladblok)"
    }
    if (-not $cfg.TokenMethod) { $cfg | Add-Member -NotePropertyName TokenMethod -NotePropertyValue 'AUTO' -Force }
    if (-not $cfg.TimeoutSeconds) { $cfg | Add-Member -NotePropertyName TimeoutSeconds -NotePropertyValue 30 -Force }

    $script:Vki = Get-VkiInstellingen $cfg
    return $cfg
}

function Get-HttpClient {
    if (-not $script:HttpClient) {
        $handler = New-Object System.Net.Http.HttpClientHandler
        $script:HttpClient = New-Object System.Net.Http.HttpClient($handler)
        $script:HttpClient.Timeout = [TimeSpan]::FromSeconds([int]$script:Config.TimeoutSeconds)
        [void]$script:HttpClient.DefaultRequestHeaders.TryAddWithoutValidation('User-Agent', $script:UserAgent)
    }
    return $script:HttpClient
}

function New-HttpResult($StatusCode, $Reason, $Body) {
    return [pscustomobject]@{
        StatusCode = [int]$StatusCode
        Reason     = [string]$Reason
        Body       = [string]$Body
        Ok         = ([int]$StatusCode -ge 200 -and [int]$StatusCode -lt 300)
    }
}

# Normale aanvraag via HttpClient
function Invoke-SkvHttp {
    param(
        [string]$Method,
        [string]$Url,
        [string]$BearerToken,
        [string]$FormBody
    )
    $client = Get-HttpClient
    $req = New-Object System.Net.Http.HttpRequestMessage((New-Object System.Net.Http.HttpMethod($Method)), $Url)
    [void]$req.Headers.TryAddWithoutValidation('Accept', 'application/json')
    if ($BearerToken) {
        $req.Headers.Authorization = New-Object System.Net.Http.Headers.AuthenticationHeaderValue('Bearer', $BearerToken)
    }
    if ($FormBody) {
        $req.Content = New-Object System.Net.Http.StringContent($FormBody, [Text.Encoding]::UTF8, 'application/x-www-form-urlencoded')
    }
    $resp = Wait-Task ($client.SendAsync($req))
    $body = Wait-Task ($resp.Content.ReadAsStringAsync())
    return New-HttpResult ([int]$resp.StatusCode) $resp.ReasonPhrase $body
}

# Ontcijfert "chunked" antwoorden (stukjes met een lengte ervoor)
function ConvertFrom-ChunkedBytes([byte[]]$Data) {
    $out = New-Object System.IO.MemoryStream
    $pos = 0
    while ($pos -lt $Data.Length) {
        $lineEnd = $pos
        while ($lineEnd -lt $Data.Length - 1 -and -not ($Data[$lineEnd] -eq 13 -and $Data[$lineEnd + 1] -eq 10)) { $lineEnd++ }
        $sizeText = [Text.Encoding]::ASCII.GetString($Data, $pos, $lineEnd - $pos).Split(';')[0].Trim()
        if ($sizeText -eq '') { break }
        $size = [Convert]::ToInt32($sizeText, 16)
        if ($size -eq 0) { break }
        $start = $lineEnd + 2
        $out.Write($Data, $start, [Math]::Min($size, $Data.Length - $start))
        $pos = $start + $size + 2
    }
    return $out.ToArray()
}

# Kijkt of een "handmatig" ontvangen antwoord compleet is
function Find-HeaderEnd([byte[]]$Data) {
    for ($i = 0; $i -lt $Data.Length - 3; $i++) {
        if ($Data[$i] -eq 13 -and $Data[$i + 1] -eq 10 -and $Data[$i + 2] -eq 13 -and $Data[$i + 3] -eq 10) { return $i }
    }
    return -1
}

function Test-RawResponseComplete([byte[]]$Data) {
    $sep = Find-HeaderEnd $Data
    if ($sep -lt 0) { return $false }
    $head = [Text.Encoding]::ASCII.GetString($Data, 0, $sep)
    if ($head -match '(?im)^Content-Length:\s*(\d+)') {
        return ($Data.Length -ge ($sep + 4 + [int]$Matches[1]))
    }
    if ($head -match '(?im)^Transfer-Encoding:\s*chunked') {
        $staart = [Text.Encoding]::ASCII.GetString($Data, [Math]::Max($sep + 4, $Data.Length - 7), [Math]::Min(7, $Data.Length - $sep - 4))
        return $staart.EndsWith("0`r`n`r`n")
    }
    return $false
}

# Aanvraag "met de hand" voor GET + formulierdata (dat kan HttpClient in PowerShell 5.1 niet)
function Invoke-RawHttpWithBody {
    param(
        [string]$Method,
        [string]$Url,
        [string]$FormBody
    )
    $uri = [Uri]$Url
    $port = $uri.Port
    $tcp = New-Object System.Net.Sockets.TcpClient
    try {
        Wait-Task ($tcp.ConnectAsync($uri.Host, $port)) | Out-Null
        $timeoutMs = [int]$script:Config.TimeoutSeconds * 1000
        $tcp.ReceiveTimeout = $timeoutMs
        $tcp.SendTimeout = $timeoutMs
        $stream = $tcp.GetStream()
        if ($uri.Scheme -eq 'https') {
            $ssl = New-Object System.Net.Security.SslStream($stream, $false)
            $ssl.AuthenticateAsClient($uri.Host, $null, [System.Security.Authentication.SslProtocols]::Tls12, $false)
            $stream = $ssl
        }
        $bodyBytes = [Text.Encoding]::UTF8.GetBytes($FormBody)
        $head = "$Method $($uri.PathAndQuery) HTTP/1.1`r`n" +
                "Host: $($uri.Host)`r`n" +
                "Accept: application/json`r`n" +
                "User-Agent: $($script:UserAgent)`r`n" +
                "Content-Type: application/x-www-form-urlencoded`r`n" +
                "Content-Length: $($bodyBytes.Length)`r`n" +
                "Connection: close`r`n`r`n"
        $headBytes = [Text.Encoding]::ASCII.GetBytes($head)
        $stream.Write($headBytes, 0, $headBytes.Length)
        $stream.Write($bodyBytes, 0, $bodyBytes.Length)
        $stream.Flush()

        $ms = New-Object System.IO.MemoryStream
        $buf = New-Object byte[] 65536
        while ($true) {
            try { $read = Wait-Task ($stream.ReadAsync($buf, 0, $buf.Length)) }
            catch { if ($ms.Length -gt 0) { break } else { throw } }
            if ($read -le 0) { break }
            $ms.Write($buf, 0, $read)
            if (Test-RawResponseComplete $ms.ToArray()) { break }
        }
        $all = $ms.ToArray()
    }
    finally {
        $tcp.Close()
    }

    # Kop en inhoud scheiden
    $sep = Find-HeaderEnd $all
    if ($sep -lt 0) { throw "Onvolledig antwoord ontvangen van $($uri.Host)." }
    $headerText = [Text.Encoding]::ASCII.GetString($all, 0, $sep)
    $bodyRaw = New-Object byte[] ($all.Length - $sep - 4)
    [Array]::Copy($all, $sep + 4, $bodyRaw, 0, $bodyRaw.Length)

    $lines = $headerText -split "`r`n"
    $statusParts = $lines[0].Split(' ', 3)
    $code = [int]$statusParts[1]
    $reason = ''
    if ($statusParts.Length -gt 2) { $reason = $statusParts[2] }
    if ($headerText -match '(?im)^Transfer-Encoding:\s*chunked') {
        $bodyRaw = ConvertFrom-ChunkedBytes $bodyRaw
    }
    return New-HttpResult $code $reason ([Text.Encoding]::UTF8.GetString($bodyRaw))
}

# Haalt een foutmelding uit een antwoord (JSON-tekst, error_description, Message, of ruwe tekst)
function Get-ResponseMessage([string]$Body) {
    if ([string]::IsNullOrWhiteSpace($Body)) { return '' }
    try {
        $parsed = $Body | ConvertFrom-Json
        if ($parsed -is [string]) { return $parsed }
        foreach ($veld in @('error_description', 'Message', 'message', 'error', 'title')) {
            if ($parsed.PSObject.Properties[$veld] -and $parsed.$veld) { return [string]$parsed.$veld }
        }
    } catch { }
    $t = $Body.Trim()
    if ($t.Length -gt 300) { $t = $t.Substring(0, 300) + '...' }
    return $t
}

# ---------------------------------------------------------------------------
#  Bearer-token
# ---------------------------------------------------------------------------
function Get-SkvToken([switch]$Force) {
    if (-not $Force -and $script:Token -and (Get-Date) -lt $script:TokenExpires) {
        return $script:Token
    }
    $cfg = $script:Config
    $form = 'grant_type=password&username=' + [Uri]::EscapeDataString([string]$cfg.Username) +
            '&password=' + [Uri]::EscapeDataString([string]$cfg.Password)

    $methode = ([string]$cfg.TokenMethod).ToUpperInvariant()
    $pogingen = @()
    switch ($methode) {
        'POST'  { $pogingen = @('POST') }
        'GET'   { $pogingen = @('GET') }
        default { $pogingen = @('POST', 'GET') }   # AUTO: eerst de standaard (POST), anders GET zoals in Postman
    }

    $fouten = New-Object System.Collections.Generic.List[string]
    foreach ($m in $pogingen) {
        try {
            if ($m -eq 'GET') { $r = Invoke-RawHttpWithBody -Method 'GET' -Url $cfg.TokenUrl -FormBody $form }
            else { $r = Invoke-SkvHttp -Method 'POST' -Url $cfg.TokenUrl -FormBody $form }
        }
        catch {
            $fouten.Add("$m : $(Get-ErrorText $_)")
            continue
        }
        if ($r.Ok) {
            $json = $null
            try { $json = $r.Body | ConvertFrom-Json } catch { }
            if ($json -and $json.access_token) {
                $script:Token = [string]$json.access_token
                $seconden = 0
                if ($json.PSObject.Properties['expires_in'] -and $json.expires_in) { [void][int]::TryParse([string]$json.expires_in, [ref]$seconden) }
                if ($seconden -gt 120) { $script:TokenExpires = (Get-Date).AddSeconds($seconden - 60) }
                else { $script:TokenExpires = (Get-Date).AddHours(23) }
                $script:TokenMethodUsed = $m
                return $script:Token
            }
            $fouten.Add("$m : antwoord bevat geen access_token")
        }
        else {
            $fouten.Add("$m : HTTP $($r.StatusCode) $($r.Reason) $(Get-ResponseMessage $r.Body)".Trim())
        }
    }
    throw "Token ophalen mislukt.`n" + ($fouten -join "`n")
}

# ---------------------------------------------------------------------------
#  Certificaat-controle
# ---------------------------------------------------------------------------
function ConvertTo-Datum($v) {
    if ($null -eq $v -or [string]::IsNullOrWhiteSpace([string]$v)) { return $null }
    if ($v -is [datetime]) { return $v }
    if ($v -is [datetimeoffset]) { return $v.DateTime }
    $dto = [datetimeoffset]::MinValue
    if ([datetimeoffset]::TryParse([string]$v, [Globalization.CultureInfo]::InvariantCulture, [Globalization.DateTimeStyles]::None, [ref]$dto)) {
        return $dto.DateTime
    }
    return $null
}

# Datum als dd-MM-yyyy (lege tekst als er geen datum is)
function Format-Datum($v) {
    $d = ConvertTo-Datum $v
    if ($d) { return $d.ToString('dd-MM-yyyy') }
    return [string]$v
}

# Beoordeelt een antwoord. Los van het scherm, zodat het te testen is.
function Get-CertificaatResultaat {
    param(
        [string]$Ubn,
        [string]$Am91,
        $Http,            # resultaat van Invoke-SkvHttp, of $null bij verbindingsfout
        [string]$Fout     # tekst bij verbindingsfout
    )
    $res = [pscustomobject]@{
        Ubn            = $Ubn
        Am91           = $Am91
        Uitkomst       = 'FOUT'     # OK / LET OP / FOUT
        Status         = ''
        GeldigTot      = ''
        Am91Type       = ''
        Certificaatnr  = ''
        AantalDieren   = ''
        Melding        = ''
        HttpCode       = ''
        RawJson        = ''
        Item           = $null
        Tijdstip       = Get-Date
        Vki            = $null      # VKI-resultaat (zie New-VkiResultaat)
    }

    if (-not $Http) {
        $res.Melding = ($Fout -replace "\s*`r?`n\s*", ' | ')
        return $res
    }

    $res.HttpCode = "$($Http.StatusCode) $($Http.Reason)".Trim()
    $res.RawJson = $Http.Body

    if (-not $Http.Ok) {
        $msg = Get-ResponseMessage $Http.Body
        $res.Melding = "HTTP $($res.HttpCode)"
        if ($msg) { $res.Melding += " - $msg" }
        return $res
    }

    $parsed = $null
    try { $parsed = $Http.Body | ConvertFrom-Json }
    catch {
        $res.Melding = 'Antwoord is geen geldige JSON'
        return $res
    }

    if ($parsed -is [string]) {
        $res.Melding = $parsed
        return $res
    }

    $items = @(@($parsed) | Where-Object { $null -ne $_ -and $_ -isnot [string] })
    if ($items.Count -eq 0) {
        $res.Melding = 'Geen certificaat gevonden (lege lijst [])'
        return $res
    }

    # Bij meerdere certificaten: het certificaat dat het langst geldig is
    $item = $items | Sort-Object -Property @{ Expression = { $d = ConvertTo-Datum $_.geldigTotDatum; if ($d) { $d } else { [datetime]::MinValue } } } -Descending | Select-Object -First 1
    $res.Item = $item
    $res.Status = [string]$item.status
    $datum = ConvertTo-Datum $item.geldigTotDatum
    if ($datum) { $res.GeldigTot = $datum.ToString('dd-MM-yyyy') } else { $res.GeldigTot = [string]$item.geldigTotDatum }
    $res.Am91Type = [string]$item.am91type
    $res.Certificaatnr = [string]$item.certificaatnummer
    if ($item.PSObject.Properties['stallijst'] -and $item.stallijst) { $res.AantalDieren = [string]@($item.stallijst).Count }

    $meldingen = New-Object System.Collections.Generic.List[string]
    $res.Uitkomst = 'OK'
    if ($res.Status -ne 'GECERTIFICEERD') {
        $res.Uitkomst = 'LET OP'
        $meldingen.Add('Status is niet GECERTIFICEERD')
    }
    if ($datum -and $datum.Date -lt (Get-Date).Date) {
        $res.Uitkomst = 'LET OP'
        $meldingen.Add('Certificaat is verlopen')
    }
    if ($items.Count -gt 1) { $meldingen.Add("$($items.Count) certificaten gevonden, langst geldige getoond") }
    $res.Melding = $meldingen -join '; '
    return $res
}

function Invoke-CertificaatControle([string]$Ubn, [string]$Am91) {
    $url = [string]$script:Config.CertificaatUrl
    if ($url.Contains('?')) { $url += '&' } else { $url += '?' }
    $url += 'UBN=' + [Uri]::EscapeDataString($Ubn) + '&AM91nummer=' + [Uri]::EscapeDataString($Am91)

    try {
        $token = Get-SkvToken
    }
    catch {
        return Get-CertificaatResultaat -Ubn $Ubn -Am91 $Am91 -Http $null -Fout (Get-ErrorText $_)
    }

    try {
        $r = Invoke-SkvHttp -Method 'GET' -Url $url -BearerToken $token
        if ($r.StatusCode -eq 401) {
            # Token verlopen of geweigerd: 1x opnieuw proberen met een nieuw token
            $token = Get-SkvToken -Force
            $r = Invoke-SkvHttp -Method 'GET' -Url $url -BearerToken $token
        }
        return Get-CertificaatResultaat -Ubn $Ubn -Am91 $Am91 -Http $r
    }
    catch {
        return Get-CertificaatResultaat -Ubn $Ubn -Am91 $Am91 -Http $null -Fout ('Geen verbinding: ' + (Get-ErrorText $_))
    }
}

# ===========================================================================
#  VKI (Voedselketeninformatie) - SKV InfoKalf-webservice
# ===========================================================================

# ---------------------------------------------------------------------------
#  Aanroep (SOAP 1.1, inloggegevens in het bericht)
# ---------------------------------------------------------------------------
function New-VkiVerzoek([string]$Ubn, [string]$Am91, [string]$Username, [string]$Password) {
    $e = { param($s) [System.Security.SecurityElement]::Escape([string]$s) }
    return '<s:Envelope xmlns:s="http://schemas.xmlsoap.org/soap/envelope/"' +
           ' xmlns:n0="http://tempuri.org/"' +
           ' xmlns:n1="http://schemas.datacontract.org/2004/07/SKVPortaal.infoKalf.Webservice.VKIget.V2">' +
           '<s:Header/><s:Body><n0:GetVKIdieren>' +
           '<n0:username>' + (& $e $Username) + '</n0:username>' +
           '<n0:password>' + (& $e $Password) + '</n0:password>' +
           '<n0:request><n1:RequestVKIDier><n1:koppelnummer>' +
           '<n1:AM91>' + (& $e $Am91) + '</n1:AM91>' +
           '<n1:HouderUBN>' + (& $e $Ubn) + '</n1:HouderUBN>' +
           '</n1:koppelnummer></n1:RequestVKIDier></n0:request>' +
           '</n0:GetVKIdieren></s:Body></s:Envelope>'
}

function Get-VkiHttpClient {
    if (-not $script:VkiHttpClient) {
        $handler = New-Object System.Net.Http.HttpClientHandler
        $script:VkiHttpClient = New-Object System.Net.Http.HttpClient($handler)
        $script:VkiHttpClient.Timeout = [TimeSpan]::FromSeconds([int]$script:Vki.TimeoutSeconden)
        [void]$script:VkiHttpClient.DefaultRequestHeaders.TryAddWithoutValidation('User-Agent', $script:UserAgent)
    }
    return $script:VkiHttpClient
}

# Doet EEN VKI-aanroep (1 UBN + 1 AM91). Geeft ook bij een HTTP-fout de inhoud van het antwoord terug.
# Gooit een fout bij geen verbinding of time-out.
function Invoke-VkiHttp([string]$Ubn, [string]$Am91) {
    $body = New-VkiVerzoek -Ubn $Ubn -Am91 $Am91 -Username $script:Vki.Username -Password $script:Vki.Password
    $client = Get-VkiHttpClient
    $req = New-Object System.Net.Http.HttpRequestMessage([System.Net.Http.HttpMethod]::Post, [string]$script:Vki.Url)
    [void]$req.Headers.TryAddWithoutValidation('SOAPAction', $script:VkiSoapAction)
    $req.Content = New-Object System.Net.Http.StringContent($body, [Text.Encoding]::UTF8, 'text/xml')
    $resp = Wait-Task ($client.SendAsync($req))
    $tekst = Wait-Task ($resp.Content.ReadAsStringAsync())
    return New-HttpResult ([int]$resp.StatusCode) $resp.ReasonPhrase $tekst
}

# ---------------------------------------------------------------------------
#  XML uitlezen (op local-name, dan maken de prefixen niet uit)
# ---------------------------------------------------------------------------
function Get-XmlKind($Node, [string]$Naam) {
    if (-not $Node) { return $null }
    foreach ($c in $Node.ChildNodes) {
        if ($c.NodeType -eq [System.Xml.XmlNodeType]::Element -and $c.LocalName -eq $Naam) { return $c }
    }
    return $null
}

# Alle element-kinderen van het kind $Naam (bijv. alle formulieren onder VKIFormulieren)
function Get-XmlLijst($Node, [string]$Naam) {
    $lijst = New-Object System.Collections.Generic.List[object]
    $c = Get-XmlKind $Node $Naam
    if ($c) {
        foreach ($k in $c.ChildNodes) {
            if ($k.NodeType -eq [System.Xml.XmlNodeType]::Element) { $lijst.Add($k) }
        }
    }
    return ,$lijst
}

function Get-XmlTekst($Node, [string]$Naam) {
    $c = Get-XmlKind $Node $Naam
    if (-not $c) { return '' }
    if ($c.GetAttribute('nil', 'http://www.w3.org/2001/XMLSchema-instance') -eq 'true') { return '' }
    # Regeleinden gelijk trekken (SKV stuurt &#xD; + newline)
    return ([string]$c.InnerText -replace "`r`n", "`n" -replace "`r", "`n").Trim()
}

function New-XmlDocument([string]$Xml) {
    $doc = New-Object System.Xml.XmlDocument
    $doc.XmlResolver = $null
    $doc.LoadXml($Xml)
    return $doc
}

# Zet een VKI-antwoordcode om naar tekst zoals op de SKV-PDF.
# Let op: het SKV-document klopt niet. Vastgesteld: 1 = JA, 0 = NEE, -1 = niet ingevuld.
function ConvertTo-VkiAntwoord([string]$Code) {
    switch ($Code.Trim()) {
        '1'  { return 'JA' }
        '0'  { return 'NEE' }
        '-1' { return '---' }
        ''   { return '---' }
        '3'  { return 'NVT' }   # volgens document; nog niet in een echt antwoord gezien
        default { return "? ($Code)" }
    }
}

function ConvertFrom-VkiAspectNode($A) {
    $nr = 0
    if (-not [int]::TryParse((Get-XmlTekst $A 'AspectVolgnummer'), [ref]$nr)) { [void][int]::TryParse((Get-XmlTekst $A 'AspectCode'), [ref]$nr) }
    $code = Get-XmlTekst $A 'Antwoord'
    $asp = [pscustomobject]@{
        Nr           = $nr
        Code         = Get-XmlTekst $A 'AspectCode'
        Omschrijving = Get-XmlTekst $A 'AspectOmschrijving'
        Toelichting  = Get-XmlTekst $A 'AspectToelichting'
        AntwoordCode = $code
        Antwoord     = ConvertTo-VkiAntwoord $code
        Opmerking    = Get-XmlTekst $A 'Opmerking'
        Middelen     = @()
        Diagnoses    = @()
        Informaties  = @()
    }
    # Let op: lijsten in een object altijd als gewone array bewaren (een List in een object-eigenschap
    # geeft in PowerShell bij @(...) de fout "Argument types do not match").
    $middelen = New-Object System.Collections.Generic.List[object]
    $diagnoses = New-Object System.Collections.Generic.List[object]
    $informaties = New-Object System.Collections.Generic.List[object]
    foreach ($m in (Get-XmlLijst $A 'VKIDierMiddelen')) {
        $einde = ConvertTo-Datum (Get-XmlTekst $m 'EindeWachttermijn')
        $eerste = ''
        if ($einde) { $eerste = $einde.AddDays(1).ToString('dd-MM-yyyy') }   # zo rekent de SKV-PDF ook
        $middelen.Add([pscustomobject]@{
            Vraag             = $nr
            Land              = Get-XmlTekst $m 'VKIDierLand'
            Levensnummer      = Get-XmlTekst $m 'VKIDierLevensnummer'
            RegNl             = Get-XmlTekst $m 'REGNL'
            Naam              = Get-XmlTekst $m 'Naam'
            Begindatum        = Format-Datum (Get-XmlTekst $m 'Begindatum')
            Einddatum         = Format-Datum (Get-XmlTekst $m 'Einddatum')
            Wachttermijn      = Get-XmlTekst $m 'Wachttermijn'
            EindeWachttermijn = Format-Datum (Get-XmlTekst $m 'EindeWachttermijn')
            EersteLevering    = $eerste
        })
    }
    foreach ($d in (Get-XmlLijst $A 'VKIDierDiagnoses')) {
        $diagnoses.Add([pscustomobject]@{
            Vraag = $nr; Soort = 'Diagnose'
            Land = Get-XmlTekst $d 'VKIDierLand'; Levensnummer = Get-XmlTekst $d 'VKIDierLevensnummer'
            Tekst = Get-XmlTekst $d 'Diagnose'
        })
    }
    foreach ($i in (Get-XmlLijst $A 'VKIDierInformaties')) {
        $informaties.Add([pscustomobject]@{
            Vraag = $nr; Soort = 'Informatie'
            Land = Get-XmlTekst $i 'VKIDierLand'; Levensnummer = Get-XmlTekst $i 'VKIDierLevensnummer'
            Tekst = Get-XmlTekst $i 'Informatie'
        })
    }
    $asp.Middelen = $middelen.ToArray()
    $asp.Diagnoses = $diagnoses.ToArray()
    $asp.Informaties = $informaties.ToArray()
    return $asp
}

function ConvertFrom-VkiFormulierNode($F, [string]$BronXml) {
    $t = { param($n) Get-XmlTekst $F $n }
    $frm = [pscustomobject]@{
        VkiCode            = & $t 'VKICode'
        VkiVersie          = & $t 'VKIVersie'
        Aanleverdatum      = Format-Datum (& $t 'Aanleverdatum')
        DatumOndertekening = Format-Datum (& $t 'DatumOndertekening')
        OndertekendOp      = ConvertTo-Datum (& $t 'DatumOndertekening')
        NaarWaarheid       = (& $t 'NaarWaarheidIngevuld').ToLowerInvariant()
        HouderUbn          = & $t 'VKIHouderUBN'
        HouderNaam         = & $t 'VKIHouderNaamVolledig'
        HouderTelefoon     = & $t 'VKIHouderTelefoonnummer'
        HouderEmail        = & $t 'VKIHouderEmail'
        Bezoek             = [pscustomobject]@{ Straat = & $t 'VKIHouderBezoekStraat'; Huisnummer = & $t 'VKIHouderBezoekHuisnummer'; Toevoeging = & $t 'VKIHouderBezoekHuisnummerToevoeging'; Postcode = & $t 'VKIHouderBezoekPostcode'; Plaats = & $t 'VKIHouderBezoekPlaats' }
        Post               = [pscustomobject]@{ Straat = & $t 'VKIHouderPostStraat'; Huisnummer = & $t 'VKIHouderPostHuisnummer'; Toevoeging = & $t 'VKIHouderPostHuisnummerToevoeging'; Postcode = & $t 'VKIHouderPostPostcode'; Plaats = & $t 'VKIHouderPostPlaats' }
        Dierenarts         = [pscustomobject]@{ Naam = & $t 'VKIDierenartsNaam'; Straat = & $t 'VKIDierenartsStraat'; Huisnummer = & $t 'VKIDierenartsHuisnummer'; Toevoeging = & $t 'VKIDierenartsHuisnummerToevoeging'; Postcode = & $t 'VKIDierenartsPostcode'; Plaats = & $t 'VKIDierenartsPlaats'; Telefoon = & $t 'VKIDierenartsTelefoon' }
        Voorletters        = & $t 'Voorletters'
        Tussenvoegsel      = & $t 'Tussenvoegsel'
        Achternaam         = & $t 'Achternaam'
        Plaats             = & $t 'Plaats'
        Aspecten           = @()
        Dieren             = @()
        Koppelnummers      = @()
        AantalDieren       = 0
        DierenZonderAm91   = 0
        Middelen           = @()
        DiagnosesEnInfo    = @()
        BronXml            = $BronXml
    }

    $aspecten = New-Object System.Collections.Generic.List[object]
    foreach ($a in (Get-XmlLijst $F 'VKIAspecten')) { $aspecten.Add((ConvertFrom-VkiAspectNode $a)) }
    $frm.Aspecten = @($aspecten | Sort-Object Nr)

    $dieren = New-Object System.Collections.Generic.List[object]
    $koppels = New-Object System.Collections.Generic.List[string]
    foreach ($d in (Get-XmlLijst $F 'VKIDieren')) {
        # Snelle variant (een formulier kan meer dan 1000 dieren hebben): kinderen 1x doorlopen
        $am = ''; $land = ''; $lev = ''
        foreach ($c in $d.ChildNodes) {
            if ($c.NodeType -ne [System.Xml.XmlNodeType]::Element) { continue }
            switch ($c.LocalName) {
                'AM91Nummer'   { $am = $c.InnerText.Trim() }
                'Land'         { $land = $c.InnerText.Trim() }
                'Levensnummer' { $lev = $c.InnerText.Trim() }
            }
        }
        $dieren.Add([pscustomobject]@{ Am91 = $am; Land = $land; Levensnummer = $lev })
        if ($am -eq '') { $frm.DierenZonderAm91++ }
        elseif (-not $koppels.Contains($am)) { $koppels.Add($am) }
    }
    $frm.Dieren = $dieren.ToArray()
    $frm.Koppelnummers = $koppels.ToArray()
    $frm.AantalDieren = $frm.Dieren.Count

    $frm.Middelen = @($frm.Aspecten | ForEach-Object { $_.Middelen })
    $frm.DiagnosesEnInfo = @($frm.Aspecten | ForEach-Object { $_.Diagnoses; $_.Informaties })
    return $frm
}

# Leest een volledig VKI-antwoord (SOAP-envelope). Gooit geen fouten; zie ParseFout/SoapFout.
function ConvertFrom-VkiXml([string]$Xml) {
    $res = [pscustomobject]@{
        HasErrors   = $false
        Fouten      = @()
        Formulieren = @()
        SoapFout    = ''
        ParseFout   = ''
    }
    if ([string]::IsNullOrWhiteSpace($Xml)) { $res.ParseFout = 'Leeg antwoord ontvangen'; return $res }
    try { $doc = New-XmlDocument $Xml }
    catch { $res.ParseFout = 'Antwoord is geen geldige XML'; return $res }

    $fault = $doc.SelectSingleNode("//*[local-name()='Fault']")
    if ($fault) {
        $tekst = Get-XmlTekst $fault 'faultstring'
        if (-not $tekst) { $tekst = ([string]$fault.InnerText).Trim() }
        $res.SoapFout = $tekst
        return $res
    }
    $result = $doc.SelectSingleNode("//*[local-name()='GetVKIdierenResult']")
    if (-not $result) { $res.ParseFout = 'Onverwacht antwoord (GetVKIdierenResult ontbreekt)'; return $res }

    $res.HasErrors = ((Get-XmlTekst $result 'HasErrors').ToLowerInvariant() -eq 'true')
    $fouten = New-Object System.Collections.Generic.List[object]
    foreach ($e in (Get-XmlLijst $result 'ErrorItems')) {
        $fouten.Add([pscustomobject]@{
            Code     = Get-XmlTekst $e 'ErrorCode'
            Tekst    = Get-XmlTekst $e 'ErrorText'
            ItemData = Get-XmlTekst $e 'ItemData'
        })
    }
    $formulieren = New-Object System.Collections.Generic.List[object]
    foreach ($f in (Get-XmlLijst $result 'VKIFormulieren')) {
        $formulieren.Add((ConvertFrom-VkiFormulierNode $f $Xml))
    }
    $res.Fouten = $fouten.ToArray()
    $res.Formulieren = $formulieren.ToArray()
    return $res
}

# Vertaalt een SKV-foutcode naar een begrijpelijke melding
function Get-VkiFoutTekst([string]$Code, [string]$Tekst) {
    switch ($Code) {
        'ERR0002' { return 'Inloggen bij VKI geweigerd (ERR0002): verkeerde gebruikersnaam/wachtwoord, geen rechten of account tijdelijk geblokkeerd' }
        'ERR0100' { return 'AM91-nummer ontbreekt (ERR0100)' }
        'ERR0101' { return 'Levensnummer ontbreekt (ERR0101)' }
        'ERR0103' { return 'UBN ontbreekt (ERR0103)' }
        'ERR0104' { return 'Landcode ontbreekt (ERR0104)' }
        'ERR0105' { return 'UBN ongeldig volgens SKV (ERR0105)' }
        'ERR0106' { return 'Geen VKI gevonden voor dit UBN + AM91 (SKV: ERR0106)' }
        'ERR0107' { return 'AM91-nummer niet bekend bij dit UBN (SKV: ERR0107)' }
        'ERR0999' { return 'Fout bij SKV (ERR0999), probeer het later opnieuw' }
        default {
            if ($Tekst -and $Code) { return "$Tekst ($Code)" }
            if ($Tekst) { return $Tekst }
            return "Onbekende fout ($Code)"
        }
    }
}

# ---------------------------------------------------------------------------
#  Beoordeling (VOORLOPIGE kleurregels, zie INSTRUCTIE hoofdstuk 5.4)
# ---------------------------------------------------------------------------
function Get-VkiAspect($Formulier, [int]$Nr) {
    foreach ($a in $Formulier.Aspecten) { if ($a.Nr -eq $Nr) { return $a } }
    return $null
}

function Get-VkiAntwoordVan($Formulier, [int]$Nr) {
    $a = Get-VkiAspect $Formulier $Nr
    if ($a) { return $a.Antwoord }
    return '---'
}

# REGEL VOOR DIERGENEESMIDDELEN - nog niet bekend (verwacht van de business).
# Nu: middelen alleen tonen en oranje markeren als er middelen zijn; nooit rood.
# De echte regel komt later HIER (en alleen hier).
function Get-VkiMiddelenStatus($Formulier) {
    $n = @($Formulier.Middelen).Count
    if ($n -eq 0) { return [pscustomobject]@{ Uitkomst = 'OK'; Melding = '' } }
    $tekst = "$n regel(s) diergeneesmiddelen"
    if ($n -eq 1) { $tekst = '1 regel diergeneesmiddelen' }
    return [pscustomobject]@{ Uitkomst = 'LET OP'; Melding = $tekst }
}

# Zichtkeuring: vraag 10 = JA, 11/12/13 = JA, 14/15/16 = NEE,
# en als vraag 17 = JA ook 18 en 20 = JA (19 = JA of NVT).   (VOORLOPIG, key-user laten bevestigen)
function Get-VkiZichtkeuring($Formulier) {
    $a = { param($n) Get-VkiAntwoordVan $Formulier $n }
    $redenen = New-Object System.Collections.Generic.List[string]
    if ((& $a 10) -ne 'JA') {
        return [pscustomobject]@{ Zichtkeuring = 'nee'; Reden = "vraag 10 = $(& $a 10)" }
    }
    foreach ($n in 11, 12, 13) { if ((& $a $n) -ne 'JA') { $redenen.Add("vraag $n = $(& $a $n)") } }
    foreach ($n in 14, 15, 16) { if ((& $a $n) -ne 'NEE') { $redenen.Add("vraag $n = $(& $a $n)") } }
    $v17 = & $a 17
    if ($v17 -eq 'JA') {
        foreach ($n in 18, 20) { if ((& $a $n) -ne 'JA') { $redenen.Add("vraag $n = $(& $a $n)") } }
        if ((& $a 19) -notin @('JA', 'NVT')) { $redenen.Add("vraag 19 = $(& $a 19)") }
    }
    elseif ($v17 -ne 'NEE') { $redenen.Add("vraag 17 = $v17") }

    if ($redenen.Count -eq 0) { return [pscustomobject]@{ Zichtkeuring = 'ja'; Reden = '' } }
    return [pscustomobject]@{ Zichtkeuring = 'nee'; Reden = ($redenen -join ', ') }
}

function Get-VkiBeoordeling($Formulier, [string]$Am91) {
    $rood = New-Object System.Collections.Generic.List[string]
    $oranje = New-Object System.Collections.Generic.List[string]

    if (-not $Formulier) {
        return [pscustomobject]@{ Uitkomst = 'FOUT'; Meldingen = @('Geen VKI-formulier'); Zicht = $null; JaVragen = @() }
    }
    if ($Formulier.NaarWaarheid -ne 'true') { $rood.Add('Niet "naar waarheid ingevuld"') }
    $v1 = Get-VkiAntwoordVan $Formulier 1
    if ($v1 -ne 'JA') { $rood.Add("Vraag 1 (wachttermijnen nageleefd) = $v1") }
    if ($Formulier.DierenZonderAm91 -gt 0) {
        if ($Formulier.DierenZonderAm91 -eq 1) { $rood.Add('1 dier zonder AM91-nummer op het formulier') }
        else { $rood.Add("$($Formulier.DierenZonderAm91) dieren zonder AM91-nummer op het formulier") }
    }
    if ($Am91 -and ($Formulier.Koppelnummers -notcontains $Am91)) { $rood.Add("AM91 $Am91 staat niet op het formulier") }

    $jaBijVragen = @(2..9 | Where-Object { (Get-VkiAntwoordVan $Formulier $_) -eq 'JA' })
    if ($jaBijVragen.Count -gt 0) { $oranje.Add('JA bij vraag ' + ($jaBijVragen -join ', ')) }
    $middelen = Get-VkiMiddelenStatus $Formulier
    if ($middelen.Uitkomst -eq 'FOUT') { $rood.Add($middelen.Melding) }
    elseif ($middelen.Uitkomst -eq 'LET OP') { $oranje.Add($middelen.Melding) }
    $nDiag = @($Formulier.DiagnosesEnInfo | Where-Object { $_.Soort -eq 'Diagnose' }).Count
    if ($nDiag -gt 0) { $oranje.Add("$nDiag diagnose(s)") }

    $uitkomst = 'OK'
    if ($rood.Count -gt 0) { $uitkomst = 'FOUT' }
    elseif ($oranje.Count -gt 0) { $uitkomst = 'LET OP' }

    $alleJa = @($Formulier.Aspecten | Where-Object { $_.Antwoord -eq 'JA' } | ForEach-Object { $_.Nr })
    $meldingen = @($rood) + @($oranje)
    return [pscustomobject]@{ Uitkomst = $uitkomst; Meldingen = $meldingen; Zicht = (Get-VkiZichtkeuring $Formulier); JaVragen = $alleJa }
}

# ---------------------------------------------------------------------------
#  Resultaat per AM91
# ---------------------------------------------------------------------------
function New-VkiResultaat([string]$Ubn, [string]$Am91) {
    return [pscustomobject]@{
        Ubn             = $Ubn
        Am91            = $Am91
        Uitkomst        = 'FOUT'      # OK / LET OP / FOUT / UIT
        Melding         = ''
        Formulier       = $null
        VkiCode         = ''
        AantalAm91      = ''
        AantalFormulier = ''
        Ondertekend     = ''
        Zichtkeuring    = ''
        ZichtReden      = ''
        JaVragen        = @()
        Fouten          = @()
        HttpCode        = ''
        RawXml          = ''
        Tijdstip        = Get-Date
    }
}

function Set-VkiResultaatUitFormulier($Res, $Formulier, [string[]]$ExtraMeldingen) {
    $b = Get-VkiBeoordeling $Formulier $Res.Am91
    $Res.Formulier = $Formulier
    $Res.Uitkomst = $b.Uitkomst
    $Res.VkiCode = $Formulier.VkiCode
    $n = 0
    foreach ($d in $Formulier.Dieren) { if ($d.Am91 -eq $Res.Am91) { $n++ } }
    $Res.AantalAm91 = [string]$n
    $Res.AantalFormulier = [string]$Formulier.AantalDieren
    $Res.Ondertekend = $Formulier.DatumOndertekening
    $Res.JaVragen = $b.JaVragen
    if ($b.Zicht) { $Res.Zichtkeuring = $b.Zicht.Zichtkeuring; $Res.ZichtReden = $b.Zicht.Reden }
    $alle = @($b.Meldingen) + @($ExtraMeldingen | Where-Object { $_ })
    $Res.Melding = ($alle -join '; ')
}

# Zoekt in de al opgehaalde formulieren het (meest recente) formulier met dit AM91
function Find-VkiFormulier($Formulieren, [string]$Am91) {
    $kandidaten = @($Formulieren | Where-Object { $_.Koppelnummers -contains $Am91 })
    if ($kandidaten.Count -eq 0) { return $null }
    $beste = $kandidaten | Sort-Object -Property @(
        @{ Expression = { if ($_.OndertekendOp) { $_.OndertekendOp } else { [datetime]::MinValue } }; Descending = $true },
        @{ Expression = { $v = 0; [void][int]::TryParse($_.VkiVersie, [ref]$v); $v }; Descending = $true }
    ) | Select-Object -First 1
    return [pscustomobject]@{ Formulier = $beste; Aantal = $kandidaten.Count }
}

<#
    De VKI-controle voor alle AM91-nummers van een UBN.
    - per AM91 een eigen aanroep (SKV koppelt fouten niet aan een AM91);
    - formulieren ontdubbelen op VKICode; staat een AM91 al op een opgehaald formulier, dan geen nieuwe aanroep;
    - ERR0002 (inlog): direct stoppen, nooit opnieuw proberen;
    - ERR0105 (UBN ongeldig): overige AM91's krijgen dezelfde melding zonder aanroep;
    - geen verbinding / time-out: overige AM91's "niet gecontroleerd".
    $Aanroep is een scriptblock { param($Ubn, $Am91) ... } dat een HttpResult teruggeeft (of een fout gooit),
    zodat deze functie ook zonder netwerk te testen is.
#>
function Invoke-VkiControle {
    param(
        [string]$Ubn,
        [string[]]$Am91Lijst,
        [scriptblock]$Aanroep,
        [scriptblock]$Voortgang
    )
    $resultaten = New-Object System.Collections.Generic.List[object]
    $formulieren = @{}          # VKICode -> formulier
    $stopMelding = ''
    $stopXml = ''
    $aanroepen = 0

    if (-not $script:Vki.Actief) {
        foreach ($am in $Am91Lijst) {
            $r = New-VkiResultaat $Ubn $am
            $r.Uitkomst = 'UIT'
            $r.Melding = $script:Vki.Reden
            $resultaten.Add($r)
        }
        return [pscustomobject]@{ Resultaten = $resultaten.ToArray(); Aanroepen = 0 }
    }
    if ($script:VkiInlogGeweigerd) {
        $stopMelding = 'Niet gecontroleerd: VKI-inlog is eerder geweigerd (ERR0002). Controleer appsettings.json en start het programma opnieuw (na minimaal 10 minuten).'
    }

    foreach ($am in $Am91Lijst) {
        $res = New-VkiResultaat $Ubn $am
        if ($stopMelding) {
            $res.Melding = $stopMelding
            $res.RawXml = $stopXml
            $resultaten.Add($res)
            continue
        }

        # Staat dit AM91 al op een eerder opgehaald formulier? Dan geen nieuwe aanroep.
        $gevonden = Find-VkiFormulier $formulieren.Values $am
        if ($gevonden) {
            Set-VkiResultaatUitFormulier $res $gevonden.Formulier @('zelfde formulier als eerder AM91-nummer')
            $res.RawXml = $gevonden.Formulier.BronXml
            $res.HttpCode = '(niet opnieuw opgevraagd)'
            $resultaten.Add($res)
            continue
        }

        if ($Voortgang) { & $Voortgang $am }
        try {
            $aanroepen++
            $http = & $Aanroep $Ubn $am
        }
        catch {
            if (Test-IsTimeout $_) { $res.Melding = "Time-out: geen antwoord van VKI binnen $($script:Vki.TimeoutSeconden) seconden" }
            else { $res.Melding = 'Geen verbinding met VKI: ' + ((Get-ErrorText $_) -replace "\s*`r?`n\s*", ' | ') }
            $stopMelding = 'Niet gecontroleerd (VKI niet bereikbaar)'
            $resultaten.Add($res)
            continue
        }

        $res.HttpCode = "$($http.StatusCode) $($http.Reason)".Trim()
        $res.RawXml = $http.Body
        $antw = ConvertFrom-VkiXml $http.Body

        if ($antw.SoapFout) {
            $res.Melding = "SOAP-fout van SKV (HTTP $($res.HttpCode)): $($antw.SoapFout)"
            $resultaten.Add($res)
            continue
        }
        if ($antw.ParseFout) {
            if ($http.Ok) { $res.Melding = $antw.ParseFout }
            else { $res.Melding = "HTTP $($res.HttpCode): $($antw.ParseFout)" }
            $resultaten.Add($res)
            continue
        }
        if ($antw.HasErrors -or $antw.Fouten.Count -gt 0) {
            $res.Fouten = @($antw.Fouten)
            $codes = @($antw.Fouten | ForEach-Object { $_.Code })
            $teksten = @($antw.Fouten | ForEach-Object { Get-VkiFoutTekst $_.Code $_.Tekst })
            if ($teksten.Count -eq 0) { $teksten = @('SKV meldt een fout zonder omschrijving') }
            $res.Melding = ($teksten -join '; ')
            if ($codes -contains 'ERR0002') {
                $script:VkiInlogGeweigerd = $true
                $res.Melding += '. Niet opnieuw geprobeerd.'
                $stopMelding = 'Niet gecontroleerd: VKI-inlog geweigerd (ERR0002)'
            }
            elseif ($codes -contains 'ERR0105') {
                $stopMelding = Get-VkiFoutTekst 'ERR0105' ''
                $stopXml = $http.Body
            }
            $resultaten.Add($res)
            continue
        }

        # Formulieren bewaren, ontdubbeld op VKICode (bij dubbel: hoogste versie)
        foreach ($f in $antw.Formulieren) {
            $sleutel = $f.VkiCode
            if (-not $sleutel) { $sleutel = [guid]::NewGuid().ToString() }
            if ($formulieren.ContainsKey($sleutel)) {
                $oud = 0; $nieuw = 0
                [void][int]::TryParse($formulieren[$sleutel].VkiVersie, [ref]$oud)
                [void][int]::TryParse($f.VkiVersie, [ref]$nieuw)
                if ($nieuw -gt $oud) { $formulieren[$sleutel] = $f }
            }
            else { $formulieren[$sleutel] = $f }
        }

        $gevonden = Find-VkiFormulier @($antw.Formulieren) $am
        if (-not $gevonden) {
            if ($antw.Formulieren.Count -eq 0) { $res.Melding = 'Geen VKI-formulier ontvangen' }
            else { $res.Melding = "VKI-antwoord bevat geen formulier met AM91 $am" }
            $resultaten.Add($res)
            continue
        }
        $extra = @()
        if ($gevonden.Aantal -gt 1) { $extra = @("$($gevonden.Aantal) formulieren gevonden, meest recente getoond") }
        Set-VkiResultaatUitFormulier $res $gevonden.Formulier $extra
        $resultaten.Add($res)
    }
    return [pscustomobject]@{ Resultaten = $resultaten.ToArray(); Aanroepen = $aanroepen }
}

# Voor testen zonder netwerk: verwerk een (opgeslagen) VKI-antwoord voor 1 of meer AM91-nummers.
# Voorbeeld:  Get-VkiResultaat -Xml (Get-Content .\test\vki-voorbeeld.xml -Raw) -Am91 @('581234')
function Get-VkiResultaat {
    param(
        [string]$Xml,
        [string[]]$Am91,
        [string]$Ubn = '0000000'
    )
    $oud = $script:Vki
    $script:Vki = [pscustomobject]@{ Actief = $true; Reden = ''; Url = 'offline'; Username = 'test'; Password = 'test'; TimeoutSeconden = 60; Waarschuwing = '' }
    try {
        # $Xml is via de aanroepketen zichtbaar in het scriptblock
        $uit = Invoke-VkiControle -Ubn $Ubn -Am91Lijst $Am91 -Aanroep { param($u, $a) New-HttpResult 200 'OK' $Xml }
        return $uit.Resultaten
    }
    finally { $script:Vki = $oud }
}

# ---------------------------------------------------------------------------
#  XML netjes inspringen
# ---------------------------------------------------------------------------
function Format-XmlTekst([string]$Xml) {
    if ([string]::IsNullOrWhiteSpace($Xml)) { return '' }
    try {
        $doc = New-XmlDocument $Xml
        $sw = New-Object System.IO.StringWriter
        $set = New-Object System.Xml.XmlWriterSettings
        $set.Indent = $true
        $set.IndentChars = '  '
        $set.NewLineChars = "`r`n"
        $set.OmitXmlDeclaration = $true
        $w = [System.Xml.XmlWriter]::Create($sw, $set)
        $doc.Save($w)
        $w.Close()
        return $sw.ToString()
    }
    catch { return $Xml }
}

# ---------------------------------------------------------------------------
#  VKI-formulier als HTML (opbouw gelijk aan de SKV InfoKalf-PDF)
# ---------------------------------------------------------------------------
function ConvertTo-HtmlTekst([string]$Tekst) {
    if ($null -eq $Tekst) { return '' }
    return ([System.Net.WebUtility]::HtmlEncode($Tekst) -replace "`r?`n", '<br>')
}

function Join-Adres($Straat, $Huisnummer, $Toevoeging) {
    $s = ([string]$Straat + ' ' + [string]$Huisnummer).Trim()
    if ($Toevoeging) { $s = ($s + ' ' + [string]$Toevoeging).Trim() }
    return $s
}

function Get-VkiBestandsnaam($Formulier, [string]$Ubn, [string]$Extensie) {
    $am = ($Formulier.Koppelnummers -join '-')
    if (-not $am) { $am = 'geen-AM91' }
    $naam = "VKI_$($Ubn)_$($am)_$($Formulier.VkiCode).$Extensie"
    foreach ($c in [IO.Path]::GetInvalidFileNameChars()) { $naam = $naam.Replace([string]$c, '_') }
    return $naam
}

function New-VkiHtml($Formulier, [datetime]$GeprintOp) {
    $f = $Formulier
    $h = { param($s) ConvertTo-HtmlTekst ([string]$s) }
    $sb = New-Object System.Text.StringBuilder
    $geprint = 'Geprint op ' + $GeprintOp.ToString('dd-MM-yyyy - HH:mm')

    $veld = {
        param($label, $waarde, [switch]$Ruimte)
        $cls = ''
        if ($Ruimte) { $cls = ' class="ruimte"' }
        [void]$sb.Append("<tr$cls><th>").Append((& $h $label)).Append('</th><td>').Append((& $h $waarde)).Append('</td></tr>')
    }

    [void]$sb.Append(@"
<!DOCTYPE html>
<html lang="nl"><head><meta charset="utf-8">
<title>VKI Formulier $(& $h $f.VkiCode)</title>
<style>
@page {
  size: A4;
  margin: 15mm 15mm 18mm 15mm;
  @bottom-left  { content: "$geprint"; font: 7.5pt 'Segoe UI', Arial, sans-serif; color: #666; }
  @bottom-right { content: "Pagina " counter(page) " van " counter(pages); font: 7.5pt 'Segoe UI', Arial, sans-serif; color: #666; }
}
* { box-sizing: border-box; }
body { font-family: 'Segoe UI', 'Open Sans', Arial, sans-serif; font-size: 9pt; color: #444; margin: 0; line-height: 1.4; }
h1 { color: #2b87c8; font-size: 14pt; margin: 0 0 14pt 0; }
h2 { color: #2b87c8; font-size: 10pt; margin: 14pt 0 7pt 0; page-break-after: avoid; }
table.velden { border-collapse: collapse; width: 100%; }
table.velden th { text-align: left; font-weight: 600; width: 45%; padding: 4pt 0; vertical-align: top; }
table.velden td { padding: 4pt 0; vertical-align: top; }
table.velden tr.ruimte th, table.velden tr.ruimte td { padding-top: 12pt; }
.vraag { display: grid; grid-template-columns: 1fr 45pt; column-gap: 10pt; margin: 0 0 8pt 0; page-break-inside: avoid; }
.vraag .nr { font-weight: 700; font-size: 10.5pt; margin-right: 4pt; }
.vraag .antw { font-weight: 700; text-align: left; }
.vraag .opm { margin-top: 3pt; font-style: italic; }
table.lijst { border: 1px solid #bbb; border-collapse: separate; border-spacing: 0; width: 100%; margin: 2pt 0 10pt 0; font-size: 8pt; }
table.lijst th { text-align: left; font-weight: 600; vertical-align: bottom; padding: 4pt 4pt; }
table.lijst td { vertical-align: top; padding: 3pt 4pt; }
table.lijst tbody tr:nth-child(odd) td { background: #eee; }
table.lijst thead { display: table-header-group; }
table.lijst tr { page-break-inside: avoid; }
table.dieren { border-collapse: collapse; width: 100%; font-size: 8pt; }
table.dieren td { width: 25%; padding: 0.5pt 0; }
.voetnoot { font-style: italic; font-size: 8pt; margin-top: 8pt; }
.blok { page-break-inside: avoid; break-inside: avoid; }
.bron { font-size: 7pt; color: #888; margin-top: 10pt; border-top: 1px solid #ddd; padding-top: 4pt; }
</style></head><body>
<h1>VKI Formulier</h1>
"@)

    # --- Gegevens kalverhouder ---
    [void]$sb.Append('<div class="blok"><h2>Gegevens kalverhouder</h2><table class="velden">')
    & $veld 'UBN:' $f.HouderUbn
    & $veld 'Naam:' $f.HouderNaam
    & $veld 'Telefoon:' $f.HouderTelefoon
    & $veld 'Bezoekadres:' (Join-Adres $f.Bezoek.Straat $f.Bezoek.Huisnummer $f.Bezoek.Toevoeging) -Ruimte
    & $veld 'Postcode/Plaats:' (($f.Bezoek.Postcode + ' ' + $f.Bezoek.Plaats).Trim())
    & $veld 'Postadres:' (Join-Adres $f.Post.Straat $f.Post.Huisnummer $f.Post.Toevoeging) -Ruimte
    & $veld 'Postcode/Plaats:' (($f.Post.Postcode + ' ' + $f.Post.Plaats).Trim())
    & $veld 'Koppelnummer(s):' ($f.Koppelnummers -join ', ') -Ruimte
    & $veld 'Aantal dieren:' ([string]$f.AantalDieren)
    [void]$sb.Append('</table></div>')

    # --- Vragen in 4 blokken ---
    $blokken = @(
        @{ Kop = 'Aspecten van af te leveren dier/dieren'; Van = 1;  Tot = 4 },
        @{ Kop = 'Informatie over herkomst bedrijf';       Van = 5;  Tot = 10 },
        @{ Kop = 'Informatie voor zichtkeuring';           Van = 11; Tot = 17 },
        @{ Kop = 'Informatie startkalveren';               Van = 18; Tot = 20 }
    )
    $gehad = @{}
    foreach ($blok in $blokken) {
        $vragen = @($f.Aspecten | Where-Object { $_.Nr -ge $blok.Van -and $_.Nr -le $blok.Tot })
        if ($vragen.Count -eq 0) { continue }
        [void]$sb.Append('<h2>').Append((& $h $blok.Kop)).Append('</h2>')
        foreach ($a in $vragen) { $gehad[$a.Nr] = $true; Add-VkiVraagHtml $sb $a }
    }
    # Eventuele extra vragen (nummer buiten 1-20)
    $overig = @($f.Aspecten | Where-Object { -not $gehad.ContainsKey($_.Nr) })
    if ($overig.Count -gt 0) {
        [void]$sb.Append('<h2>Overige vragen</h2>')
        foreach ($a in $overig) { Add-VkiVraagHtml $sb $a }
    }

    # --- Dierenarts ---
    [void]$sb.Append('<div class="blok"><h2>Gegevens dierenarts</h2><table class="velden">')
    & $veld 'Naam:' $f.Dierenarts.Naam
    & $veld 'Straat:' (Join-Adres $f.Dierenarts.Straat $f.Dierenarts.Huisnummer $f.Dierenarts.Toevoeging)
    & $veld 'Postcode:' $f.Dierenarts.Postcode
    & $veld 'Telefoon:' $f.Dierenarts.Telefoon
    [void]$sb.Append('</table></div>')

    # --- Ondertekening ---
    [void]$sb.Append('<div class="blok"><h2>Aldus verklaard</h2><table class="velden">')
    & $veld 'Voorletters:' $f.Voorletters
    & $veld 'Tussenvoegsel:' $f.Tussenvoegsel
    & $veld 'Achternaam:' $f.Achternaam
    & $veld 'Plaats:' $f.Plaats
    & $veld 'Datum ondertekening:' $f.DatumOndertekening
    [void]$sb.Append('</table></div>')

    # --- Diernummers (4 kolommen, gesorteerd zoals op de SKV-PDF) ---
    [void]$sb.Append('<h2>Diernummers</h2><table class="dieren">')
    $nummers = @($f.Dieren | Sort-Object Land, Levensnummer | ForEach-Object { ($_.Land + ' ' + $_.Levensnummer).Trim() })
    for ($i = 0; $i -lt $nummers.Count; $i += 4) {
        [void]$sb.Append('<tr>')
        for ($k = $i; $k -lt $i + 4; $k++) {
            $w = ''
            if ($k -lt $nummers.Count) { $w = & $h $nummers[$k] }
            [void]$sb.Append('<td>').Append($w).Append('</td>')
        }
        [void]$sb.Append('</tr>')
    }
    [void]$sb.Append('</table>')
    [void]$sb.Append('<div class="voetnoot">De diernummers, die op het VKI formulier staan vermeld, zijn allen kalveren uit het door SKV gecertificeerde AM91 koppel(s) en de diernummers van kalveren met een leeftijd van minimaal 5 maanden tot 1 jaar, die op het moment van opmaak van het VKI formulier op het bedrijf (UBN) aanwezig waren. Vaststelling van aanwezige kalveren op het UBN geschiet o.b.v. de offici&euml;le I&amp;R database. Daarnaast kan de kalverhouder individuele ID codes van kalveren hebben geselecteerd en/of toegevoegd aan het VKI formulier.</div>')

    [void]$sb.Append('<div class="bron">Samengesteld door Partijcontrole Veal uit SKV InfoKalf VKI-gegevens (VKI-code ').Append((& $h $f.VkiCode)).Append(', versie ').Append((& $h $f.VkiVersie)).Append('). ').Append((& $h $geprint)).Append('. Dit is geen offici&euml;le SKV-print.</div>')
    [void]$sb.Append('</body></html>')
    return $sb.ToString()
}

function Add-VkiVraagHtml($Sb, $A) {
    $h = { param($s) ConvertTo-HtmlTekst ([string]$s) }
    [void]$Sb.Append('<div class="vraag"><div><span class="nr">').Append($A.Nr).Append('</span> ').Append((& $h $A.Omschrijving))
    if ($A.Opmerking) { [void]$Sb.Append('<div class="opm">Opmerking: ').Append((& $h $A.Opmerking)).Append('</div>') }
    [void]$Sb.Append('</div><div class="antw">').Append((& $h $A.Antwoord)).Append('</div></div>')

    if ($A.Middelen.Count -gt 0) {
        [void]$Sb.Append('<table class="lijst"><thead><tr><th>Levensnummer</th><th>Reg NL<br>/ Diergeneesmiddel</th><th>Begin<br>behandeling</th><th>Einde<br>behandeling</th><th>Wacht-<br>termijn</th><th>Einde<br>wachttermijn</th><th>Eerste<br>levering</th></tr></thead><tbody>')
        foreach ($m in $A.Middelen) {
            $mid = ([string]$m.RegNl + ' ' + [string]$m.Naam).Trim()
            [void]$Sb.Append('<tr><td>').Append((& $h (($m.Land + ' ' + $m.Levensnummer).Trim()))).Append('</td><td>').Append((& $h $mid))
            [void]$Sb.Append('</td><td>').Append((& $h $m.Begindatum)).Append('</td><td>').Append((& $h $m.Einddatum))
            [void]$Sb.Append('</td><td>').Append((& $h $m.Wachttermijn)).Append('</td><td>').Append((& $h $m.EindeWachttermijn))
            [void]$Sb.Append('</td><td>').Append((& $h $m.EersteLevering)).Append('</td></tr>')
        }
        [void]$Sb.Append('</tbody></table>')
    }
    $extra = @($A.Diagnoses) + @($A.Informaties)
    if ($extra.Count -gt 0) {
        [void]$Sb.Append('<table class="lijst"><thead><tr><th style="width:25%">Levensnummer</th><th style="width:15%">Soort</th><th>Omschrijving</th></tr></thead><tbody>')
        foreach ($d in $extra) {
            [void]$Sb.Append('<tr><td>').Append((& $h (($d.Land + ' ' + $d.Levensnummer).Trim()))).Append('</td><td>').Append((& $h $d.Soort))
            [void]$Sb.Append('</td><td>').Append((& $h $d.Tekst)).Append('</td></tr>')
        }
        [void]$Sb.Append('</tbody></table>')
    }
}

# ---------------------------------------------------------------------------
#  HTML -> PDF met Microsoft Edge (headless). Geen extra software nodig.
# ---------------------------------------------------------------------------
function Find-Edge {
    $kandidaten = @()
    if (${env:ProgramFiles(x86)}) { $kandidaten += (Join-Path ${env:ProgramFiles(x86)} 'Microsoft\Edge\Application\msedge.exe') }
    if ($env:ProgramFiles) { $kandidaten += (Join-Path $env:ProgramFiles 'Microsoft\Edge\Application\msedge.exe') }
    if ($env:LOCALAPPDATA) { $kandidaten += (Join-Path $env:LOCALAPPDATA 'Microsoft\Edge\Application\msedge.exe') }
    foreach ($k in $kandidaten) { if (Test-Path $k) { return $k } }
    return $null
}

# Geeft @{ Ok; Fout } terug. De tijdelijke HTML (bevat persoonsgegevens) wordt altijd verwijderd.
function Convert-HtmlNaarPdf([string]$Html, [string]$PdfPad, [int]$TimeoutSeconden = 90) {
    $edge = Find-Edge
    if (-not $edge) { return [pscustomobject]@{ Ok = $false; Fout = 'Microsoft Edge is niet gevonden op deze pc.' } }

    $tmpMap = Join-Path ([IO.Path]::GetTempPath()) ('PartijcontroleVeal_' + [guid]::NewGuid().ToString('N'))
    [void](New-Item -ItemType Directory -Path $tmpMap)
    $htmlPad = Join-Path $tmpMap 'vki.html'
    $tmpPdf = Join-Path $tmpMap 'vki.pdf'
    $profiel = Join-Path $tmpMap 'edge-profiel'
    try {
        [IO.File]::WriteAllText($htmlPad, $Html, (New-Object System.Text.UTF8Encoding($true)))
        $uri = ([Uri]$htmlPad).AbsoluteUri
        # Eigen profielmap: anders geeft een al geopende Edge de opdracht door en wordt er niets geprint
        $argumenten = @(
            '--headless', '--disable-gpu', '--no-first-run', '--no-default-browser-check',
            '--disable-extensions', '--disable-sync',
            "--user-data-dir=`"$profiel`"",
            '--no-pdf-header-footer', '--print-to-pdf-no-header',
            "--print-to-pdf=`"$tmpPdf`"",
            "`"$uri`""
        )
        $p = Start-Process -FilePath $edge -ArgumentList $argumenten -PassThru -WindowStyle Hidden
        $tot = (Get-Date).AddSeconds($TimeoutSeconden)
        while (-not $p.HasExited -and (Get-Date) -lt $tot) {
            if ($script:UiPump) { & $script:UiPump }
            Start-Sleep -Milliseconds 100
        }
        if (-not $p.HasExited) {
            try { $p.Kill() } catch { }
            return [pscustomobject]@{ Ok = $false; Fout = "Edge reageerde niet binnen $TimeoutSeconden seconden." }
        }
        # Edge schrijft het bestand soms net na het afsluiten weg
        for ($i = 0; $i -lt 20 -and -not (Test-Path $tmpPdf); $i++) { Start-Sleep -Milliseconds 100 }
        if (-not (Test-Path $tmpPdf) -or (Get-Item $tmpPdf).Length -eq 0) {
            return [pscustomobject]@{ Ok = $false; Fout = 'Edge heeft geen PDF gemaakt.' }
        }
        Copy-Item -LiteralPath $tmpPdf -Destination $PdfPad -Force
        return [pscustomobject]@{ Ok = $true; Fout = '' }
    }
    catch {
        return [pscustomobject]@{ Ok = $false; Fout = (Get-ErrorText $_) }
    }
    finally {
        for ($i = 0; $i -lt 10 -and (Test-Path $tmpMap); $i++) {
            try { Remove-Item -LiteralPath $tmpMap -Recurse -Force -ErrorAction Stop } catch { Start-Sleep -Milliseconds 300 }
        }
    }
}

# ---------------------------------------------------------------------------
#  JSON netjes inspringen (zonder waarden aan te passen)
# ---------------------------------------------------------------------------
$script:JsonFormatterCode = @'
using System;
using System.Text;
public static class SkvJsonFormatter
{
    public static string Format(string json, bool shorten, int maxLen)
    {
        if (string.IsNullOrEmpty(json)) return "";
        StringBuilder sb = new StringBuilder(json.Length + 1024);
        int ind = 0;
        int n = json.Length;
        int i = 0;
        while (i < n)
        {
            char c = json[i];
            if (c == '"')
            {
                int j = i + 1;
                while (j < n)
                {
                    if (json[j] == '\\') { j += 2; continue; }
                    if (json[j] == '"') break;
                    j++;
                }
                if (j >= n) j = n - 1;
                int len = j - i + 1;
                if (shorten && len > maxLen)
                    sb.Append(json, i, 60).Append("... [").Append(len - 2).Append(" tekens ingekort]\"");
                else
                    sb.Append(json, i, len);
                i = j + 1;
                continue;
            }
            if (c == '{' || c == '[')
            {
                char close = (c == '{') ? '}' : ']';
                int k = i + 1;
                while (k < n && char.IsWhiteSpace(json[k])) k++;
                if (k < n && json[k] == close) { sb.Append(c).Append(close); i = k + 1; continue; }
                ind++;
                sb.Append(c).Append("\r\n").Append(' ', ind * 4);
            }
            else if (c == '}' || c == ']')
            {
                ind = Math.Max(0, ind - 1);
                sb.Append("\r\n").Append(' ', ind * 4).Append(c);
            }
            else if (c == ',') sb.Append(",\r\n").Append(' ', ind * 4);
            else if (c == ':') sb.Append(": ");
            else if (!char.IsWhiteSpace(c)) sb.Append(c);
            i++;
        }
        return sb.ToString();
    }
}
'@

# Zet JSON netjes onder elkaar (ingebouwde C#-code voor snelheid; waarden blijven ongewijzigd)
function Format-JsonText([string]$Json, [bool]$Inkorten, [int]$MaxLengte = 150) {
    if (-not ('SkvJsonFormatter' -as [type])) {
        Add-Type -TypeDefinition $script:JsonFormatterCode -Language CSharp
    }
    return [SkvJsonFormatter]::Format($Json, $Inkorten, $MaxLengte)
}

# ---------------------------------------------------------------------------
#  Export naar tekstbestand
# ---------------------------------------------------------------------------
function Get-ExportTekst($Ubn, $Resultaten) {
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.AppendLine('Partijcontrole Veal (SKV-certificaat en VKI)')
    [void]$sb.AppendLine('Datum/tijd : ' + (Get-Date).ToString('dd-MM-yyyy HH:mm:ss'))
    [void]$sb.AppendLine('UBN        : ' + $Ubn)
    [void]$sb.AppendLine(('-' * 100))
    foreach ($r in $Resultaten) {
        $label = '[' + $r.Uitkomst + ']'
        [void]$sb.Append($label.PadRight(10)).Append('AM91 ').Append($r.Am91.PadRight(12))
        if ($r.Item) {
            [void]$sb.Append('STATUS: ').Append($r.Status.PadRight(16)).Append('GELDIGTOTDATUM: ').Append($r.GeldigTot)
            [void]$sb.AppendLine()
            [void]$sb.Append(' ' * 27).Append('Type: ').Append($r.Am91Type).Append('   Certificaatnr: ').Append($r.Certificaatnr)
            if ($r.AantalDieren) { [void]$sb.Append('   Dieren op stallijst: ').Append($r.AantalDieren) }
            if ($r.Melding) { [void]$sb.AppendLine().Append(' ' * 27).Append('Let op: ').Append($r.Melding) }
        }
        else {
            [void]$sb.Append($r.Melding)
        }
        [void]$sb.AppendLine()

        $v = $r.Vki
        if ($v) {
            [void]$sb.Append(' ' * 10).Append(('VKI [' + $v.Uitkomst + ']').PadRight(17))
            if ($v.Formulier) {
                [void]$sb.Append('VKI-code: ').Append($v.VkiCode).Append('   Dieren AM91: ').Append($v.AantalAm91)
                [void]$sb.Append(' / formulier: ').Append($v.AantalFormulier).Append('   Ondertekend: ').Append($v.Ondertekend)
                [void]$sb.AppendLine()
                $ja = '-'
                if (@($v.JaVragen).Count -gt 0) { $ja = (@($v.JaVragen) -join ', ') }
                [void]$sb.Append(' ' * 27).Append('JA bij vraag: ').Append($ja).Append('   Zichtkeuring: ').Append($v.Zichtkeuring)
                if ($v.ZichtReden) { [void]$sb.Append(' (').Append($v.ZichtReden).Append(')') }
                if ($v.Melding) { [void]$sb.AppendLine().Append(' ' * 27).Append('Let op: ').Append($v.Melding) }
            }
            else {
                [void]$sb.Append($v.Melding)
            }
            [void]$sb.AppendLine()
        }
    }
    [void]$sb.AppendLine(('-' * 100))
    return $sb.ToString()
}

# Testmodus: alleen de functies laden, geen scherm tonen
if ($env:PARTIJCONTROLE_TEST -eq '1' -or $env:SKV_TEST -eq '1') { return }

# ===========================================================================
#  SCHERM
# ===========================================================================
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

$script:UiPump = { [System.Windows.Forms.Application]::DoEvents() }

function Show-Fout([string]$Tekst, [string]$Titel = $script:AppNaam) {
    [void][System.Windows.Forms.MessageBox]::Show($Tekst, $Titel, 'OK', 'Error')
}

try { $script:Config = Read-Config }
catch {
    Show-Fout ("Het programma kan niet starten.`n`n" + (Get-ErrorText $_))
    return
}
if ($script:Vki.Waarschuwing) {
    [void][System.Windows.Forms.MessageBox]::Show($script:Vki.Waarschuwing, $script:AppNaam, 'OK', 'Warning')
}

$fontNormaal = New-Object System.Drawing.Font('Segoe UI', 9)
$fontVet = New-Object System.Drawing.Font('Segoe UI', 10, [System.Drawing.FontStyle]::Bold)
$fontMono = New-Object System.Drawing.Font('Consolas', 9.5)
$kleurOk = [System.Drawing.Color]::FromArgb(0, 128, 0)
$kleurLetOp = [System.Drawing.Color]::FromArgb(214, 120, 0)
$kleurFout = [System.Drawing.Color]::FromArgb(200, 0, 0)
$kleurUit = [System.Drawing.Color]::DimGray
$achtergrondFout = [System.Drawing.Color]::FromArgb(255, 235, 235)
$teken = @{ 'OK' = [string][char]0x2713; 'LET OP' = '!'; 'FOUT' = '!'; 'UIT' = [string][char]0x2013; 'BEZIG' = [string][char]0x2026 }

function Get-UitkomstKleur([string]$Uitkomst) {
    switch ($Uitkomst) {
        'OK'     { return $kleurOk }
        'LET OP' { return $kleurLetOp }
        'UIT'    { return $kleurUit }
        'BEZIG'  { return $kleurUit }
        default  { return $kleurFout }
    }
}

$form = New-Object System.Windows.Forms.Form
$form.Text = $script:AppNaam
$form.Size = New-Object System.Drawing.Size(1320, 600)
$form.MinimumSize = New-Object System.Drawing.Size(900, 400)
$form.StartPosition = 'CenterScreen'
$form.Font = $fontNormaal
try {
    $ico = Get-ChildItem -Path $script:ScriptDir -Filter '*.ico' -File -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($ico) { $form.Icon = New-Object System.Drawing.Icon($ico.FullName) }
} catch { }

# --- UBN ---
$lblUbn = New-Object System.Windows.Forms.Label
$lblUbn.Text = 'UBN:'
$lblUbn.Location = New-Object System.Drawing.Point(12, 16)
$lblUbn.AutoSize = $true
$form.Controls.Add($lblUbn)

$txtUbn = New-Object System.Windows.Forms.TextBox
$txtUbn.Location = New-Object System.Drawing.Point(120, 12)
$txtUbn.Width = 140
$txtUbn.MaxLength = 15
$form.Controls.Add($txtUbn)

# --- AM91 ---
$lblAm = New-Object System.Windows.Forms.Label
$lblAm.Text = 'AM91-nummers:'
$lblAm.Location = New-Object System.Drawing.Point(12, 50)
$lblAm.AutoSize = $true
$form.Controls.Add($lblAm)

$alleenCijfers = {
    param($sender, $e)
    if (-not [char]::IsDigit($e.KeyChar) -and -not [char]::IsControl($e.KeyChar)) { $e.Handled = $true }
}
$txtUbn.Add_KeyPress($alleenCijfers)

$amVelden = @()
for ($k = 0; $k -lt 5; $k++) {
    $tb = New-Object System.Windows.Forms.TextBox
    $tb.Location = New-Object System.Drawing.Point((120 + $k * 125), 46)
    $tb.Width = 115
    $tb.MaxLength = 10
    $tb.Add_KeyPress($alleenCijfers)
    $form.Controls.Add($tb)
    $amVelden += $tb
}

# --- VKI aan/uit melding rechtsboven ---
$lblVki = New-Object System.Windows.Forms.Label
$lblVki.AutoSize = $true
$lblVki.Location = New-Object System.Drawing.Point(760, 50)
if ($script:Vki.Actief) {
    $lblVki.Text = 'VKI-controle: aan'
    $lblVki.ForeColor = $kleurOk
}
else {
    $lblVki.Text = 'VKI-controle: ' + $script:Vki.Reden
    $lblVki.ForeColor = $kleurUit
}
$form.Controls.Add($lblVki)

# --- Knoppen ---
$btnCheck = New-Object System.Windows.Forms.Button
$btnCheck.Text = 'Controleer'
$btnCheck.Location = New-Object System.Drawing.Point(12, 84)
$btnCheck.Size = New-Object System.Drawing.Size(110, 30)
$form.Controls.Add($btnCheck)
$form.AcceptButton = $btnCheck

$btnWis = New-Object System.Windows.Forms.Button
$btnWis.Text = 'Wissen'
$btnWis.Location = New-Object System.Drawing.Point(130, 84)
$btnWis.Size = New-Object System.Drawing.Size(90, 30)
$form.Controls.Add($btnWis)

$btnExport = New-Object System.Windows.Forms.Button
$btnExport.Text = 'Exporteren...'
$btnExport.Location = New-Object System.Drawing.Point(228, 84)
$btnExport.Size = New-Object System.Drawing.Size(110, 30)
$btnExport.Enabled = $false
$form.Controls.Add($btnExport)

$lblTip = New-Object System.Windows.Forms.Label
$lblTip.Text = 'Tip: dubbelklik op een regel voor alle details (SKV-certificaat en VKI, met PDF-knoppen).'
$lblTip.Location = New-Object System.Drawing.Point(350, 91)
$lblTip.AutoSize = $true
$lblTip.ForeColor = [System.Drawing.Color]::DimGray
$form.Controls.Add($lblTip)

# --- Resultatenlijst ---
$grid = New-Object System.Windows.Forms.DataGridView
$grid.Location = New-Object System.Drawing.Point(12, 124)
$grid.Size = New-Object System.Drawing.Size(1280, 400)
$grid.Anchor = 'Top, Bottom, Left, Right'
$grid.ReadOnly = $true
$grid.AllowUserToAddRows = $false
$grid.AllowUserToDeleteRows = $false
$grid.AllowUserToResizeRows = $false
$grid.RowHeadersVisible = $false
$grid.SelectionMode = 'FullRowSelect'
$grid.MultiSelect = $false
$grid.BackgroundColor = [System.Drawing.SystemColors]::Window
$grid.AutoSizeColumnsMode = 'None'
$grid.ColumnHeadersDefaultCellStyle.WrapMode = 'True'
$grid.ColumnHeadersHeightSizeMode = 'AutoSize'
$grid.ScrollBars = 'Both'

$kolommen = @(
    @{ Naam = 'Icoon';          Kop = '';                     Breedte = 32 },
    @{ Naam = 'Am91';           Kop = 'AM91';                 Breedte = 85 },
    @{ Naam = 'Uitkomst';       Kop = 'Resultaat certificaat'; Breedte = 80 },
    @{ Naam = 'Status';         Kop = 'STATUS';               Breedte = 115 },
    @{ Naam = 'GeldigTot';      Kop = 'GELDIGTOTDATUM';       Breedte = 105 },
    @{ Naam = 'Melding';        Kop = 'Melding certificaat';  Breedte = 180 },
    @{ Naam = 'VkiIcoon';       Kop = 'VKI';                  Breedte = 38 },
    @{ Naam = 'VkiCode';        Kop = 'VKI-code';             Breedte = 105 },
    @{ Naam = 'VkiAantalAm91';  Kop = 'Aantal dieren (AM91)'; Breedte = 70 },
    @{ Naam = 'VkiAantalForm';  Kop = 'Aantal dieren (formulier)'; Breedte = 75 },
    @{ Naam = 'VkiOndertekend'; Kop = 'Onder-tekend';         Breedte = 80 },
    @{ Naam = 'VkiZicht';       Kop = 'Zicht-keuring';        Breedte = 62 },
    @{ Naam = 'VkiMelding';     Kop = 'Melding VKI';          Breedte = 240 }
)
foreach ($kol in $kolommen) {
    $col = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $col.Name = $kol.Naam
    $col.HeaderText = $kol.Kop
    $col.Width = $kol.Breedte
    $col.SortMode = 'NotSortable'
    [void]$grid.Columns.Add($col)
}
$grid.Columns['Melding'].AutoSizeMode = 'Fill'
$grid.Columns['Melding'].FillWeight = 40
$grid.Columns['Melding'].MinimumWidth = 140
$grid.Columns['VkiMelding'].AutoSizeMode = 'Fill'
$grid.Columns['VkiMelding'].FillWeight = 60
$grid.Columns['VkiMelding'].MinimumWidth = 180
foreach ($n in @('Icoon', 'VkiIcoon')) {
    $grid.Columns[$n].DefaultCellStyle.Alignment = 'MiddleCenter'
    $grid.Columns[$n].DefaultCellStyle.Font = $fontVet
}
foreach ($n in @('VkiAantalAm91', 'VkiAantalForm')) { $grid.Columns[$n].DefaultCellStyle.Alignment = 'MiddleRight' }
# Scheiding tussen het certificaat-deel en het VKI-deel
$grid.Columns['Melding'].DividerWidth = 3
$form.Controls.Add($grid)

# --- Statusbalk ---
$statusStrip = New-Object System.Windows.Forms.StatusStrip
$statusLabel = New-Object System.Windows.Forms.ToolStripStatusLabel
$statusLabel.Text = 'Klaar. Vul een UBN en 1 tot 5 AM91-nummers in.'
[void]$statusStrip.Items.Add($statusLabel)
$form.Controls.Add($statusStrip)

$script:Resultaten = New-Object System.Collections.Generic.List[object]
$script:LaatsteUbn = ''

function Update-TokenStatus {
    if ($script:Token) {
        $statusLabel.Text += '   |   Token geldig tot ' + $script:TokenExpires.ToString('dd-MM HH:mm')
    }
}

function Add-ResultaatRegel($r) {
    $idx = $grid.Rows.Add($teken[$r.Uitkomst], $r.Am91, $r.Uitkomst, $r.Status, $r.GeldigTot, $r.Melding,
                          $teken['BEZIG'], '', '', '', '', '', 'nog niet opgevraagd')
    $row = $grid.Rows[$idx]
    $row.Tag = $script:Resultaten.Count - 1
    $kleur = Get-UitkomstKleur $r.Uitkomst
    $row.Cells['Icoon'].Style.ForeColor = $kleur
    $row.Cells['Icoon'].Style.SelectionForeColor = $kleur
    $row.Cells['Uitkomst'].Style.ForeColor = $kleur
    if ($r.Uitkomst -eq 'FOUT') {
        foreach ($n in @('Icoon', 'Am91', 'Uitkomst', 'Status', 'GeldigTot', 'Melding')) { $row.Cells[$n].Style.BackColor = $achtergrondFout }
    }
    $row.Cells['VkiIcoon'].Style.ForeColor = $kleurUit
    $row.Cells['VkiMelding'].Style.ForeColor = $kleurUit
    return $idx
}

function Update-VkiCellen([int]$RijIndex, $v) {
    $row = $grid.Rows[$RijIndex]
    $kleur = Get-UitkomstKleur $v.Uitkomst
    $row.Cells['VkiIcoon'].Value = $teken[$v.Uitkomst]
    $row.Cells['VkiCode'].Value = $v.VkiCode
    $row.Cells['VkiAantalAm91'].Value = $v.AantalAm91
    $row.Cells['VkiAantalForm'].Value = $v.AantalFormulier
    $row.Cells['VkiOndertekend'].Value = $v.Ondertekend
    $row.Cells['VkiZicht'].Value = $v.Zichtkeuring
    $melding = $v.Melding
    if ($v.Uitkomst -eq 'OK' -and -not $melding) { $melding = 'in orde' }
    $row.Cells['VkiMelding'].Value = $melding
    $row.Cells['VkiMelding'].ToolTipText = $melding
    $row.Cells['VkiIcoon'].Style.ForeColor = $kleur
    $row.Cells['VkiIcoon'].Style.SelectionForeColor = $kleur
    if ($v.Uitkomst -eq 'UIT') { $row.Cells['VkiMelding'].Style.ForeColor = $kleurUit }
    elseif ($v.Uitkomst -eq 'FOUT') { $row.Cells['VkiMelding'].Style.ForeColor = $kleurFout }
    elseif ($v.Uitkomst -eq 'LET OP') { $row.Cells['VkiMelding'].Style.ForeColor = $kleurLetOp }
    else { $row.Cells['VkiMelding'].Style.ForeColor = [System.Drawing.SystemColors]::ControlText }
    if ($v.Uitkomst -eq 'FOUT') {
        foreach ($n in @('VkiIcoon', 'VkiCode', 'VkiAantalAm91', 'VkiAantalForm', 'VkiOndertekend', 'VkiZicht', 'VkiMelding')) { $row.Cells[$n].Style.BackColor = $achtergrondFout }
    }
}

function Update-DetailTekst {
    $d = $script:Det
    if ([string]::IsNullOrEmpty($d.R.RawJson)) { $d.Txt.Text = '(geen antwoord ontvangen)' + "`r`n`r`n" + $d.R.Melding }
    else { $d.Txt.Text = Format-JsonText $d.R.RawJson $d.Chk.Checked }
    $d.Txt.SelectionStart = 0
    $d.Txt.SelectionLength = 0
}

function Save-Pdf([string]$Base64, [string]$Naam) {
    $sfd = New-Object System.Windows.Forms.SaveFileDialog
    $sfd.Filter = 'PDF-bestand (*.pdf)|*.pdf'
    $sfd.FileName = $Naam
    if ($sfd.ShowDialog() -eq 'OK') {
        try {
            [IO.File]::WriteAllBytes($sfd.FileName, [Convert]::FromBase64String($Base64))
            $antw = [System.Windows.Forms.MessageBox]::Show("Opgeslagen:`n$($sfd.FileName)`n`nNu openen?", 'PDF opgeslagen', 'YesNo', 'Information')
            if ($antw -eq 'Yes') { Start-Process $sfd.FileName }
        }
        catch { Show-Fout ("PDF opslaan mislukt.`n" + (Get-ErrorText $_)) }
    }
}

# Maakt het VKI-formulier als PDF (via Edge). Lukt dat niet, dan als HTML openen in de browser.
function Save-VkiPdf($V, $Eigenaar) {
    $f = $V.Formulier
    if (-not $f) { return }
    $sfd = New-Object System.Windows.Forms.SaveFileDialog
    $sfd.Filter = 'PDF-bestand (*.pdf)|*.pdf'
    $sfd.FileName = Get-VkiBestandsnaam $f $V.Ubn 'pdf'
    if ($sfd.ShowDialog() -ne 'OK') { return }
    $pad = $sfd.FileName

    $html = New-VkiHtml $f (Get-Date)
    $Eigenaar.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    try { $uit = Convert-HtmlNaarPdf $html $pad }
    finally { $Eigenaar.Cursor = [System.Windows.Forms.Cursors]::Default }

    if ($uit.Ok) {
        $antw = [System.Windows.Forms.MessageBox]::Show("Opgeslagen:`n$pad`n`nNu openen?", 'PDF opgeslagen', 'YesNo', 'Information')
        if ($antw -eq 'Yes') { Start-Process $pad }
        return
    }
    $htmlPad = [IO.Path]::ChangeExtension($pad, '.html')
    $vraag = "De PDF kon niet automatisch gemaakt worden:`n$($uit.Fout)`n`n" +
             "Wil je het formulier als HTML-bestand opslaan en openen in de browser?`n" +
             "Daar kies je dan zelf Afdrukken > Opslaan als PDF.`n`n$htmlPad"
    $antw = [System.Windows.Forms.MessageBox]::Show($vraag, 'PDF maken mislukt', 'YesNo', 'Warning')
    if ($antw -eq 'Yes') {
        try {
            [IO.File]::WriteAllText($htmlPad, $html, (New-Object System.Text.UTF8Encoding($true)))
            Start-Process $htmlPad
        }
        catch { Show-Fout ("HTML opslaan mislukt.`n" + (Get-ErrorText $_)) }
    }
}

function Save-VkiXml($V) {
    if (-not $V.RawXml) { return }
    $sfd = New-Object System.Windows.Forms.SaveFileDialog
    $sfd.Filter = 'XML-bestand (*.xml)|*.xml'
    if ($V.Formulier) { $sfd.FileName = Get-VkiBestandsnaam $V.Formulier $V.Ubn 'xml' }
    else { $sfd.FileName = "VKI_$($V.Ubn)_$($V.Am91)_antwoord.xml" }
    if ($sfd.ShowDialog() -eq 'OK') {
        try {
            $tekst = (Format-XmlTekst $V.RawXml) -replace "(?<!`r)`n", "`r`n"
            [IO.File]::WriteAllText($sfd.FileName, $tekst, (New-Object System.Text.UTF8Encoding($true)))
            [void][System.Windows.Forms.MessageBox]::Show("Opgeslagen:`n$($sfd.FileName)", 'XML opgeslagen', 'OK', 'Information')
        }
        catch { Show-Fout ("XML opslaan mislukt.`n" + (Get-ErrorText $_)) }
    }
}

function New-DetailGrid {
    $g = New-Object System.Windows.Forms.DataGridView
    $g.Dock = 'Fill'
    $g.ReadOnly = $true
    $g.AllowUserToAddRows = $false
    $g.AllowUserToDeleteRows = $false
    $g.AllowUserToResizeRows = $false
    $g.RowHeadersVisible = $false
    $g.SelectionMode = 'FullRowSelect'
    $g.BackgroundColor = [System.Drawing.SystemColors]::Window
    $g.AutoSizeRowsMode = 'AllCells'
    $g.DefaultCellStyle.WrapMode = 'True'
    $g.ColumnHeadersDefaultCellStyle.WrapMode = 'True'
    return $g
}

function Add-DetailKolom($Grid, [string]$Kop, [int]$Breedte, [switch]$Vullen) {
    $col = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $col.HeaderText = $Kop
    $col.Width = $Breedte
    $col.SortMode = 'NotSortable'
    if ($Vullen) { $col.AutoSizeMode = 'Fill' }
    [void]$Grid.Columns.Add($col)
}

# Vult het VKI-tabblad van het detailvenster
function Add-VkiTabInhoud($Tab, $V) {
    $info = New-Object System.Windows.Forms.Label
    $info.Location = New-Object System.Drawing.Point(8, 8)
    $info.Size = New-Object System.Drawing.Size(950, 38)
    $info.Anchor = 'Top, Left, Right'
    $t = "VKI: $($V.Uitkomst)     HTTP: $($V.HttpCode)     Opgevraagd: $($V.Tijdstip.ToString('dd-MM-yyyy HH:mm:ss'))"
    if ($V.Melding) { $t += "`n$($V.Melding)" }
    $info.Text = $t
    $info.ForeColor = Get-UitkomstKleur $V.Uitkomst
    if ($V.Uitkomst -eq 'OK') { $info.ForeColor = [System.Drawing.SystemColors]::ControlText }
    $Tab.Controls.Add($info)

    $btnVkiPdf = New-Object System.Windows.Forms.Button
    $btnVkiPdf.Text = 'VKI-formulier (PDF) opslaan'
    $btnVkiPdf.Location = New-Object System.Drawing.Point(8, 50)
    $btnVkiPdf.Size = New-Object System.Drawing.Size(200, 28)
    $btnVkiPdf.Enabled = [bool]$V.Formulier
    $Tab.Controls.Add($btnVkiPdf)

    $btnVkiXml = New-Object System.Windows.Forms.Button
    $btnVkiXml.Text = 'VKI-antwoord (XML) opslaan'
    $btnVkiXml.Location = New-Object System.Drawing.Point(216, 50)
    $btnVkiXml.Size = New-Object System.Drawing.Size(200, 28)
    $btnVkiXml.Enabled = [bool]$V.RawXml
    $Tab.Controls.Add($btnVkiXml)

    $btnVkiPdf.Add_Click({ Save-VkiPdf $script:Det.Vki $script:Det.Dlg })
    $btnVkiXml.Add_Click({ Save-VkiXml $script:Det.Vki })

    $f = $V.Formulier
    $sub = New-Object System.Windows.Forms.TabControl
    $sub.Location = New-Object System.Drawing.Point(8, 86)
    $sub.Size = New-Object System.Drawing.Size(950, 520)
    $sub.Anchor = 'Top, Bottom, Left, Right'
    $Tab.Controls.Add($sub)

    if ($f) {
        # --- Kopgegevens ---
        $tpKop = New-Object System.Windows.Forms.TabPage('Gegevens')
        $txtKop = New-Object System.Windows.Forms.TextBox
        $txtKop.Multiline = $true; $txtKop.ReadOnly = $true; $txtKop.ScrollBars = 'Vertical'
        $txtKop.Dock = 'Fill'; $txtKop.Font = $fontMono; $txtKop.BackColor = [System.Drawing.SystemColors]::Window
        $nWaar = 'nee'
        if ($f.NaarWaarheid -eq 'true') { $nWaar = 'ja' }
        $ondertekenaar = (@($f.Voorletters, $f.Tussenvoegsel, $f.Achternaam) | Where-Object { $_ }) -join ' '
        $zicht = $V.Zichtkeuring
        if ($V.ZichtReden) { $zicht += " ($($V.ZichtReden))" }
        $ja = '-'
        if (@($V.JaVragen).Count -gt 0) { $ja = (@($V.JaVragen) -join ', ') }
        $regels = @(
            "UBN                    : $($f.HouderUbn)",
            "Houder                 : $($f.HouderNaam)",
            "Telefoon / e-mail      : $($f.HouderTelefoon)  $($f.HouderEmail)",
            "Bezoekadres            : $(Join-Adres $f.Bezoek.Straat $f.Bezoek.Huisnummer $f.Bezoek.Toevoeging), $($f.Bezoek.Postcode) $($f.Bezoek.Plaats)",
            '',
            "Koppelnummer(s)        : $($f.Koppelnummers -join ', ')",
            "Aantal dieren          : $($f.AantalDieren)   (waarvan AM91 $($V.Am91): $($V.AantalAm91))",
            "Dieren zonder AM91     : $($f.DierenZonderAm91)",
            '',
            "VKI-code / versie      : $($f.VkiCode) / $($f.VkiVersie)",
            "Aanleverdatum          : $($f.Aanleverdatum)",
            "Datum ondertekening    : $($f.DatumOndertekening)   door $ondertekenaar, $($f.Plaats)",
            "Naar waarheid ingevuld : $nWaar",
            '',
            "JA bij vraag           : $ja",
            "Zichtkeuring           : $zicht",
            "Diergeneesmiddelen     : $(@($f.Middelen).Count) regel(s)",
            "Diagnoses / informatie : $(@($f.DiagnosesEnInfo).Count) regel(s)",
            '',
            "Dierenarts             : $($f.Dierenarts.Naam), $(Join-Adres $f.Dierenarts.Straat $f.Dierenarts.Huisnummer $f.Dierenarts.Toevoeging), $($f.Dierenarts.Postcode) $($f.Dierenarts.Plaats), tel. $($f.Dierenarts.Telefoon)",
            '',
            '(De levensnummers staan in de PDF.)'
        )
        $txtKop.Text = $regels -join "`r`n"
        $tpKop.Controls.Add($txtKop)
        $sub.TabPages.Add($tpKop)

        # --- Vragen ---
        $tpVr = New-Object System.Windows.Forms.TabPage('Vragen (20)')
        $gv = New-DetailGrid
        Add-DetailKolom $gv 'Nr' 36
        Add-DetailKolom $gv 'Vraag' 500 -Vullen
        Add-DetailKolom $gv 'Antwoord' 75
        Add-DetailKolom $gv 'Opmerking' 220
        $gv.Columns[1].DefaultCellStyle.WrapMode = 'False'
        $gv.AutoSizeRowsMode = 'None'
        foreach ($a in $f.Aspecten) {
            $eersteRegel = ($a.Omschrijving -split "`n")[0]
            $i = $gv.Rows.Add($a.Nr, $eersteRegel, $a.Antwoord, $a.Opmerking)
            $tip = $a.Omschrijving
            if ($a.Toelichting) { $tip += "`n`nToelichting: " + $a.Toelichting }
            $gv.Rows[$i].Cells[1].ToolTipText = $tip
            $kleur = $null
            if ($a.Nr -eq 1 -and $a.Antwoord -ne 'JA') { $kleur = $kleurFout }
            elseif ($a.Nr -ge 2 -and $a.Nr -le 9 -and $a.Antwoord -eq 'JA') { $kleur = $kleurLetOp }
            if ($kleur) { $gv.Rows[$i].Cells[2].Style.ForeColor = $kleur }
            $gv.Rows[$i].Cells[2].Style.Font = New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Bold)
        }
        $tpVr.Controls.Add($gv)
        $sub.TabPages.Add($tpVr)

        # --- Diergeneesmiddelen ---
        $tpMid = New-Object System.Windows.Forms.TabPage("Diergeneesmiddelen ($(@($f.Middelen).Count))")
        $gm = New-DetailGrid
        Add-DetailKolom $gm 'Vraag' 45
        Add-DetailKolom $gm 'Levensnummer' 115
        Add-DetailKolom $gm 'Reg NL' 70
        Add-DetailKolom $gm 'Diergeneesmiddel' 200 -Vullen
        Add-DetailKolom $gm 'Begin behandeling' 85
        Add-DetailKolom $gm 'Einde behandeling' 85
        Add-DetailKolom $gm 'Wacht-termijn' 60
        Add-DetailKolom $gm 'Einde wachttermijn' 85
        Add-DetailKolom $gm 'Eerste levering' 85
        foreach ($m in $f.Middelen) {
            [void]$gm.Rows.Add($m.Vraag, (($m.Land + ' ' + $m.Levensnummer).Trim()), $m.RegNl, $m.Naam, $m.Begindatum, $m.Einddatum, $m.Wachttermijn, $m.EindeWachttermijn, $m.EersteLevering)
        }
        $tpMid.Controls.Add($gm)
        $sub.TabPages.Add($tpMid)

        # --- Diagnoses en informatie ---
        $tpDi = New-Object System.Windows.Forms.TabPage("Diagnoses / informatie ($(@($f.DiagnosesEnInfo).Count))")
        $gd = New-DetailGrid
        Add-DetailKolom $gd 'Vraag' 45
        Add-DetailKolom $gd 'Soort' 80
        Add-DetailKolom $gd 'Levensnummer' 115
        Add-DetailKolom $gd 'Omschrijving' 400 -Vullen
        foreach ($d in $f.DiagnosesEnInfo) {
            [void]$gd.Rows.Add($d.Vraag, $d.Soort, (($d.Land + ' ' + $d.Levensnummer).Trim()), $d.Tekst)
        }
        $tpDi.Controls.Add($gd)
        $sub.TabPages.Add($tpDi)
    }

    # --- Ruw antwoord ---
    $tpXml = New-Object System.Windows.Forms.TabPage('Antwoord (XML)')
    $txtXml = New-Object System.Windows.Forms.TextBox
    $txtXml.Multiline = $true; $txtXml.ReadOnly = $true; $txtXml.ScrollBars = 'Both'; $txtXml.WordWrap = $false
    $txtXml.MaxLength = [int]::MaxValue
    $txtXml.Dock = 'Fill'; $txtXml.Font = $fontMono; $txtXml.BackColor = [System.Drawing.SystemColors]::Window
    if ($V.RawXml) { $txtXml.Text = (Format-XmlTekst $V.RawXml) -replace "(?<!`r)`n", "`r`n" }
    else { $txtXml.Text = '(geen antwoord ontvangen)' + "`r`n`r`n" + $V.Melding }
    $tpXml.Controls.Add($txtXml)
    $sub.TabPages.Add($tpXml)
}

# --- Detailvenster ---
function Show-Detail($r, [string]$StartTab = 'SKV') {
    $dlg = New-Object System.Windows.Forms.Form
    $dlg.Text = "Resultaat  UBN $($r.Ubn)  -  AM91 $($r.Am91)"
    $dlg.Size = New-Object System.Drawing.Size(1000, 720)
    $dlg.MinimumSize = New-Object System.Drawing.Size(700, 450)
    $dlg.StartPosition = 'CenterParent'
    $dlg.Font = $fontNormaal
    if ($form.Icon) { $dlg.Icon = $form.Icon }

    $tabs = New-Object System.Windows.Forms.TabControl
    $tabs.Dock = 'Fill'
    $dlg.Controls.Add($tabs)

    # ===== Tabblad SKV-certificaat (zoals voorheen) =====
    $tabSkv = New-Object System.Windows.Forms.TabPage("SKV-certificaat: $($r.Uitkomst)")
    $tabs.TabPages.Add($tabSkv)

    $info = New-Object System.Windows.Forms.Label
    $info.Location = New-Object System.Drawing.Point(8, 8)
    $info.Size = New-Object System.Drawing.Size(950, 36)
    $info.Anchor = 'Top, Left, Right'
    $t = "Resultaat: $($r.Uitkomst)     HTTP: $($r.HttpCode)     Opgevraagd: $($r.Tijdstip.ToString('dd-MM-yyyy HH:mm:ss'))"
    if ($r.AantalDieren) { $t += "     Dieren op stallijst: $($r.AantalDieren)" }
    if ($r.Melding) { $t += "`n$($r.Melding)" }
    $info.Text = $t
    $tabSkv.Controls.Add($info)

    $chk = New-Object System.Windows.Forms.CheckBox
    $chk.Text = 'Lange velden inkorten (PDF-data)'
    $chk.Checked = $true
    $chk.AutoSize = $true
    $chk.Location = New-Object System.Drawing.Point(8, 52)
    $tabSkv.Controls.Add($chk)

    $btnCopy = New-Object System.Windows.Forms.Button
    $btnCopy.Text = 'Kopieer JSON'
    $btnCopy.Location = New-Object System.Drawing.Point(258, 48)
    $btnCopy.Size = New-Object System.Drawing.Size(110, 28)
    $tabSkv.Controls.Add($btnCopy)

    $btnPdf = New-Object System.Windows.Forms.Button
    $btnPdf.Text = 'Certificaat-PDF opslaan'
    $btnPdf.Location = New-Object System.Drawing.Point(376, 48)
    $btnPdf.Size = New-Object System.Drawing.Size(170, 28)
    $tabSkv.Controls.Add($btnPdf)

    $btnStal = New-Object System.Windows.Forms.Button
    $btnStal.Text = 'Stallijst-PDF opslaan'
    $btnStal.Location = New-Object System.Drawing.Point(554, 48)
    $btnStal.Size = New-Object System.Drawing.Size(160, 28)
    $tabSkv.Controls.Add($btnStal)

    $item = $r.Item
    $btnPdf.Enabled = [bool]($item -and $item.PSObject.Properties['certificaatPDF'] -and $item.certificaatPDF)
    $btnStal.Enabled = [bool]($item -and $item.PSObject.Properties['stallijstPDF'] -and $item.stallijstPDF)

    $txt = New-Object System.Windows.Forms.TextBox
    $txt.Multiline = $true
    $txt.ReadOnly = $true
    $txt.ScrollBars = 'Both'
    $txt.WordWrap = $false
    $txt.MaxLength = [int]::MaxValue
    $txt.Font = $fontMono
    $txt.BackColor = [System.Drawing.SystemColors]::Window
    $txt.Location = New-Object System.Drawing.Point(8, 84)
    $txt.Size = New-Object System.Drawing.Size(950, 560)
    $txt.Anchor = 'Top, Bottom, Left, Right'
    $tabSkv.Controls.Add($txt)

    # ===== Tabblad VKI =====
    $v = $r.Vki
    $vkiKop = 'VKI'
    if ($v) { $vkiKop = "VKI: $($v.Uitkomst)" }
    $tabVki = New-Object System.Windows.Forms.TabPage($vkiKop)
    $tabs.TabPages.Add($tabVki)

    # Het venster is "modaal" (blokkeert het hoofdscherm), dus deze gedeelde variabelen zijn veilig
    $script:Det = @{ R = $r; Item = $item; Txt = $txt; Chk = $chk; Vki = $v; Dlg = $dlg }
    Update-DetailTekst
    $chk.Add_CheckedChanged({ Update-DetailTekst })

    if ($v) { Add-VkiTabInhoud $tabVki $v }
    else {
        $lbl = New-Object System.Windows.Forms.Label
        $lbl.Text = 'Geen VKI-gegevens voor deze regel.'
        $lbl.AutoSize = $true
        $lbl.Location = New-Object System.Drawing.Point(8, 8)
        $tabVki.Controls.Add($lbl)
    }

    $btnCopy.Add_Click({
        if ($script:Det.R.RawJson) { [System.Windows.Forms.Clipboard]::SetText((Format-JsonText $script:Det.R.RawJson $false)) }
    })

    $btnPdf.Add_Click({
        $d = $script:Det
        $naam = "$($d.R.Ubn)_$($d.R.Am91)_certificaat.pdf"
        if ($d.Item.PSObject.Properties['bestandnaam'] -and $d.Item.bestandnaam) { $naam = [string]$d.Item.bestandnaam }
        Save-Pdf ([string]$d.Item.certificaatPDF) $naam
    })

    $btnStal.Add_Click({
        $d = $script:Det
        Save-Pdf ([string]$d.Item.stallijstPDF) "$($d.R.Ubn)_$($d.R.Am91)_stallijst.pdf"
    })

    if ($StartTab -eq 'VKI') { $tabs.SelectedTab = $tabVki }
    [void]$dlg.ShowDialog($form)
    $dlg.Dispose()
}

# --- Acties ---
$btnCheck.Add_Click({
    $ubn = $txtUbn.Text.Trim()
    if ($ubn -notmatch '^\d+$') {
        Show-Fout 'Vul een geldig UBN in (alleen cijfers).'
        [void]$txtUbn.Focus()
        return
    }
    $lijst = New-Object System.Collections.Generic.List[string]
    foreach ($tb in $amVelden) {
        $v = $tb.Text.Trim()
        if ($v -eq '') { continue }
        if ($v -notmatch '^\d{1,10}$') {
            Show-Fout "AM91-nummer '$v' is ongeldig (alleen cijfers, maximaal 10)."
            [void]$tb.Focus()
            return
        }
        if (-not $lijst.Contains($v)) { $lijst.Add($v) }
    }
    if ($lijst.Count -eq 0) {
        Show-Fout 'Vul minimaal 1 AM91-nummer in.'
        [void]$amVelden[0].Focus()
        return
    }

    $btnCheck.Enabled = $false
    $btnWis.Enabled = $false
    $btnExport.Enabled = $false
    $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    $grid.Rows.Clear()
    $script:Resultaten.Clear()
    $script:LaatsteUbn = $ubn

    try {
        # ---- 1. SKV-certificaat ----
        # Eerst het token ophalen. Lukt dat niet, dan heeft de certificaatcontrole geen zin,
        # maar de VKI-controle (eigen inlog) gaat wel door.
        $statusLabel.Text = 'Token ophalen...'
        [System.Windows.Forms.Application]::DoEvents()
        $tokenFout = ''
        try { [void](Get-SkvToken) }
        catch {
            $tokenFout = Get-ErrorText $_
            $statusLabel.Text = 'Token ophalen mislukt.'
            Show-Fout ($tokenFout + "`n`nControleer de gebruikersnaam, het wachtwoord en TokenUrl in appsettings.json.")
        }

        $rijen = @{}
        $nr = 0
        foreach ($am in $lijst) {
            $nr++
            if ($tokenFout) {
                $r = Get-CertificaatResultaat -Ubn $ubn -Am91 $am -Http $null -Fout ('Token ophalen mislukt: ' + $tokenFout)
            }
            else {
                $statusLabel.Text = "Certificaat $nr van $($lijst.Count): AM91 $am ..."
                [System.Windows.Forms.Application]::DoEvents()
                $r = Invoke-CertificaatControle -Ubn $ubn -Am91 $am
            }
            $script:Resultaten.Add($r)
            $rijen[$am] = Add-ResultaatRegel $r
            [System.Windows.Forms.Application]::DoEvents()
        }

        # ---- 2. VKI (per AM91 een eigen aanroep) ----
        $statusLabel.Text = 'VKI ophalen...'
        [System.Windows.Forms.Application]::DoEvents()
        $vkiUit = Invoke-VkiControle -Ubn $ubn -Am91Lijst $lijst.ToArray() `
            -Aanroep { param($u, $a) Invoke-VkiHttp $u $a } `
            -Voortgang { param($a) $statusLabel.Text = "VKI ophalen: AM91 $a ..."; [System.Windows.Forms.Application]::DoEvents() }
        foreach ($v in $vkiUit.Resultaten) {
            foreach ($r in $script:Resultaten) { if ($r.Am91 -eq $v.Am91) { $r.Vki = $v } }
            if ($rijen.ContainsKey($v.Am91)) { Update-VkiCellen $rijen[$v.Am91] $v }
        }
        if ($script:VkiInlogGeweigerd -and $vkiUit.Aanroepen -gt 0 -and @($vkiUit.Resultaten | Where-Object { $_.Melding -like '*ERR0002*' }).Count -gt 0) {
            Show-Fout ("Inloggen bij de VKI-webservice is geweigerd (ERR0002).`n`n" +
                       "Controleer bij `"Vki`" in appsettings.json de gebruikersnaam en het wachtwoord.`n" +
                       "Het programma probeert het NIET opnieuw: SKV blokkeert het account 10 minuten na 3 foute pogingen.`n" +
                       "Start het programma na het aanpassen opnieuw (wacht bij twijfel 10 minuten).")
        }

        $ok = @($script:Resultaten | Where-Object { $_.Uitkomst -eq 'OK' }).Count
        $tekst = "Klaar. Certificaat: $ok van $($lijst.Count) OK."
        if ($script:Vki.Actief) {
            $vOk = @($vkiUit.Resultaten | Where-Object { $_.Uitkomst -eq 'OK' }).Count
            $tekst += "   VKI: $vOk van $($lijst.Count) OK ($($vkiUit.Aanroepen) aanroep(en))."
        }
        else { $tekst += '   VKI: ' + $script:Vki.Reden }
        $statusLabel.Text = $tekst
        Update-TokenStatus
    }
    catch {
        Show-Fout ("Onverwachte fout.`n" + (Get-ErrorText $_))
        $statusLabel.Text = 'Er ging iets mis.'
    }
    finally {
        $form.Cursor = [System.Windows.Forms.Cursors]::Default
        $btnCheck.Enabled = $true
        $btnWis.Enabled = $true
        $btnExport.Enabled = ($script:Resultaten.Count -gt 0)
    }
})

$btnWis.Add_Click({
    $txtUbn.Clear()
    foreach ($tb in $amVelden) { $tb.Clear() }
    $grid.Rows.Clear()
    $script:Resultaten.Clear()
    $btnExport.Enabled = $false
    $statusLabel.Text = 'Klaar. Vul een UBN en 1 tot 5 AM91-nummers in.'
    Update-TokenStatus
    [void]$txtUbn.Focus()
})

$btnExport.Add_Click({
    if ($script:Resultaten.Count -eq 0) { return }
    $sfd = New-Object System.Windows.Forms.SaveFileDialog
    $sfd.Filter = 'Tekstbestand (*.txt)|*.txt'
    $sfd.FileName = 'Partijcontrole_' + $script:LaatsteUbn + '_' + (Get-Date).ToString('yyyyMMdd_HHmm') + '.txt'
    $sfd.InitialDirectory = [Environment]::GetFolderPath('MyDocuments')
    if ($sfd.ShowDialog() -eq 'OK') {
        try {
            $tekst = Get-ExportTekst $script:LaatsteUbn $script:Resultaten
            [IO.File]::WriteAllText($sfd.FileName, $tekst, (New-Object System.Text.UTF8Encoding($true)))
            $statusLabel.Text = "Geexporteerd naar $($sfd.FileName)"
        }
        catch { Show-Fout ("Exporteren mislukt.`n" + (Get-ErrorText $_)) }
    }
})

$grid.Add_CellDoubleClick({
    param($sender, $e)
    if ($e.RowIndex -lt 0) { return }
    $idx = [int]$grid.Rows[$e.RowIndex].Tag
    $start = 'SKV'
    if ($e.ColumnIndex -ge 0 -and $grid.Columns[$e.ColumnIndex].Name -like 'Vki*') { $start = 'VKI' }
    Show-Detail $script:Resultaten[$idx] $start
})

$form.Add_Shown({ [void]$txtUbn.Focus() })
$form.Add_FormClosed({
    if ($script:HttpClient) { $script:HttpClient.Dispose() }
    if ($script:VkiHttpClient) { $script:VkiHttpClient.Dispose() }
})

[void]$form.ShowDialog()
$form.Dispose()
