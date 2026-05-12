param(
  [int]$Port = 5173
)

$root = (Resolve-Path $PSScriptRoot).Path
[System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12
[System.Net.ServicePointManager]::Expect100Continue = $false

if ([string]::IsNullOrWhiteSpace($env:CLOUDSDK_CONFIG)) {
  $localGcloudConfig = Join-Path $env:LOCALAPPDATA "gcloud-config"
  New-Item -ItemType Directory -Force $localGcloudConfig | Out-Null
  $env:CLOUDSDK_CONFIG = $localGcloudConfig
}

$listener = [System.Net.Sockets.TcpListener]::new([System.Net.IPAddress]::Loopback, $Port)
$listener.Start()

Write-Host "Invoice Input PoC server"
Write-Host "Root: $root"
Write-Host "URL : http://127.0.0.1:$Port/"
Write-Host "Press Ctrl+C to stop."

function Get-MimeType {
  param([string]$Path)

  switch ([System.IO.Path]::GetExtension($Path).ToLowerInvariant()) {
    ".html" { "text/html; charset=utf-8"; break }
    ".css" { "text/css; charset=utf-8"; break }
    ".js" { "text/javascript; charset=utf-8"; break }
    ".json" { "application/json; charset=utf-8"; break }
    ".csv" { "text/csv; charset=utf-8"; break }
    ".xlsx" { "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"; break }
    ".xls" { "application/vnd.ms-excel"; break }
    ".xlsm" { "application/vnd.ms-excel.sheet.macroEnabled.12"; break }
    ".png" { "image/png"; break }
    ".jpg" { "image/jpeg"; break }
    ".jpeg" { "image/jpeg"; break }
    ".svg" { "image/svg+xml"; break }
    default { "application/octet-stream" }
  }
}

function Send-Response {
  param(
    [System.Net.Sockets.NetworkStream]$Stream,
    [int]$StatusCode,
    [string]$StatusText,
    [string]$ContentType,
    [byte[]]$Body,
    [hashtable]$ExtraHeaders = @{}
  )

  $header = "HTTP/1.1 $StatusCode $StatusText`r`nContent-Length: $($Body.Length)`r`nContent-Type: $ContentType`r`n"
  foreach ($key in $ExtraHeaders.Keys) {
    $header += "${key}: $($ExtraHeaders[$key])`r`n"
  }
  $header += "Connection: close`r`n`r`n"
  $headerBytes = [System.Text.Encoding]::ASCII.GetBytes($header)
  $Stream.Write($headerBytes, 0, $headerBytes.Length)
  if ($Body.Length -gt 0) {
    $Stream.Write($Body, 0, $Body.Length)
  }
}

function Send-Json {
  param(
    [System.Net.Sockets.NetworkStream]$Stream,
    [int]$StatusCode,
    [string]$StatusText,
    [string]$Json
  )

  $body = [System.Text.Encoding]::UTF8.GetBytes($Json)
  Send-Response $Stream $StatusCode $StatusText "application/json; charset=utf-8" $body
}

function Read-HttpRequest {
  param([System.Net.Sockets.NetworkStream]$Stream)

  $memory = [System.IO.MemoryStream]::new()
  $buffer = [byte[]]::new(16384)
  $headerEnd = -1
  $contentLength = 0
  $headerText = ""

  while ($true) {
    $read = $Stream.Read($buffer, 0, $buffer.Length)
    if ($read -le 0) {
      break
    }

    $memory.Write($buffer, 0, $read)
    $bytes = $memory.ToArray()
    $ascii = [System.Text.Encoding]::ASCII.GetString($bytes)
    $headerEnd = $ascii.IndexOf("`r`n`r`n")

    if ($headerEnd -ge 0) {
      $headerText = $ascii.Substring(0, $headerEnd)
      foreach ($line in ($headerText -split "`r`n")) {
        if ($line -match "^Content-Length:\s*(\d+)\s*$") {
          $contentLength = [int]$Matches[1]
        }
      }

      $expectedLength = $headerEnd + 4 + $contentLength
      if ($memory.Length -ge $expectedLength) {
        break
      }
    }
  }

  if ($memory.Length -eq 0 -or $headerEnd -lt 0) {
    return $null
  }

  $allBytes = $memory.ToArray()
  $bodyStart = $headerEnd + 4
  $bodyBytes = [byte[]]::new($contentLength)
  if ($contentLength -gt 0) {
    [Array]::Copy($allBytes, $bodyStart, $bodyBytes, 0, $contentLength)
  }

  $requestLine = ($headerText -split "`r`n")[0]
  return @{
    RequestLine = $requestLine
    Body = [System.Text.Encoding]::UTF8.GetString($bodyBytes)
  }
}

function Get-DotEnvValue {
  param([string]$Name)

  $envPath = Join-Path $root ".env"
  if ([System.IO.File]::Exists($envPath)) {
    foreach ($line in [System.IO.File]::ReadAllLines($envPath)) {
      $trimmed = $line.Trim()
      if ($trimmed.Length -eq 0 -or $trimmed.StartsWith("#")) {
        continue
      }

      $index = $trimmed.IndexOf("=")
      if ($index -le 0) {
        continue
      }

      $key = $trimmed.Substring(0, $index).Trim()
      if ($key -ne $Name) {
        continue
      }

      $value = $trimmed.Substring($index + 1).Trim()
      return $value.Trim('"').Trim("'")
    }
  }

  return [Environment]::GetEnvironmentVariable($Name)
}

function Get-GeminiProvider {
  $provider = Get-DotEnvValue "GEMINI_PROVIDER"
  if ([string]::IsNullOrWhiteSpace($provider)) {
    $provider = Get-DotEnvValue "GEMINI_API_MODE"
  }
  if ([string]::IsNullOrWhiteSpace($provider)) {
    return "generative"
  }

  $normalized = $provider.Trim().ToLowerInvariant()
  if ($normalized -eq "vertex" -or $normalized -eq "vertexai" -or $normalized -eq "vertex-ai") {
    return "vertex"
  }

  return "generative"
}

function Get-DefaultGeminiModel {
  $model = Get-DotEnvValue "GEMINI_DEFAULT_MODEL"
  if ([string]::IsNullOrWhiteSpace($model)) {
    return "gemini-3.1-flash-lite"
  }

  return $model.Trim()
}

function Normalize-GeminiModelName {
  param([string]$Model)

  if ([string]::IsNullOrWhiteSpace($Model)) {
    $Model = Get-DefaultGeminiModel
  }

  $model = $Model.Trim()
  $lowerModel = $model.ToLowerInvariant()
  $marker = "/models/"
  $markerIndex = $lowerModel.LastIndexOf($marker)
  if ($markerIndex -ge 0) {
    return $model.Substring($markerIndex + $marker.Length)
  }

  if ($lowerModel.StartsWith("models/")) {
    return $model.Substring(7)
  }

  return $model
}

function Get-GcloudCommandPath {
  $gcloud = Get-Command "gcloud" -ErrorAction SilentlyContinue
  if ($gcloud) {
    return $gcloud.Source
  }

  $candidates = @(
    (Join-Path $env:LOCALAPPDATA "Google\Cloud SDK\google-cloud-sdk\bin\gcloud.cmd"),
    "C:\Program Files (x86)\Google\Cloud SDK\google-cloud-sdk\bin\gcloud.cmd",
    "C:\Program Files\Google\Cloud SDK\google-cloud-sdk\bin\gcloud.cmd"
  )

  foreach ($candidate in $candidates) {
    if ([System.IO.File]::Exists($candidate)) {
      return $candidate
    }
  }

  return $null
}

function Activate-GcloudServiceAccountFromKey {
  param([string]$GcloudPath)

  $keyFile = Get-DotEnvValue "GOOGLE_APPLICATION_CREDENTIALS"
  if ([string]::IsNullOrWhiteSpace($keyFile)) {
    return $false
  }

  $resolvedKeyFile = [Environment]::ExpandEnvironmentVariables($keyFile.Trim())
  if (-not [System.IO.File]::Exists($resolvedKeyFile)) {
    Write-Host "[gcloud.auth] GOOGLE_APPLICATION_CREDENTIALS file was not found: $resolvedKeyFile"
    return $false
  }

  $projectId = Get-DotEnvValue "GOOGLE_CLOUD_PROJECT"
  $args = @("auth", "activate-service-account", "--key-file=$resolvedKeyFile")
  if (-not [string]::IsNullOrWhiteSpace($projectId)) {
    $args += "--project=$projectId"
  }

  $activateOutput = & $GcloudPath @args 2>&1
  if ($LASTEXITCODE -ne 0) {
    Write-Host "[gcloud.auth] service account activation failed: $activateOutput"
    return $false
  }

  return $true
}

function Get-GcloudAccessToken {
  $gcloudPath = Get-GcloudCommandPath
  if ([string]::IsNullOrWhiteSpace($gcloudPath)) {
    return $null
  }

  $tokenOutput = & $gcloudPath auth application-default print-access-token 2>&1
  if ($LASTEXITCODE -eq 0) {
    $token = (($tokenOutput | Select-Object -Last 1) -as [string]).Trim()
    if (-not [string]::IsNullOrWhiteSpace($token)) {
      return $token
    }
  }

  $tokenOutput = & $gcloudPath auth print-access-token 2>&1
  if ($LASTEXITCODE -eq 0) {
    $token = (($tokenOutput | Select-Object -Last 1) -as [string]).Trim()
    if (-not [string]::IsNullOrWhiteSpace($token)) {
      return $token
    }
  }

  if (Activate-GcloudServiceAccountFromKey $gcloudPath) {
    $tokenOutput = & $gcloudPath auth print-access-token 2>&1
    if ($LASTEXITCODE -eq 0) {
      $token = (($tokenOutput | Select-Object -Last 1) -as [string]).Trim()
      if (-not [string]::IsNullOrWhiteSpace($token)) {
        return $token
      }
    }
  }

  return $null
}

function Get-VertexAuthHeaders {
  $accessToken = Get-DotEnvValue "VERTEX_ACCESS_TOKEN"
  if ([string]::IsNullOrWhiteSpace($accessToken)) {
    $accessToken = Get-GcloudAccessToken
  }

  if (-not [string]::IsNullOrWhiteSpace($accessToken)) {
    return @{ "Authorization" = "Bearer $accessToken" }
  }

  $apiKey = Get-DotEnvValue "VERTEX_API_KEY"
  if (-not [string]::IsNullOrWhiteSpace($apiKey)) {
    return @{ "x-goog-api-key" = $apiKey }
  }

  throw "Vertex AI credentials were not found. Set VERTEX_ACCESS_TOKEN or VERTEX_API_KEY in .env, or run gcloud auth application-default login."
}

function Send-GeminiError {
  param(
    [System.Net.Sockets.NetworkStream]$Stream,
    [int]$StatusCode,
    [string]$Message
  )

  if ([string]::IsNullOrWhiteSpace($Message)) {
    $Message = "Gemini request failed."
  }

  if ($Message.TrimStart().StartsWith("{")) {
    Send-Json $Stream $StatusCode "Gemini Error" $Message
    return
  }

  $escaped = $Message.Replace("\", "\\").Replace('"', '\"').Replace("`r", "\r").Replace("`n", "\n")
  Send-Json $Stream $StatusCode "Gemini Error" "{""error"":{""message"":""$escaped""}}"
}

function Invoke-GeminiGenerateContent {
  param(
    [System.Net.Sockets.NetworkStream]$Stream,
    [string]$RequestBody
  )

  try {
    $requestJson = $RequestBody | ConvertFrom-Json
    $model = Normalize-GeminiModelName $requestJson.model

    $geminiBody = [ordered]@{
      contents = $requestJson.contents
      generationConfig = $requestJson.generationConfig
    } | ConvertTo-Json -Depth 100 -Compress

    $provider = Get-GeminiProvider
    $modelPath = [System.Uri]::EscapeDataString($model)
    $headers = @{}

    if ($provider -eq "vertex") {
      $projectId = Get-DotEnvValue "GOOGLE_CLOUD_PROJECT"
      if ([string]::IsNullOrWhiteSpace($projectId)) {
        throw "GOOGLE_CLOUD_PROJECT is missing in .env for Vertex AI."
      }

      $location = Get-DotEnvValue "GOOGLE_CLOUD_LOCATION"
      if ([string]::IsNullOrWhiteSpace($location)) {
        $location = "asia-northeast1"
      }

      $vertexHost = "aiplatform.googleapis.com"
      if ($location -ne "global") {
        $vertexHost = "$location-aiplatform.googleapis.com"
      }

      $projectPath = [System.Uri]::EscapeDataString($projectId.Trim())
      $locationPath = [System.Uri]::EscapeDataString($location.Trim())
      $uri = "https://$vertexHost/v1/projects/$projectPath/locations/$locationPath/publishers/google/models/$modelPath`:generateContent"
      $headers = Get-VertexAuthHeaders
    }
    else {
      $apiKey = Get-DotEnvValue "GEMINI_API_KEY"
      if ([string]::IsNullOrWhiteSpace($apiKey)) {
        throw "GEMINI_API_KEY is missing in .env"
      }

      $uri = "https://generativelanguage.googleapis.com/v1beta/models/$modelPath`:generateContent"
      $headers = @{ "x-goog-api-key" = $apiKey }
    }

    $response = Invoke-WebRequest `
      -Uri $uri `
      -Method Post `
      -Headers $headers `
      -ContentType "application/json" `
      -Body $geminiBody `
      -UseBasicParsing `
      -TimeoutSec 90

    $body = [System.Text.Encoding]::UTF8.GetBytes($response.Content)
    Send-Response $Stream ([int]$response.StatusCode) "OK" "application/json; charset=utf-8" $body
  }
  catch {
    $statusCode = 500
    $message = $_.Exception.Message
    $response = $_.Exception.Response

    if ($response) {
      try {
        $statusCode = [int]$response.StatusCode
        $reader = [System.IO.StreamReader]::new($response.GetResponseStream())
        $message = $reader.ReadToEnd()
        $reader.Dispose()
        if ([string]::IsNullOrWhiteSpace($message)) {
          $message = "Gemini request failed with HTTP $statusCode $($response.StatusDescription)."
        }
      }
      catch {
        $message = $_.Exception.Message
      }
    }

    if ([string]::IsNullOrWhiteSpace($message)) {
      $message = "Gemini request failed."
    }

    Write-Host "[gemini.error] status=$statusCode message=$message"
    Send-GeminiError $Stream $statusCode $message
  }
}

function Send-ExtractionTemplate {
  param([System.Net.Sockets.NetworkStream]$Stream)

  $localWorkbookExtensions = @(".xlsx", ".xls", ".xlsm", ".csv")
  $localWorkbook = Get-ChildItem -LiteralPath $root -File |
    Where-Object { $localWorkbookExtensions -contains $_.Extension.ToLowerInvariant() -and $_.Name -notlike "skyberry-import-*.xlsx" -and $_.Name -notlike "~$*" } |
    Sort-Object LastWriteTime -Descending |
    Select-Object -First 1

  if ($localWorkbook) {
    $bodyBytes = [System.IO.File]::ReadAllBytes($localWorkbook.FullName)
    Send-Response $Stream 200 "OK" (Get-MimeType $localWorkbook.FullName) $bodyBytes @{ "X-Template-Source" = "local-workbook"; "X-Template-Name" = [System.Uri]::EscapeDataString($localWorkbook.Name) }
    return
  }

  $templateUrl = Get-DotEnvValue "TEMPLATE_EXCEL_URL"
  if (-not [string]::IsNullOrWhiteSpace($templateUrl)) {
    $tempPath = [System.IO.Path]::GetTempFileName()
    try {
      $response = Invoke-WebRequest `
        -Uri $templateUrl `
        -Method Get `
        -OutFile $tempPath `
        -UseBasicParsing `
        -TimeoutSec 8

      $contentType = [string]$response.Headers["Content-Type"]
      if ([string]::IsNullOrWhiteSpace($contentType)) {
        $contentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
      }

      if (-not $contentType.ToLowerInvariant().Contains("text/html")) {
        $bodyBytes = [System.IO.File]::ReadAllBytes($tempPath)
        if ($bodyBytes.Length -gt 0) {
          Send-Response $Stream 200 "OK" $contentType $bodyBytes @{ "X-Template-Source" = "sharepoint" }
          return
        }
      }
    }
    catch {
    }
    finally {
      Remove-Item -LiteralPath $tempPath -Force -ErrorAction SilentlyContinue
    }
  }

  $fallbackPath = Join-Path $root "templates\extraction-fields.csv"
  if ([System.IO.File]::Exists($fallbackPath)) {
    $bodyBytes = [System.IO.File]::ReadAllBytes($fallbackPath)
    Send-Response $Stream 200 "OK" (Get-MimeType $fallbackPath) $bodyBytes @{ "X-Template-Source" = "local-fallback" }
    return
  }

  Send-Json $Stream 404 "Not Found" '{"error":{"message":"Template file was not found."}}'
}

function Write-ClientLog {
  param(
    [System.Net.Sockets.NetworkStream]$Stream,
    [string]$RequestBody
  )

  $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
  try {
    $payload = $RequestBody | ConvertFrom-Json
    $eventName = [string]$payload.event
    if ([string]::IsNullOrWhiteSpace($eventName)) {
      $eventName = "client.log"
    }

    $detail = "{}"
    if ($null -ne $payload.detail) {
      $detail = $payload.detail | ConvertTo-Json -Depth 50 -Compress
    }

    Write-Host "[$timestamp] [$eventName] $detail"
  }
  catch {
    Write-Host "[$timestamp] [client.log] $RequestBody"
  }

  Send-Json $Stream 200 "OK" '{"ok":true}'
}

try {
  while ($true) {
    $client = $listener.AcceptTcpClient()
    try {
      $stream = $client.GetStream()
      $request = Read-HttpRequest $stream
      if ($null -eq $request) {
        continue
      }

      $parts = $request.RequestLine -split " "
      $method = if ($parts.Length -ge 1) { $parts[0] } else { "GET" }
      $rawPath = if ($parts.Length -ge 2) { $parts[1] } else { "/" }
      $pathOnly = ($rawPath -split "\?")[0]

      if ($method -eq "GET" -and $pathOnly -eq "/api/health") {
        $provider = Get-GeminiProvider
        $defaultModel = Get-DefaultGeminiModel
        Send-Json $stream 200 "OK" "{""ok"":true,""provider"":""$provider"",""defaultModel"":""$defaultModel""}"
        continue
      }

      if ($method -eq "GET" -and $pathOnly -eq "/api/template/extraction-fields") {
        Send-ExtractionTemplate $stream
        continue
      }

      if ($method -eq "POST" -and $pathOnly -eq "/api/gemini/generate") {
        Invoke-GeminiGenerateContent $stream $request.Body
        continue
      }

      if ($method -eq "POST" -and $pathOnly -eq "/api/log") {
        Write-ClientLog $stream $request.Body
        continue
      }

      $decodedPath = [System.Uri]::UnescapeDataString($pathOnly).TrimStart("/")

      if ([string]::IsNullOrWhiteSpace($decodedPath)) {
        $decodedPath = "index.html"
      }

      $fileName = [System.IO.Path]::GetFileName($decodedPath)
      if ($fileName -eq ".env" -or $decodedPath.StartsWith(".git")) {
        $body = [System.Text.Encoding]::UTF8.GetBytes("403 Forbidden")
        Send-Response $stream 403 "Forbidden" "text/plain; charset=utf-8" $body
        continue
      }

      $candidate = [System.IO.Path]::GetFullPath((Join-Path $root $decodedPath))

      if (-not $candidate.StartsWith($root, [System.StringComparison]::OrdinalIgnoreCase)) {
        $body = [System.Text.Encoding]::UTF8.GetBytes("403 Forbidden")
        Send-Response $stream 403 "Forbidden" "text/plain; charset=utf-8" $body
        continue
      }

      if (-not [System.IO.File]::Exists($candidate)) {
        $body = [System.Text.Encoding]::UTF8.GetBytes("404 Not Found")
        Send-Response $stream 404 "Not Found" "text/plain; charset=utf-8" $body
        continue
      }

      $bodyBytes = [System.IO.File]::ReadAllBytes($candidate)
      Send-Response $stream 200 "OK" (Get-MimeType $candidate) $bodyBytes
    }
    catch {
      $message = $_.Exception.Message
      if ([string]::IsNullOrWhiteSpace($message)) {
        $message = "Local server request failed."
      }

      Write-Host "[server.error] $method $pathOnly $message"
      try {
        Send-GeminiError $stream 500 $message
      }
      catch {
      }
    }
    finally {
      $client.Close()
    }
  }
}
finally {
  $listener.Stop()
}
