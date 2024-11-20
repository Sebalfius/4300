$headers = @{
    "Content-Type" = "application/json"
    "Authorization" = "Bearer eyJhbGciOiJIUzUxMiJ9.eyJzdWIiOiJjb25vc3VyICxhcGlfc2ViYXN0aWFuICwyMDAuNjEuMTc4LjEyNiIsImV4cCI6MTczMjIxODA2NH0.AWOGWTPMVL3_X2rOzh1ZSJOcHSjOTRe_lujQN6aM4HTrJ-JfxmZD1OKNENGwu_Kwy8YjLXBsAF9ldSEdIoRpuQ" 
}

$responses = @()

$dateforurl = Get-Date -Format "dd/MM/yyyy"
#$yesterday = datetime.now() - timedelta(days=1)
$cuentas = 4300

foreach ($i in $cuentas) {
    $uri = "https://conosur.aunesa.com/Irmo/api/operaciones/informes?cuenta=$i&fechaDesde=$dateforurl&fechaHasta=$dateforurl"    
    
    try {
        $response = Invoke-WebRequest -Uri $uri -Method Get -Headers $headers
        $responses += ($response.Content | ConvertFrom-Json)
    } catch {
        Write-Host "Error informe de operaciones for account  $i : $_"
    }
}

$finalJson = $responses | ConvertTo-Json -Depth 100

$currentDate = Get-Date -Format "yyyy-MM-dd"
$fileName = "InformedeOperaciones_$currentDate.json"
$finalJson = $responses | ConvertTo-Json -Depth 100

$finalJson | Out-File -FilePath $fileName
