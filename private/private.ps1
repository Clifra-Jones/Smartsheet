
$script:BaseURI = "https://api.smartsheet.com/2.0"

$global:SSFormat = $ServerInfo | ConvertFrom-Json

#Private function
function Read-Config () {
    $ConfigPath = "$home/.smartsheet/config.json"
    $config = Get-Content -Raw -Path $ConfigPath | ConvertFrom-Json
    return $config
}

function ConvertTo-UTime () {
    Param(
        [datetime]$DateTime
    )

    $uTime = ([System.DateTimeOffset]$DateTime).ToUnixTimeMilliseconds() / 1000

    return $Utime
}

function ConvertFrom-UTime() {
    Param(
        [decimal]$Utime
    )

    [DateTime]$DateTime = [System.DateTimeOffset]::FromUnixTimeMilliseconds(1000 * $Utime).LocalDateTime

    return $DateTime
}

function Get-Headers() {
    Param(
        [string]$ContentType = 'application/json',
        $ContentDisposition,
        [string]$filename,
        [switch]$AuthOnly
    )
    $config = Read-Config
    $Authorization = "Bearer {0}" -f $Config.APIKey
    #$headers = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
    $Headers = @{}
    
    $Headers.Add("Authorization", $Authorization)
    if ($AuthOnly) { return $Headers}

    if ($ContentType) {
        $Headers.Add('Content-Type', $ContentType)
    } else {
        $Header.Add('application/json')
    }

    if ($ContentDisposition) {
        if ($filename) {
            $file = Get-Item $filename            
            $ContentDisposition += "; filename=`"{0}`"" -f $file.Name
            $Headers.Add('Content-Disposition', $ContentDisposition)        
            #$Headers.Add("Content-Length", $size)
        } else {
            $Headers.Add('Content-Disposition', $ContentDisposition)        
        }
    }

    return $Headers
}