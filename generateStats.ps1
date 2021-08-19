
function Show-ProgressV3 {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true, Position=0, ValueFromPipeline=$true)]
        [PSObject[]]$InputObject,
        [string]$Activity = "Processing items"
    )

        [int]$TotItems = $Input.Count
        [int]$Count = 0

        $Input|foreach {
            $_
            $Count++
            [int]$percentComplete = ($Count/$TotItems* 100)
            Write-Progress -Activity $Activity -PercentComplete $percentComplete -Status ("Working - " + $percentComplete + "%") -CurrentOperation (""+$Count+"/"+$TotItems+" - "+$_.Name)
        }
}

$count = 0
$Visio = New-Object -ComObject Visio.Application

(Get-ChildItem -Recurse -Include @("*.vss", "*.vssx")) | Show-ProgressV3 | Foreach-Object {
    $doc = $Visio.Documents.OpenEx($_.FullName, 192)
    $count += $doc.Masters.Count
    $doc.close()

    Start-Sleep 1
}

$Visio.Quit()

[System.Runtime.Interopservices.Marshal]::ReleaseComObject($Visio) | Out-Null

Write-Host "Old template files:" (Get-ChildItem -Recurse -Include "*.vss").Count
Write-Host "New template files:" (Get-ChildItem -Recurse -Include "*.vssx").Count
Write-Host "Total visio stencils:" $count
