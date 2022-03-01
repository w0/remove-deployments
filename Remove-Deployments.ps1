
[CmdletBinding()]
param (
    [Parameter(Mandatory,HelpMessage='Enter the Collection name. WildCards Accepted (*)')]
    [string]
    $CollectionName,

    [Parameter(Mandatory,HelpMessage='ConfigMgr Site Server address')]
    [string]
    $SiteServer,

    [Parameter(Mandatory,HelpMessage='ConfigMgr Site Code')]
    [string]
    $SiteCode
)

begin {

    if (-not (Get-Module -Name 'ConfigurationManager')) {
        Write-Progress 'Importing Configuration Manager PSModule'
        Import-Module (Join-Path (Split-Path $env:SMS_ADMIN_UI_PATH -Parent) 'ConfigurationManager.psd1')
    }

    if (-not (Get-PSDrive -Name $SiteCode -ErrorAction SilentlyContinue)) {
        Write-Progress 'Creating PS Site Drive'
        New-PSDrive -Name $SiteCode -PSProvider 'CMSite' -Root $SiteServer -Description 'ConfigMgr Site'
    }
    
    Push-Location -Path ('{0}:' -f $SiteCode)

    Write-Progress 'Getting Collection Information.'
    $Collections = Get-CMCollection -Name $CollectionName

}

process {

    foreach ($Collection in $Collections) {
        Write-Progress -Id 0 -Activity ('Checking: {0}' -f $Collection.Name)
        $Deployments = Get-CMApplicationDeployment -CollectionId $Collection.CollectionID

        if ($Deployments.Count -le 5) { continue }

        $Deployments = $Deployments | Sort-Object -Property CreationTime | Select-Object -First $($Deployments.Count - 5)
        
        $Deployments | ForEach-Object {
            Write-Progress -Id 1 -Parent 0 -Activity ('Removing: {0}' -f $_.AssignmentName)
            Remove-CMApplicationDeployment -InputObject $_ -Confirm:$false -Force:$true
        }
    }
}

end {
    Pop-Location
}
