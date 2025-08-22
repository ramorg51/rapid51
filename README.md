rapid51
=======
def collectAppPools = { remoteHost, windowsCredentials ->
    // remoteHost: comma-separated string of servers
    // windowsCredentials: Jenkins credential ID (username/password)
    
    withCredentials([usernamePassword(credentialsId: windowsCredentials, usernameVariable: 'USR', passwordVariable: 'PSW')]) {
        def psScript = """
param([string]\$ServerList)

\$cred = New-Object System.Management.Automation.PSCredential (
    '\$env:USR', ('\$env:PSW' | ConvertTo-SecureString -AsPlainText -Force)
)

\$servers = \$ServerList -split "," | ForEach-Object { \$_ .Trim() }

\$appPoolsResults = @()
\$virtualDirResults = @()

foreach (\$server in \$servers) {
    try {
        Write-Host "Fetching data from \$server..." -ForegroundColor Cyan

        # Application Pools
        \$appPools = Invoke-Command -ComputerName \$server -Credential \$cred -ScriptBlock {
            Import-Module WebAdministration
            Get-ChildItem IIS:\\AppPools | Select-Object Name, State
        }

        foreach (\$pool in \$appPools) {
            \$appPoolsResults += [PSCustomObject]@{
                ServerName = \$server
                AppPool    = \$pool.Name
                State      = \$pool.State
            }
        }

        # Virtual Directories
        \$virtualDirs = Invoke-Command -ComputerName \$server -Credential \$cred -ScriptBlock {
            Import-Module WebAdministration
            Get-WebVirtualDirectory | Select-Object Name, PhysicalPath
        }

        foreach (\$vd in \$virtualDirs) {
            \$virtualDirResults += [PSCustomObject]@{
                ServerName     = \$server
                VirtualDirName = \$vd.Name
                PhysicalPath   = \$vd.PhysicalPath
            }
        }
    }
    catch {
        Write-Warning "Failed to fetch data from \$server. Error: \$_"
    }
}

# Export results
\$outputAppPools = "C:\\Temp\\IIS_AppPools.xlsx"
\$outputVirtualDirs = "C:\\Temp\\IIS_VirtualDirectories.xlsx"

if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
    Install-Module ImportExcel -Force -Scope CurrentUser
}

\$appPoolsResults | Export-Excel -Path \$outputAppPools -AutoSize -Title "IIS App Pools"
\$virtualDirResults | Export-Excel -Path \$outputVirtualDirs -AutoSize -Title "IIS Virtual Directories"

Write-Host "Data collection complete."
Write-Host "AppPools file: \$outputAppPools"
Write-Host "VirtualDirs file: \$outputVirtualDirs"
"""

        // Run PowerShell script on Jenkins slave
        powershell(script: psScript, returnStatus: true, args: ["-ServerList", remoteHost])
    }
}