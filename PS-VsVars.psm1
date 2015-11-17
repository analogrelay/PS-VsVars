$VisualStudios = @()
Export-ModuleMember -Variable VisualStudios

$EditionNodes = @(
    "VisualStudio",
    "VCExpress",
    "VPDExpress",
    "VSWinExpress",
    "VWDExpress",
    "WDExpress"
)

# I assume there will be a 15.0 :), who knows for sure? :P
$CtpVersions = @(
    New-Object System.Version "15.0"
)

$SearchRoot = "Software\Microsoft"
if($env:PROCESSOR_ARCHITECTURE -eq "AMD64") {
    $SearchRoot = "Software\Wow6432Node\Microsoft"
}

$EditionNodes | ForEach-Object {
    $edition = $_;
    $root = "HKLM:\$SearchRoot\$edition"
    if(Test-Path $root) {
        dir "$root" | Where-Object {
            ($_.Name -match "\d+\.\d+") -and
            (![String]::IsNullOrEmpty((Get-ItemProperty "$root\$($_.PSChildName)").InstallDir)) 
        } | ForEach-Object {
            $regPath = "$root\$($_.PSChildName)"
            
            # Gather VS data
            $installDir = (Get-ItemProperty $regPath).InstallDir

            $vsVars = $null;
            if(Test-Path "$installDir\..\..\VC\vcvarsall.bat") {
                $vsVars = Convert-Path "$installDir\..\..\VC\vcvarsall.bat"
            }
            $devenv = (Get-ItemProperty "$regPath\Setup\VS").EnvironmentPath;
            
            # Make a VSInfo object
            $ver = New-Object System.Version $_.PSChildName;
            $vsInfo = [PSCustomObject]@{
                "Edition" = $edition;
                "Version" = $ver;
                "RegistryRoot" = $_;
                "InstallDir" = $installDir;
                "VsVarsPath" = $vsVars;
                "DevEnv" = $devenv;
                "Prerelease" = ($CtpVersions -contains $ver)
            }

            # Add it to the dictionary
            $VisualStudios += $vsInfo
        }
    }
}

$DefaultVisualStudio = $VisualStudios | Where { $_.Edition -eq "VisualStudio" } | sort Version -desc | select -first 1
Export-ModuleMember -Variable DefaultVisualStudio

function Import-VsVars {
    param(
        [Parameter(Mandatory=$false)][string]$Version = $null,
        [Parameter(Mandatory=$false)][string]$Edition = "VisualStudio",
        [Parameter(Mandatory=$false)][string]$VsVarsPath = $null,
        [Parameter(Mandatory=$false)][string]$Architecture = $env:PROCESSOR_ARCHITECTURE,
        [Parameter(Mandatory=$false)][switch]$PrereleaseAllowed
    )
    
    if([String]::IsNullOrEmpty($VsVarsPath)) {
        Write-Debug "Finding vcvarsall.bat automatically..."

        # Find all versions of the specified edition
        $vers = $VisualStudios | Where { $_.Edition -eq $Edition }

        $Vs = Get-VisualStudio -Version $Version -Edition $Edition -PrereleaseAllowed:$PrereleaseAllowed;
        
        if(!$Vs) {
            throw "No $Edition Environments found"
        } else {
            Write-Debug "Found VS $($Vs.Version) in $($Vs.InstallDir)"
            $VsVarsPath = $Vs.VsVarsPath
        }
    }
    if($VsVarsPath -and (Test-Path $VsVarsPath)) {
        # Run the cmd script
        Write-Debug "Invoking: `"$VsVarsPath`" $Architecture"
        Invoke-CmdScript "$VsVarsPath" $Architecture
        "Imported Visual Studio $VsVersion Environment into current shell"
    } else {
        throw "Could not find VsVars batch file at: $VsVarsPath!"
    }
}
Export-ModuleMember -Function Import-VsVars

function Get-VisualStudio {
    param(
        [Parameter(Mandatory=$false, Position=1)][string]$Version,
        [Parameter(Mandatory=$false, Position=1)][string]$Edition = "VisualStudio",
        [Parameter(Mandatory=$false)][switch]$PrereleaseAllowed)
    # Find all versions of the specified edition
    $vers = $VisualStudios | Where { $_.Edition -eq $Edition }

    $Vs = $null;
    if($Version) {
        $Vs = $vers | where { $_.Version -eq [System.Version]$Version } | select -first 1
    } else {
        $Vs = $vers | where { $PrereleaseAllowed -or !($_.Prerelease) } | sort Version -desc | select -first 1
    }

    if(!$Vs) {
        if($Version) {
            throw "Could not find $Edition $Version!"
        } else {
            throw "Could not find any $Edition version!"
        }
    }
    $Vs
}
Export-ModuleMember -Function Get-VisualStudio

function Invoke-VisualStudio {
    param(
        [Parameter(Mandatory=$false, Position=0)][string]$Solution,
        [Parameter(Mandatory=$false, Position=1)][string]$Version,
        [Parameter(Mandatory=$false, Position=2)][string]$Edition,
        [Parameter(Mandatory=$false)][switch]$Elevated,
        [Parameter(Mandatory=$false)][switch]$PrereleaseAllowed,
        [Parameter(Mandatory=$false)][switch]$WhatIf)

    # Load defaults from ".vslaunch" file
    if(Test-Path ".vslaunch") {
        $config = ConvertFrom-Json (cat -Raw ".vslaunch")
        if($config) {
            $Edition = if($Edition) { $Edition } else { $config.edition }
            $Version = if($Version) { $Version } else { $config.version }
            $Solution = if($Solution) { $Solution } else { $config.solution }
            $Elevated = if($Elevated) { $Elevated } else { $config.elevated }
        }
    }
    if(!$Edition) {
        $Edition = "VisualStudio";
    }
    Write-Debug "Launching: Edition=$Edition, Version=$Version, Solution=$Solution, Elevated=$Elevated"

    if([String]::IsNullOrEmpty($Solution)) {
        $Solution = "*.sln"
    }
    elseif(!$Solution.EndsWith(".sln")) {
        $Solution = "*" + $Solution + "*.sln";
    }

    $devenvargs = @();
    if(!(Test-Path $Solution)) {
        Write-Host "Could not find any solutions. Launching VS without opening a solution."
    }
    else {
        $slns = @(dir $Solution)
        if($slns.Length -gt 1) {
            $names = [String]::Join(",", @($slns | foreach { $_.Name }))
            throw "Ambiguous matches for $($Solution): $names";
        }
        $devenvargs = @($slns[0])
    }

    $Vs = Get-VisualStudio -Version $Version -Edition $Edition -PrereleaseAllowed:$PrereleaseAllowed
    $devenv = $Vs.DevEnv

    if($devenv) {
        if($WhatIf) {
            Write-Host "Launching '$devenv $devenvargs'"
        }
        else {
            if($Elevated) {
                elevate $devenv @devenvargs
            }
            else {
                &$devenv @devenvargs
            }
        }
    } else {
        throw "Could not find desired Visual Studio Edition/Version: $Edition v$Version"
    }
}
Set-Alias -Name vs -Value Invoke-VisualStudio
Export-ModuleMember -Function Invoke-VisualStudio -Alias vs
