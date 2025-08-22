# Copyright (C) AVI-SPL. All Rights Reserved.
# The intellectual and technical concepts contained herein are proprietary to AVI-SPL, Inc. and subject to AVI-SPL's standard software license agreement. 
# These materials may not be copied, reproduced, distributed or disclosed, in whole or in part, in any way without the written permission of an authorized representative of AVI-SPL.
# All references to AVI-SPL, Inc. shall also be references to AVI-SPL, Inc. affiliates. 

param (
    [Parameter( HelpMessage = "Create the zip without uploading a Gitlab release.")]
    [Switch]$LocalOnly,
    [Parameter( HelpMessage = "Pause before zipping the release folder and uploading it to GitLab.")]
    [Switch]$PauseBeforeZipping,
    [Parameter( HelpMessage = "Path to the build properties json file.")]
    [String]$PropertiesFile = '.\build.properties',
    [Parameter( HelpMessage = "Comma separated list of Groups to include.")]
    [String[]]$IncludeGroups,
    [Parameter( HelpMessage = "Comma separated list of Groups to exclude.")]
    [String[]]$ExcludeGroups,
    [Parameter( HelpMessage = "Exclude all files that have a Group definition")]
    [Switch]$ExcludeAllGroups,
    [Parameter( HelpMessage = "Override the configured group release suffix definition.")]
    [String]$GroupReleaseSuffix,
    [Parameter( HelpMessage = "Scan directories for files and update the existing build.properties.")]
    [Switch]$UpdateFiles,
    [Parameter( HelpMessage = "Override access_token within ~/.gitlab_api for this execution.")]
    [String]$OverrideApiToken,
    [Parameter(HelpMessage = "Skip VerifySimplWindowsRefs checks on .smw files.")]
    [Switch]$SkipVerifySmw,
    [Parameter(HelpMessage = "Skip pre-script execution.")]
    [Switch]$SkipPre,
    [Parameter(HelpMessage = "Skip post-script execution.")]
    [Switch]$SkipPost,
    [Parameter( HelpMessage = "Whether to show additional debugging information.", Mandatory = $false)]
    [Switch]$DebugEnable,
    [Parameter(HelpMessage = "Whether to disable milestone association feature.", Mandatory = $false)]
    [Switch]$MilestonePromptDisable
)

# vars
$GitlabV4ApiUrl = "https://gitlab.avispl.com/api/v4"
$GenerateReleaseProjectId = "545"
$VerifySimplWindowsRefsProjectId = "8602"
$TagSemverRegexPattern = "^v(0|[1-9]\d*)\.(0|[1-9]\d*)\.(0|[1-9]\d*)(?:-((?:0|[1-9]\d*|\d*[a-zA-Z-][0-9a-zA-Z-]*)(?:\.(?:0|[1-9]\d*|\d*[a-zA-Z-][0-9a-zA-Z-]*))*))?(?:\+([0-9a-zA-Z-]+(?:\.[0-9a-zA-Z-]+)*))?$"
$ScriptVersion = @{
    Major = 1
    Minor = 10
    Patch = 0
    Pre   = @()
}
$BuildProperties = $null;
$DefaultStagingPath = '.releases'
$DefaultReleaseName = 'PROJECTi-YY-NUMBER - ClientNameOrInitials LocationIfDesired - RoomNameOrRoomTypePleaseUseSpaces'
$DefaultMilestoneNames = @("Job Number and Description", "Original Job Number and Description", "Case Number and Description")
$DefaultReleaseNameHistory = $DefaultReleaseName, '{ProjectNumber}-{Client}-{Room Name or Type}'
$ErrorActionPreference = "Stop" # stop after any error
$MaxInstructionFileSize = 20000  # Instruction files larger than this will trigger a warning to the user.
$MaxUploadFileSizeMb=[int]500

$CrestronProgramArtifacts = @('*.lpz', '*.cpz', '*.sig', '*.spz')
$ExtronProgramArtifacts = @('*.gcplus', '*.gcpro')
$CrestronUIArtifacts = @('*.vtz', '*.ch5z', '*.Core3z', '*.c3prj')
$CrestronProgramSourceArtifacts = @('*_archive.zip')
$CrestronUISourceArtifacts = @('*.vta')
$CrestronExcludes = @('CrestronXPanel installer.*')  # The latest files are hosted on Crestron's website.
$ExtronUIArtifacts = @('*.gdl')
$QscProgramArtifacts = @('*.qsys')
$QscUIArtifacts = @('*.uci')
$BiampDesignerArtifacts = @('*.pdprj')
$ConfigFileTypes = @('.json')
$ConfigFolderExclusions = @('Configs')
$DocumentationFileTypes = @('*.doc*', '*.pdf', '*.md', '*.xls*')
$DocumentationPath = '.\Documentation\' # Define to prevent script from grabbing all doc file types in any folder.
$GitRemotePattern = "git@(?'host'.*):(?'remote'.*)"
$TokenExpiryDays = 15 # This value is set based on data gathered from GitLab

function Set-Exit($code) {
    exit($code)
}

function Write-Log() {
    param (
        [string]
        $Message,
        [Parameter()]
        [string]
        $LogLevel
    )

    switch -Exact ($LogLevel) {
        'Warn' { 
            $foregroundColor = "Yellow"
        }
        'Error' {
            $foregroundColor = "Red"
        }
        'Notice' {
            $foregroundColor = "Green"
        }
        Default {
            $foregroundColor = "Cyan"
        }
    } 
    Microsoft.PowerShell.Utility\Write-Host $Message -ForegroundColor $foregroundColor
}

function Remove-DirIfExists($path) {
    if (Microsoft.PowerShell.Management\Test-Path -LiteralPath $path) {
        try {
            Microsoft.PowerShell.Management\Remove-Item -recurse -force $path -erroraction stop
            return $true
        }
        catch {
            return $false
        }
    }
    else {
        return $true
    }
}

function Remove-FileIfExists($path) {
    if (Microsoft.PowerShell.Management\Test-Path -LiteralPath $path) {
        try {
            Microsoft.PowerShell.Management\Remove-Item $path -erroraction stop
            return $true
        }
        catch {
            return $false;
        }
    }
    else {
        return $true
    }
}

function Confirm-YesOrNo($question) {
    try {
        do {
            Write-Log "${question}? [y/n]" -LogLevel Warn
            $c = Microsoft.PowerShell.Utility\Read-Host
            if ($c -like 'n*') {
                throw New-GenerateReleaseException 100 "User chose to terminate the release."
            }
        } until($c -like 'y*')
    }
    catch [GenerateReleaseException] {
        throw $_
    } catch {
        Write-Log "Confirm-YesOrNo: Failed to execute Read-Host. Assuming automated environment..."
    }
}

# Check for access token
function Import-APIProperties($ApiPropertiesPath) {
    Write-Log "Looking for API properties at $ApiPropertiesPath"
    if (Microsoft.PowerShell.Management\Test-Path -LiteralPath "$ApiPropertiesPath") {
        Write-Log "Found API properties at $ApiPropertiesPath"
        $contents = Microsoft.PowerShell.Management\Get-Content $ApiPropertiesPath
        $propertiesJson = $contents -Join "`n" | Microsoft.PowerShell.Utility\ConvertFrom-Json
        try {
            $tokenInfo = Get-TokenInfo "$($propertiesJson.url)" "$($propertiesJson.access_token)"
        } catch {
            throw New-GenerateReleaseException 38 "Retrieving access_token information has failed. The token within $ApiPropertiesPath is invalid or has expired." "Generate a new token with API access checked from https://gitlab.avispl.com/-/profile/personal_access_tokens"
        }

        $now=$(Microsoft.PowerShell.Utility\Get-Date)
        $tokenExpiry=$(Microsoft.PowerShell.Utility\Get-Date $tokenInfo.expires_at)
        $timeSpan = Microsoft.PowerShell.Utility\New-TimeSpan -Start $now -End $tokenExpiry
        if ($timeSpan.Days -gt $TokenExpiryDays) {
            return $propertiesJson
        }

        Write-Log "The access token within $ApiPropertiesPath is going to expire within $TokenExpiryDays days and will be rotated." -LogLevel Warn
        try {
            $rotateResponse = Rotate-Token "$($propertiesJson.url)" "$($propertiesJson.access_token)"
        } catch {
            throw New-GenerateReleaseException 39 "Token rotation failed. Generate a new token with API access checked from https://gitlab.avispl.com/-/profile/personal_access_tokens and paste it into $ApiPropertiesPath"
        }

        $propertiesJson.access_token = $rotateResponse.token
        Microsoft.PowerShell.Management\Set-Content $ApiPropertiesPath $($propertiesJson | Microsoft.PowerShell.Utility\ConvertTo-Json -Depth 5)
        Write-Log "Access token rotated and saved within $ApiPropertiesPath. The new token will expire on $($rotateResponse.expires_at)" -LogLevel Notice
        return $propertiesJson
    }
    else {
        Write-Log "GitLab API properties not found." -LogLevel Warn
        return
    }
}

# Generate a new default API properties file
function New-APIProperties($ApiPropertiesPath, $ApiURL) {
    Write-Log "If you do not have a GitLab Personal Access Token, please generate a new token with API access checked from https://gitlab.avispl.com/-/profile/personal_access_tokens" 
    $props = @{ 
        'access_token' = Microsoft.PowerShell.Utility\Read-Host "Enter your Gitlab Personal Access Token"
        'url'          = $ApiURL
    }
    $save = Microsoft.PowerShell.Utility\Read-Host "Save this token for later use by this script? [y/n]"
    switch ($save) {
        "y" {
            Microsoft.PowerShell.Management\New-Item -ItemType File -Path $ApiPropertiesPath -Value ($props | Microsoft.PowerShell.Utility\ConvertTo-Json)
            Write-Log "Your Gitlab API connection info has been saved to $ApiPropertiesPath" -LogLevel Notice
        }
    }
    return $props
}

# Loads the build.properties file if it exists
function Import-BuildProperties ($FileName) {
    try {
        Write-Log "Loading build properties at $FileName"

        $properties = (Microsoft.PowerShell.Management\Get-Content $FileName) -join "`n" | Microsoft.PowerShell.Utility\ConvertFrom-Json

        if ($($properties.VersionNumber -match "^[a-zA-Z0-9\.-]+$") -eq $false) {
            throw New-GenerateReleaseException 2 "Invalid characters in VersionNumber. Only alpha, numeric, dot, and dash are allowed."
        }

        if ($GroupReleaseSuffix -and $GroupReleaseSuffix -match "^[a-zA-Z0-9\.-]+$" -eq $false) {
            throw New-GenerateReleaseException 2 "Invalid characters in GroupReleaseSuffix. Only alpha, numeric, dot, and dash are allowed."
        }

        if ($null -eq $($properties.StagingPath)) {
            Write-Log "The $PropertiesFile file does not contain a 'StagingPath' property. Defaulting to '$DefaultStagingPath'. Define the 'StagingPath' in $PropertiesFile to change the location." -LogLevel Warn
            Microsoft.PowerShell.Utility\Add-Member -InputObject $properties -MemberType NoteProperty -Name "StagingPath" -Value $DefaultStagingPath
        }

        if ($PSVersionTable.PSVersion.Major -eq 7) {
            $invalidReleaseNameChars = @("\", '/', "|", "␦")
        } else {
            $invalidReleaseNameChars = @("\", '/', "|")
        }
        if ($($properties.ReleaseName).IndexOfAny($invalidReleaseNameChars) -ne -1) {
            throw New-GenerateReleaseException 36 "The 'ReleaseName' property contains characters that are invalid with GenerateRelease. Modify 'ReleaseName' to not contain any of these characters: $($invalidReleaseNameChars -join ", ")"
        }

        if ($($properties.ReleaseName) -in $DefaultReleaseNameHistory) {
            throw New-GenerateReleaseException 3 "The 'ReleaseName' property is still the default value '$($properties.ReleaseName)'. Fill in the appropriate values in the $PropertiesFile file for your release."
        }

        foreach ($file in $properties.Copy){
            if ($null -eq $file){
                throw New-GenerateReleaseException 37 "The 'Copy' property contains null entries. Ensure that the 'Copy' property is a well-formed array."
            }
            if (!$file.FileName) {
                throw New-GenerateReleaseException 37 "The 'Copy' property contains entries with no 'FileName' property. Ensure that all entries have a valid 'FileName' property."
            }
        }

        return $properties
    }
    catch [GenerateReleaseException] {
        throw $_
    }
    catch {
        throw New-GenerateReleaseException 1 "$($PropertiesFile) is not a valid JSON object. Check the formatting of the file, or regenerate it."
    }
}

function Update-BuildProperties($Path, $ExistingProperties) {
    $existingFiles = $ExistingProperties.Copy
    $newFiles = $(Get-CopyProperty $Path)
    $fileDiff = @()

    foreach ($newFile in $newFiles) {
        if (!$($existingFiles.FileName -contains $newFile.FileName)) {
            $existingFiles += $newFile
            $fileDiff += $newFile
        }
    }

    Write-Log "Updated $PropertiesFile with $($fileDiff.Length) new files" -LogLevel Warn
    foreach ($diff in $fileDiff) {
        Write-Log "  - $($diff.FileName)" -LogLevel Warn
    }
    $ExistingProperties.Copy = $existingFiles
    return $ExistingProperties
}

# Generates a new build.properties
function New-BuildProperties($Path) {
    $BuildProperties = [ordered]@{};
    $BuildProperties['VersionNumber'] = "v1.0.00"; # Provide example starting release number so initial releases start at 1.0.00 instead of 0.0.1. 
    $BuildProperties['ReleaseName'] = $DefaultReleaseName;
    $BuildProperties['StagingPath'] = $DefaultStagingPath;
    $BuildProperties['ReleaseNotes'] = "";
    $BuildProperties['InstructionsFile'] = ""; # Optional file used when populating the release notes
    $BuildProperties['Copy'] = $(Get-CopyProperty $Path)
    return $BuildProperties
}

function Get-CopyProperty($Path) {
    $Files = @();
    $artifacts = Find-BuildArtifacts $Path
    foreach ($artifact in $artifacts) {
        $Files += [ordered]@{
            'FileName' = "$(Resolve-RelativePath $artifact.FullName)";
            'Rename'   = "false"
        };
    }
    $sourceArtifacts = Find-CrestronSourceArtifacts $Path
    foreach ($artifact in $sourceArtifacts) {
        $Files += [ordered]@{
            'FileName'        = "$(Resolve-RelativePath $artifact.FullName)";
            'Rename'          = "false";
            'OutputDirectory' = "$(Resolve-RelativePath $artifact.FullName | Microsoft.PowerShell.Management\Split-Path)\Source (Programmer Use ONLY)\";
        };
    }
    $configFiles = Find-ConfigFiles $Path
    foreach ($file in $configFiles) {
        $Files += [ordered]@{
            'FileName' = "$(Resolve-RelativePath $file.FullName)";
            'Rename'   = "false"
        };
    }

    $Files += [ordered]@{
        'FileName' = "$(Resolve-RelativePath 'LICENSE')";
        'Rename'   = "false"
    };

    return $Files
}

# Get built output files
function Find-BuildArtifacts($Path) {
    Write-Log "Finding build artifacts in '$Path'"

    $crestronArtifacts = Get-GenericArtifacts ($CrestronProgramArtifacts + $CrestronUIArtifacts) $Path
    # filter out Crestron Construct output
    $crestronArtifacts = $crestronArtifacts | Microsoft.PowerShell.Core\Where-Object { 
        return $_.FullName -notlike "*__TEMP*" 
    }
    Write-Log "Found $($crestronArtifacts.Count) Crestron artifact(s)"

    $crestronExecutableXpanels = Get-ExecutableXPanels $Path
    Write-Log "Added $($crestronExecutableXpanels.Count) Executable XPanel artifact(s)"

    $extronArtifacts = Get-ExtronArtifacts $Path
    Write-Log "Found $($extronArtifacts.Count) Extron artifact(s)"

    $qscArtifacts = Get-GenericArtifacts ($QscProgramArtifacts + $QscUIArtifacts) $Path
    $qscArtifacts = $qscArtifacts | Microsoft.PowerShell.Core\Where-Object { 
        return $_.FullName -notlike "*Configs\*"
    }
    Write-Log "Found $($qscArtifacts.Count) QSC artifact(s)"

    $amxArtifacts = Get-AmxArtifacts $Path
    Write-Log "Found $($amxArtifacts.Count) AMX artifact(s)"

    $biampArtifacts = Get-GenericArtifacts $BiampDesignerArtifacts $Path
    Write-Log "Found $($biampArtifacts.Count) Biamp artifact(s)"

    $documentArtifacts = Get-GenericArtifacts $DocumentationFileTypes $DocumentationPath
    Write-Log "Found $($documentArtifacts.Count) documentation artifact(s)"

    $allArtifacts = @($crestronArtifacts) + @($extronArtifacts) + @($qscArtifacts) + @($amxArtifacts) + @($biampArtifacts) + @($documentArtifacts) | Microsoft.PowerShell.Core\Where-Object { $_.FullName -notlike "*$DefaultStagingPath*" }
    $files = $allArtifacts | Microsoft.PowerShell.Core\Where-Object { $_.Directory }
    $directories = $allArtifacts | Microsoft.PowerShell.Core\Where-Object { !$_.Directory }
    $additionalFiles = $directories | Microsoft.PowerShell.Core\ForEach-Object { 
        if ($null -ne $_.FullName -and $_.FullName.EndsWith(".c3prj")) {
            # issue-128 -> do not recurse into .c3prj dirs; keep them verbatim.
            $_
        } else {
            Microsoft.PowerShell.Management\Get-ChildItem -path $_ -recurse -Exclude $CrestronExcludes -Attributes !Directory 
        }
    }

    $finalFiles = @($files) + @($additionalFiles) + @($crestronExecutableXpanels)
    if (!$finalFiles) {
        Write-Log "No artifact files were found in this repository. Manually edit the $PropertiesFile to add your files." -LogLevel Warn
    }

    return $finalFiles;
}

# These get stashed in a separate "Source" folder
function Find-CrestronSourceArtifacts($Path) {
    Write-Log "Finding Crestron source build artifacts in '$Path'"

    $crestronArtifacts = Get-GenericArtifacts ($CrestronProgramSourceArtifacts + $CrestronUISourceArtifacts) $Path
    $constructSolutions = Get-GenericArtifacts "*.csln" $Path  
    $allArtifacts = @($crestronArtifacts) | Microsoft.PowerShell.Core\Where-Object { $_.FullName -notlike "*$DefaultStagingPath*" }
    $files = $allArtifacts | Microsoft.PowerShell.Core\Where-Object { $_.Directory }
    $directories = $allArtifacts | Microsoft.PowerShell.Core\Where-Object { !$_.Directory }
    $additionalFiles = $directories | Microsoft.PowerShell.Core\ForEach-Object {
        Microsoft.PowerShell.Management\Get-ChildItem -path $_ -recurse -Attributes !Directory 
    }
    $fileArray = @($files)   
    foreach ($solution in $constructSolutions) {
        $parent = $solution.Directory
        $fileArray += $parent
    } 
    
    $allFiles = @($fileArray) + @($additionalFiles)
    $finalFiles = Remove-AllFilesInGitignore $allFiles

    Write-Log "Found $($finalFiles.Count) Crestron source artifact(s)"
    
    if (!$finalFiles) {
        Write-Log "No crestron source (*_archive.zip, *.csln, *.vta) artifact files were found in this repository. Manually edit the $PropertiesFile to add your files." -LogLevel Warn
    }

    return $finalFiles;
}

# Use this function when special file/directory handling is not required.
# Calling example: Get-GenericArtifacts ($ListOne + $ListTwo + $ListThree + ...) $Subdirectory
function Get-GenericArtifacts($ArtifactList, $SearchPath) {
    if (!(Microsoft.PowerShell.Management\Test-Path $SearchPath)) {
        Write-Log "$SearchPath does not exist. Skipping artifact detection..." -LogLevel Warn
        return @()
    }
    $artifacts = Microsoft.PowerShell.Management\Get-ChildItem -Path $SearchPath -Include $ArtifactList -Recurse
    return @($artifacts)
}

function Get-ExtronArtifacts($Path) {
    $extronIncludes = $ExtronProgramArtifacts + $ExtronUIArtifacts
    $extronArtifacts = Microsoft.PowerShell.Management\Get-ChildItem -Path $Path -Include $extronIncludes -Recurse |
    Microsoft.PowerShell.Core\Where-Object { $_.FullName -notlike "*_bak.g*" }
    return @($extronArtifacts)
}

function Get-AmxArtifacts($Path) {
    $amxDirectories = Microsoft.PowerShell.Management\Get-ChildItem -Path $Path -Filter "Code" -Recurse
    $amxApwArtifacts = $amxDirectories | Microsoft.PowerShell.Management\Get-ChildItem -Filter "*.apw"
    if (!$amxApwArtifacts) {
        return @()
    }

    Write-Log "Found $(@($amxApwArtifacts).Count) .apw file(s). Finding AMX artifacts..."

    $amxArtifactPaths = @()
    $masterFiles = @()
    foreach ($apw in $amxApwArtifacts) {
        $apwParent = Microsoft.PowerShell.Management\Split-Path -Path $apw.FullName -Parent
        $apwXml = [xml]$(Microsoft.PowerShell.Management\Get-Content -Path $apw.FullName)
        foreach ($file in $($apwXml.Workspace.Project.System.File)) {
            if ($file.Type -eq "MasterSrc") {
                $masterFiles += ("$apwParent\$($file.FilePathName)")
            } elseif ($file.FilePathName.EndsWith(".ftl")) {
                if (!($file.FilePathName -like "*Configs*")){
                    $amxArtifactPaths += ("$apwParent\$($file.FilePathName)")
                }
            } else {
                $amxArtifactPaths += ("$apwParent\$($file.FilePathName)")
            }
        }
    }

    $amxArtifactPaths = @($amxArtifactPaths | % { Resolve-RelativePath $_ })
    $amxArtifactPaths = @($amxArtifactPaths | Microsoft.PowerShell.Utility\Sort-Object -unique)
    $masterFiles = @($masterFiles | Microsoft.PowerShell.Utility\Sort-Object -unique)

    foreach ($file in $masterFiles) {
        $parent = Microsoft.PowerShell.Management\Split-Path -Path $file -Parent
        $fileName = Microsoft.PowerShell.Management\Split-Path -Path $file -Leaf
        $fileNameExt = $fileName.Split('.') | Microsoft.PowerShell.Utility\Select-Object -Last 1
        $fileNameNoExt = $fileName.Substring(0, $fileName.Length - $fileNameExt.Length - 1) # powershell 5 support
        foreach ($ext in "tkn", "src", "tko") {
            $amxArtifactPaths += "$parent\$fileNameNoExt.$ext"
        }
    }
    
    $amxFileObjs = @()
    foreach ($amxFilePath in $amxArtifactPaths) {
        try {
            $afi = Microsoft.PowerShell.Management\Get-Item $amxFilePath
            $amxFileObjs += $afi
        } catch {
            Write-Log "Expected $amxFilePath to exist, but it did not. Did you forget to compile the workspace?" -LogLevel Warn
        }
    }

    $ftlArtifacts = $amxFileObjs | % { if ($_.FullName.EndsWith(".ftl")) {$_} } 
    $otherFtlArtifacts = Get-ChildItem -path "$Path/Configs" -recurse -filter "*.ftl"
    $existingFtlArtifacts = @($ftlArtifacts) + @($otherFtlArtifacts)
    if (!$existingFtlArtifacts) {
        Write-Log "There is no .ftl file in this project. See the link below for a guide on how to generate one:" -LogLevel Warn
        Write-Log "https://gitlab.avispl.com/avispl/modules/tools/generate-release/-/wikis/AMX:-Netlinx-File-Transfer-List-generation" -LogLevel Warn
    }
    else {
        # check timestamp of apw vs ftl and warn about stale one.
        foreach ($ftl in $existingFtlArtifacts) {
            try {
                $ftlLwt = $(Microsoft.PowerShell.Management\Get-Item $ftl.FullName).LastWriteTime
                foreach ($apw in $amxApwArtifacts) {
                    $apwLwt = $(Microsoft.Powershell.Management\Get-Item $apw.FullName).LastWriteTime
                    if ($ftlLwt -lt $apwLwt) {
                        Write-Log "'$($ftl.Name)' was created before the last modification to '$($apw.Name)'. Consider re-exporting '$($ftl.Name)'" -LogLevel Warn
                        Write-Log ".ftl guide: https://gitlab.avispl.com/avispl/modules/tools/generate-release/-/wikis/AMX:-Netlinx-File-Transfer-List-generation" -LogLevel Warn
                    }
                }
            } catch {}
        }
    }

    return @($amxFileObjs)
}

function Get-ExecutableXPanels($SearchPath){
    if (!(Microsoft.PowerShell.Management\Test-Path $SearchPath)){
        Write-Log "$SearchPath does not exist. Skipping Executable XPanel detection..." -LogLevel Warn
        return @()
    }
    $underscoreConfigInis = Microsoft.PowerShell.Management\Get-ChildItem -Path $SearchPath -Filter "_config_ini_" -Recurse

    $executableXpanels = @()
    foreach ($file in $underscoreConfigInis) {
        $dir = $file.Directory
        Write-Log "Should '$(Microsoft.PowerShell.Management\Resolve-Path $dir -Relative)' be included as an Executable XPanel? [y/n]" -LogLevel Warn
        try {
            $c = Microsoft.PowerShell.Utility\Read-Host
        if ($c -like "y*") {
            $executableXpanels += $dir
        }
        } catch {
            Write-Log "Get-ExecutableXPanels: Failed to execute Read-Host. $dir not added to artifact listing."
        }

    }

    return @($executableXpanels)
}

function Find-ConfigFiles($Path) {
    Write-Log "Finding all files within '$Path/Configs'"
    $configs = Microsoft.PowerShell.Management\Get-ChildItem -Path "$Path/Configs" -exclude ".git*", ".cbh*" -Recurse -Attributes !Directory |
    Microsoft.PowerShell.Core\Where-Object { $_.FullName -notlike "*node_modules*" }
    Write-Log "Finding '$ConfigFileTypes' files in '$Path', excluding '$ConfigFolderExclusions' folders"
    $otherFiles = Microsoft.PowerShell.Management\Get-ChildItem -Path $Path -exclude $ConfigFolderExclusions |
    Microsoft.PowerShell.Management\Get-ChildItem -recurse -Attributes !Directory |
    Microsoft.PowerShell.Core\Where-Object { $ConfigFileTypes.Contains($_.Extension) `
            -and $_.FullName -notlike "*node_modules*" `
            -and $_.FullName -notlike "*$DefaultStagingPath*" `
            -and $_.FullName -notlike "*.vscode*" }
    $allConfigs = @($configs) + @($otherFiles)   
    $finalConfigs = Remove-AllFilesInGitignore $allConfigs
    if (!$finalConfigs) {
        Write-Log "No config files were found in this repository. Manually edit the $PropertiesFile to add your files." -LogLevel Warn
    }
    else {
        Write-Log "Found $($finalConfigs.Count) configuration file(s)"
    }

    return $finalConfigs;
}

# Build a valid file path from the running directory and the specified file path.
# This will help support both absolute and relative paths in the build.properties
function Build-Path($base, $path) {
    if (Microsoft.PowerShell.Management\Test-Path -LiteralPath $path) {
        return $path
    }

    $UpdatedPath = $("$base\$path").Replace("\\", "\").Replace("/", "\")
    if (Microsoft.PowerShell.Management\Test-Path -LiteralPath $UpdatedPath) {
        return $UpdatedPath
    }

    if (Microsoft.PowerShell.Management\Test-Path $UpdatedPath) {
        return $UpdatedPath
    }

    throw New-GenerateReleaseException 4 "Cannot find path $UpdatedPath because it does not exist. Is it within the git repository?"
}

# Generate release notes
function New-ReleaseNotes($BuildProperties, $AssetInfo, $projectLink) {
    $releaseNotes = ''

    # Add the instructions file if it is defined and we can find the file.
    if ($BuildProperties.InstructionsFile) {
        try {
            $file = $(Microsoft.PowerShell.Management\Get-Item $($BuildProperties.InstructionsFile))
            if ($file.length -gt $MaxInstructionFileSize) {
                Write-Log "`nThe InstructionsFile is over $MaxInstructionFileSize characters!  You're going to overwhelm them with so many instructions.  Please consider shortening your instructions.`n`nCTRL+C to quit." -LogLevel Warn
                Pause
            }
        }
        catch {
            throw New-GenerateReleaseException 4 "Cannot find path to InstructionsFile: $($BuildProperties.InstructionsFile) because it does not exist. Is it within the git repository?"
        }

        Write-Log "New-ReleaseNotes.  Reading in InstructionsFile $($BuildProperties.InstructionsFile)"
        $releaseNotes = (Get-Content -Path $BuildProperties.InstructionsFile) -Join "`n"
        $releaseNotes += "`n`n"
    }
    else {
        # Add the default instructions
        $releaseNotes = "----`n`n"
        $releaseNotes += "# Instructions - START HERE`n`n"
        $releaseNotes += " - Download the **_package (.zip)_** linked above that _includes the version number in the name_.`n"
    }

    # add notes here from BuildProperties.ReleaseNotes?
    if ($BuildProperties.ReleaseNotes) {
        $releaseNotes += "`n## Version Notes`n`n"
        $releaseNotes += $BuildProperties.ReleaseNotes
        $releaseNotes += "`n`n"
    }

    $releaseNotes += "##### GitLab Project Link`n`n"
    $releaseNotes += "[$projectLink]($projectLink)`n`n`n"

    $releaseNotes += "##### Files contained in this automatic release:`n`n"
    foreach ($item in $BuildProperties.Copy) {
        if ($item.Included) {
            $filePath = $item.FileName
            $fileName = [IO.Path]::GetFileName($filePath)
            $releaseNotes += "> $filePath  "
            if ($AssetInfo.ContainsKey($fileName)) {
                $releaseNotes += "`n"
                foreach ($key in $AssetInfo[$fileName].Keys) {
                    $releaseNotes += ">> **${key}**: $($AssetInfo[$fileName][$key])  `n"
                }
                $releaseNotes += ">`n"
            }
            else {
                $releaseNotes += "  `n"
            }
        }
    }

    if ($SkipVerifySmw){
        $releaseNotes += "`nThis release was generated by bypassing the verification of all source code existing in the repository."
    }

    return $releaseNotes
}

# Get git remote from .git folder
function Get-GitRemoteUrl  {
    $remote =  (git remote get-url --push origin) | Microsoft.PowerShell.Utility\Select-String -Pattern $GitRemotePattern
    if ($remote.Matches) {
        return $remote.Matches[0].Groups['remote'].Value
    }

    return ""
}

# Get GitLab project info
function Build-GitLabProjectUrl($gitRemote) {
    $dotGit=".git"
    $lastIndexDotGit = $gitRemote.LastIndexOf($dotGit)
    if ($lastIndexDotGit -gt 0) {
        $gitRemote = $gitRemote.Remove($lastIndexDotGit, $dotGit.Length)
    }

    return $gitRemote
}

function Get-GitLabReleaseInfo($ProjectURL, $Token) {
    $requestParams = @{
        Method = 'Get'
        Uri = "$ProjectURL/releases"
        Headers = @{
            'Private-Token' = $Token
            'Content-Type' = 'application/json'
        }
    }
    
    return Microsoft.PowerShell.Utility\Invoke-RestMethod @requestParams
}

function Get-Milestones($ProjectURL, $Token) {
    $requestParams = @{
        Method = 'Get'
        Uri = "$ProjectURL/milestones"
        Headers = @{
            'Private-Token' = $Token
            'Content-Type' = 'application/json'
        }
    }

    if ($PSVersionTable.PSVersion.Major -ge 7){
        return Microsoft.PowerShell.Utility\Invoke-RestMethod @requestParams -FollowRelLink
    } else {
        return Microsoft.PowerShell.Utility\Invoke-RestMethod @requestParams
    }
}

function Update-GitLabMilestoneTitle($ProjectURL, $Token, $MilestoneId, $NewTitle) {
    $payload = @{
        title = "$NewTitle"
    }
    $requestParams = @{
        Method = 'Put'
        Uri = "$ProjectURL/milestones/$MilestoneId"
        Headers = @{
            'Private-Token' = $Token
            'Content-Type' = 'application/json'
        }
        Body = $payload | Microsoft.PowerShell.Utility\ConvertTo-Json 
    }

    return Microsoft.PowerShell.Utility\Invoke-RestMethod @requestParams
}

function Read-MilestoneTitle($Prompt, $ExistingMilestoneTitles) {
    Write-Log $Prompt -LogLevel Warn
    do {
        $title = $(Microsoft.PowerShell.Utility\Read-Host)
        if (!$title) {
            Write-Log "Empty milestone name provided. Provide a title that includes the job or case number and a description:" -LogLevel Error
        }
        if ($DefaultMilestoneNames -contains $title) {
            Write-Log "The new name is still a default name. Provide a title that includes the job or case number and a description:" -LogLevel Warn
        } elseif ($ExistingMilestoneTitles -contains $title) {
            Write-Log "There is already a milestone with the title '$title'. Provide a different title that includes the job or case number and a description:" -LogLevel Warn
        }
    } until ($title -and !($DefaultMilestoneNames -contains $title) -and !($ExistingMilestoneTitles -contains $title))

    return $title
}

function New-GitLabMilestone($ProjectURL, $Token, $Title) {
    $payload = @{
        title = $Title
    }
    $requestParams = @{
        Method = 'Post'
        Uri = "$ProjectURL/milestones"
        Headers = @{
            'Private-Token' = $Token
            'Content-Type' = 'application/json'
        }
        Body = $payload | Microsoft.PowerShell.Utility\ConvertTo-Json
    }

    return Microsoft.PowerShell.Utility\Invoke-RestMethod @requestParams
}

function New-Milestone() {
    param(
        [string]$ProjectUrl,
        [string]$Token,
        [string]$ProjectUrlNoEncode,
        [string]$Prompt,
        [array]$ExistingMilestoneTitles
    )
    $title = Read-MilestoneTitle "$Prompt" $ExistingMilestoneTitles
    $newMilestoneResponse = New-GitLabMilestone "$ProjectURL" "$Token" "$title"
    if ($newMilestoneResponse.message -like "*Failed to create milestone*") {
        throw New-GenerateReleaseException 35 "Failed to create milestone '$title'."
    }

    Write-Log "Created milestone '$title'." -LogLevel Notice
    Write-Log "Make sure to associate your issues with the milestone here https://gitlab.avispl.com/$ProjectUrlNoEncode/-/issues using the Bulk Edit feature." -LogLevel Warn
    return $title
}

function Update-DefaultMilestoneTitle() {
    param(
        $ProjectUrl,
        $Token,
        $GitlabMilestone,
        $ExistingMilestoneTitles
    )
    $newTitle = Read-MilestoneTitle "Milestone ID $($GitlabMilestone.id) has default title '$($GitlabMilestone.title)'. Provide a new title for it that includes the job or case number and a description:" $ExistingMilestoneTitles
    $updateMilestoneResponse = Update-GitLabMilestoneTitle "$ProjectURL" "$Token" $GitlabMilestone.id "$newTitle"
    Write-Log "Milestone ID $($updateMilestoneResponse.id) title updated to '$($updateMilestoneResponse.title)'"
    return $updateMilestoneResponse.title
}

function New-GitLabRelease($ProjectURL, $Token, $Tag, $BranchRef, $MilestoneTitle) {
    if ($MilestoneTitle) {
        $payload = @{
            name = "$Tag"
            tag_name = "$Tag"
            ref = "$BranchRef"
            milestones = @($MilestoneTitle)
        }
    } else {
        $payload = @{
            name = "$Tag"
            tag_name = "$Tag"
            ref = "$BranchRef"
        }
    }
    if ($MilestoneTitle) {
        Write-Log "Creating release with name $Tag on commit $BranchRef associated with milestone(s) '$MilestoneTitle'"
    } else {
        Write-Log "Creating release with name $Tag on commit $BranchRef"
    }

    $releaseParams = @{
        Method  = 'Post'
        Uri     = "$ProjectURL/releases"
        Headers = @{
            'Private-Token' = $Token
            'Content-Type'  = 'application/json'
        }
        Body    = $payload | Microsoft.PowerShell.Utility\ConvertTo-Json
    }

    return Microsoft.PowerShell.Utility\Invoke-RestMethod @releaseParams
}

function Get-TokenInfo($ApiBaseUrl, $Token){ 
    $tokenParams = @{
        Method = 'Get'
        Uri = "$ApiBaseUrl/personal_access_tokens/self"
        Headers = @{
            'Private-Token' = $Token
            'Content-Type'="application/json"
        }
    }

    return Microsoft.PowerShell.Utility\Invoke-RestMethod @tokenParams
}

function Rotate-Token($ApiBaseUrl, $Token) {
    $newExpiry = $(get-date (get-date).ToUniversalTime().AddYears(1) -UFormat '%Y-%m-%d')
    $payload = @{
        expires_at = $newExpiry
    }

    $updateTokenParams = @{
        Method = 'Post'
        Uri = "$ApiBaseUrl/personal_access_tokens/self/rotate"
        Headers = @{
            'Private-Token' = $Token
            'Content-Type' = "application/json"
        }
        Body = $payload | Microsoft.PowerShell.Utility\ConvertTo-Json
    }

    return Microsoft.PowerShell.Utility\Invoke-RestMethod @updateTokenParams
}

function Update-ProjectFile($ProjectURL, $Token, $FilePath) {
    Write-Log "Uploading $FilePath to the project."
    if ($PSVersionTable.PSVersion.Major -eq 7) {
        $uploadParams = @{
            Method  = 'Post'
            Uri     = "$ProjectURL/uploads"
            Headers = @{
                'Private-Token' = $Token
            }
            Form    = @{file = (Microsoft.PowerShell.Management\Get-Item "$FilePath") }
        }
    }
    elseif ($PSVersionTable.PSVersion.Major -eq 5) {
        $enc = [System.Text.Encoding]::GetEncoding("iso-8859-1")
        $fileBin = [System.IO.File]::ReadAllBytes("$FilePath")
        $fileEnc = $enc.GetString($fileBin)
        $fileName = Microsoft.PowerShell.Management\Split-Path $FilePath -Leaf -Resolve
        # We need a boundary (something random() will do best)
        $boundary = [System.Guid]::NewGuid().ToString()
        # Linefeed character
        $LF = "`r`n"
        $bodyLines = @(
            "--$boundary",
            "Content-Disposition: form-data; name=`"file`"; filename=`"$fileName`"",
            "Content-Type: application/octet-stream$LF",
            $fileEnc,
            "--$boundary--$LF"
        ) -join $LF

        $uploadParams = @{
            Method  = 'Post'
            Uri     = "$ProjectURL/uploads"
            Headers = @{
                'Private-Token' = $Token
                'Content-Type'  = "multipart/form-data; boundary=$boundary"
            }
            Body    = $bodyLines
        }
    }
    
    return Microsoft.PowerShell.Utility\Invoke-RestMethod @uploadParams
}

function Update-ReleaseLink($ProjectURL, $Token, $TagName, $LinkName, $AssetUrl, $Type) {
    Write-Log "Updating release with asset links"
    $payload = @{
        name      = $LinkName
        url       = $AssetUrl
        link_type = $Type
    }

    $releaseParams = @{
        Method  = 'Post'
        Uri     = "$ProjectURL/releases/$TagName/assets/links"
        Headers = @{
            'Private-Token' = $Token
            'Content-Type'  = 'application/json'
        }
        Body    = $payload | Microsoft.PowerShell.Utility\ConvertTo-Json
    }

    return  Microsoft.PowerShell.Utility\Invoke-RestMethod @releaseParams
}

function Update-ReleaseDescription($ProjectURL, $Token, $TagName, $Description) {
    Write-Log "Updating release description"
    $payload = @{
        description = $Description
    }

    $releaseParams = @{
        Method  = 'Put'
        Uri     = "$ProjectURL/releases/$TagName/"
        Headers = @{
            'Private-Token' = $Token
            'Content-Type'  = 'application/json'
        }
        Body    = $payload | Microsoft.PowerShell.Utility\ConvertTo-Json
    }

    return Microsoft.PowerShell.Utility\Invoke-RestMethod @releaseParams
}

function Test-ArrayContains([array]$groups, [array]$iore) {
    foreach($term in $groups) {
        if ($iore -contains $term) {
            return $true
        }
    }

    return $false
}

function Remove-AllFilesInGitignore {
    param(
        [array]$files
    )
    $finalFiles = @()
    foreach($file in $files) {
        $resolvedPath = Resolve-RelativePath $file.FullName
        if (!(Test-IsInGitIgnore $resolvedPath)) {
            $finalFiles+=$file
        }
    }

    return $finalFiles
}

function Test-IsInGitignore($path) {
    $output = Invoke-GitCheckIgnore $path
    if ($output) { 
        return $true
    }

    return $false
}

function Set-Included($files) {
    $includedCount=0
    Write-Log "Determining file inclusion via -IncludeGroups / -ExcludeGroups"
    foreach($file in $files) {
        $groups = $file.Groups
        if ($groups) {
            $included = $false
            if ($IncludeGroups) {
                if (Test-ArrayContains $groups $IncludeGroups) {
                    $included = $true
                }
            }
            elseif ($ExcludeGroups) {
                if (Test-ArrayContains $groups $ExcludeGroups) {
                    $included = $false
                } else {
                    $included = $true
                }
            }
            elseif ($ExcludeAllGroups) {
                $included = $false
            } else{
                $included = $true
            }
        } else {
            $included = $true
        }
        if ($included) {
            $includedCount += 1
        }
        $file | Add-Member -NotePropertyName Included -NotePropertyValue $included
    }

    return $includedCount
}

function Build-PackageName() {
    $pn = "$($BuildProperties.ReleaseName)-$(Build-TagName)"
    return $pn
}

function Build-TagName {
    $tag = $BuildProperties.VersionNumber
    $suffix = ""
    if ($GroupReleaseSuffix) {
        return "$tag-$GroupReleaseSuffix"
    }
    elseif ($BuildProperties.GroupReleaseSuffixes -and $IncludeGroups) {
        $ht = [ordered]@{}
        $BuildProperties.GroupReleaseSuffixes.PSObject.Properties | Microsoft.PowerShell.Core\ForEach-Object { $ht[$_.Name]=$_.Value}
        [String[]]$includes = @()
        foreach ($key in $ht.Keys) {
            if ($IncludeGroups -contains $key) {
                if ($($ht[$key] -match "^[a-zA-Z0-9\.-]+$") -eq $false) {
                    throw New-GenerateReleaseException 2 "Invalid characters in GroupReleaseSuffixes[$key] = $($ht[$key]). Only alpha, numeric, dot, and dash are allowed."
                }
                $includes += $ht[$key]
            }
        }
        $suffix=$($includes -join "_")
        return "$tag-$suffix"
    }

    return $tag
}

class GenerateReleaseException : Exception {
    [int]$code

    
    GenerateReleaseException([int]$code, [string[]]$messages)
    : base($messages -join "`n") {
        $this.code = $code
    }
}

function New-GenerateReleaseException() {
    param(
        [Parameter(Mandatory=$true)][int]$code,
        [Parameter(ValueFromRemainingArguments=$true)][string[]]$messages
    )
    $gre = [GenerateReleaseException]::new($code, $messages)
    return $gre
}

function Handle-TerminatingGenerateReleaseException([GenerateReleaseException]$exception) {
    $m = $exception.Message
    $c = $exception.code

    if ($m) {
        write-log $m -LogLevel Error  
    }

    Set-Exit $c
}

# Invokes the script at $path with parameter $param, returning its exit code.
function Invoke-VerifySimplWindowsRefs($path, $param) {
    return $((& $path "$param" -NoUpdate); $LASTEXITCODE)
}

# draws the current menu state.
function Draw-Menu {
    param(
        $Options,
        $Position,
        $CurrentSelection,
        $Multiselect,
        $AddNewIndex
    )

    for($i=0;$i -le $Options.Count;$i++){
        if ($null -ne $Options[$i]){
            $Option=$Options[$i]
            if($Multiselect){
                if ($CurrentSelection -contains $i) {
                    $Option="[*] $Option"
                } else {
                    if ($i -ne $AddNewIndex) {
                        $Option="[ ] $Option"
                    } else {
                        $Option="$Option"
                    }
                }
            }
            
            if ($i -eq $Position){
                Microsoft.PowerShell.Utility\Write-Host "> $Option" -ForegroundColor Green
            } else {
                if ($i -eq $AddNewIndex) {
                    Microsoft.PowerShell.Utility\Write-Host "  $Option" -ForegroundColor Yellow
                } else {
                    Microsoft.PowerShell.Utility\Write-Host "  $Option"
                }
            }
        }
    }
}

function Set-MenuMultiSelect {
    param(
        $Position,
        [array]$CurrentSelection
    )

    if ($Position -eq $AddNewIndex) {
        # code goes here to execute the option addition.
        # then, clear out current selection and return the list as though there was a new option 
        # that was the only one that was selected
    }

    if ($CurrentSelection -contains $Position){
        $Result = $CurrentSelection | Microsoft.PowerShell.Core\Where-Object { $_ -ne $Position }
    } else {
        $CurrentSelection += $Position
        $result = $CurrentSelection
    }

    return $Result
}

function Read-RawUiKey {
    return $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown").VirtualKeyCode
}

function Set-ConsoleCursorVisible {
    param(
        [bool]$NewVisibility
    )
    [System.Console]::CursorVisible=$NewVisibility
}

function Get-ConsoleCursorTop {
    return [System.Console]::CursorTop
}

function Set-ConsoleCursorPosition {
    param(
        [int]$NewPosition
    )
    [System.Console]::SetCursorPosition(0, $NewPosition)
}


function New-MilestoneMenu {
    param(
        [string]$ProjectUrl,
        [string]$Token,
        [string]$ProjectUrlNoEncode,
        [string]$Title,
        [array]$Milestones,
        [switch]$Multiselect,
        [array]$ExistingMilestoneTitles
    )

    $Key = 0
    $Position = 0
    $CurrentSelection = @()
    $Options=@()
    foreach ($milestone in $Milestones){
        # test milestone title for default values
        if ($DefaultMilestoneNames -contains $milestone.title) {
           $updatedTitle = Update-DefaultMilestoneTitle -ProjectUrl $ProjectUrl -Token $Token -GitlabMilestone $milestone -ExistingMilestoneTitles $existingMilestoneTitles
           $Options += $updatedTitle
        } else {
           $Options+=$milestone.title
        }
    }

    $Options += "Create and use a new milestone for this release"
    $AddNewMilestoneIndex = $Options.Count - 1
    if ($Options.Count -gt 0){
        try {
            Write-Log "$Title"
            Write-Log " Use the up/down arrow keys to select a milestone."
            Write-Log " Press spacebar to select an option."
            Write-Log " Press enter to confirm selections."
            Write-Log " Press escape to cancel associating this release with a milestone."
            Set-ConsoleCursorVisible $False
            Draw-Menu -Options $Options -Position $Position -CurrentSelection $CurrentSelection -Multiselect $Multiselect -AddNewIndex $AddNewMilestoneIndex
            # Key code 13 is the enter key. Key code 27 is the escape key.
            while ($Key -ne 13 -and $Key -ne 27) {
                $Key = Read-RawUiKey
                # Up arrow
                if ($Key -eq 38) {
                    $Position--
                }
                # Down arrow
                if ($Key -eq 40){
                    $Position++
                }
                # Down arrow at final option wraps to first option.
                if ($Position -eq $Options.Count){
                    $Position = 0
                }
                # Up arrow at first option wraps to last option.
                if ($Position -lt 0) {
                    $Position = $Options.Count -1
                }

                # Spacebar selects the current item
                if ($Key -eq 32) {
                    # if the user is selecting the "Add new milestone index"
                    if ($Position -eq $AddNewMilestoneIndex) {
                        $AddNewChosen = $true
                        $newMilestoneTitle = New-Milestone -ProjectUrl $ProjectUrl -Token $Token -ProjectUrlNoEncode $ProjectUrlNoEncode -Prompt "Provide a title for a new milestone that includes the job or case number and a description:" -ExistingMilestoneTitles $Options
                        # ignore other selections and use only the newly created one.
                        return $newMilestoneTitle
                    } else {
                        $CurrentSelection = Set-MenuMultiSelect -Position $Position -CurrentSelection $CurrentSelection
                    }
                }

                # escape cancels all selections and the menu
                if ($Key -eq 27) {
                    return $null
                }

                # Re-draw the menu with new position.
                if ($Key -ne 27){ 
                    try {
                        $cursorTop = Get-ConsoleCursorTop
                        $NewPosition = $cursorTop-$Options.Count
                        Set-ConsoleCursorPosition $NewPosition
                    }catch {
                        Clear-Host
                    }

                    Draw-Menu -Options $Options -Position $Position -Multiselect $Multiselect -CurrentSelection $CurrentSelection -AddNewIndex $AddNewMilestoneIndex
                }
            }
        }
        finally {
            try
            { 
                if (!$AddNewChosen) {
                    # If the user chose to add a new milestone, the cursor position will not be correctly tracked
                    # and thus we should not reset its position.
                    $updatedPosition = $NewPosition + $Options.Count
                    Set-ConsoleCursorPosition $updatedPosition
                }
            }
            catch {
                Clear-Host
            }
            Set-ConsoleCursorVisible $true
        }
    } else {
        $Position = $null
    }

    if ($Multiselect) {
        return $Options[$CurrentSelection]
    } else {
        return $Options[$Position]
    }
}

# returns a list of Hashtables, where each hashtable contains an "item", "destination", and "isDir" key.
# recurses through all directories in $path, skipping directories and files that are in the gitignore.
# this is used by GenerateRelease to copy directories that contain Crestron Construct solutions, preserving 
# directory structure while also not iterating over every file, since __TEMP contains thousands of files that 
# can't efficiently be checked for .gitignore presence.  
# destinationDir is used to compute final destination of files and directories.
function Get-UnignoredFiles {
    param(
        $path,
        $destinationDir
    )

    $finalItems=@()
    $resolvedPath = Resolve-RelativePath $path
    $destinationPath = "$destinationDir/$(Microsoft.PowerShell.Management\Split-Path -Path $resolvedPath -Leaf)"
    $finalItems+=@{item=$destinationPath; isDir=$true}
    $currentItems = Microsoft.PowerShell.Management\Get-ChildItem "$path"
    foreach ($item in $currentItems) {
        $resolvedItemPath = Resolve-RelativePath $item.FullName
        if (Test-IsInGitignore $resolvedItemPath) {
            # skip directories and files that are in the .gitignore
            continue
        }

        if ($item.PSIsContainer) {
            # recurse into directories that aren't in the .gitignore
            $finalItems += Get-UnignoredFiles $resolvedItemPath $destinationPath
        } else {
            # base recursion case
            $finalItems += @{item=$resolvedItemPath;destination=$destinationPath; isDir=$false}
        }
    }

    return $finalItems
}

# in this function, $destinationDir is ALWAYS a dir and is not the location of a renamed file.
# if it doesn't exist, it will be created. this means that recursive directory copies 
# will be in a folder with the same name under $destinationDir; the copy process will 
# never *create* $destinationDir.
function Copy-ReleaseFile($sourcePath, $destinationDir, $vswrPath){
    try {
        $sourceItem = $(Microsoft.PowerShell.Management\Get-Item $sourcePath)
    }
    catch {
        throw New-GenerateReleaseException 4 "Cannot find path $sourcePath. Is it within the git repository?"
    }

    $resolvedOutputPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($destinationDir)
    if (!(Microsoft.PowerShell.Management\Test-Path -LiteralPath "$resolvedOutputPath")) {
        Microsoft.PowerShell.Management\New-Item -ItemType Directory -Path $resolvedOutputPath > $null
    }

    if ($sourceItem.PSIsContainer) {
        Write-Log "Copying directory $sourcePath"
        if (Microsoft.PowerShell.Management\Test-Path "$sourcePath\*.csln") {
            # toCopy will be a list of objects with 'item', 'isDir', and 'destination' properties.
            # item denotes a source file path, 'destination' denotes its containing directory
            # if isDir is true, this indicates a directory with path 'item' should be created and the 'destination' property is unused.
            $toCopy = Get-UnignoredFiles $sourcePath $destinationDir
            foreach ($file in $toCopy){
                try{
                    if ($file.isDir){
                        Microsoft.PowerShell.Management\New-Item -ItemType Directory -Path "$($file.item)" > $null
                    } else {
                        Microsoft.PowerShell.Management\Copy-Item "$($file.item)" -Destination "$($file.destination)"
                    }
                } catch {
                    throw New-GenerateReleaseException 4 "Could not copy file $($file.item) into $($file.destination)"
                }
            }
        } else {
            try{
                Microsoft.PowerShell.Management\Copy-Item "$sourcePath" -Destination "$destinationDir" -Recurse
            } catch {
                throw New-GenerateReleaseException 4 "Could not copy directory $sourcePath into $destinationDir"
            }
        }
    } else {
        Write-Log "Copying file $sourcePath"
        if ($sourcePath.EndsWith(".lpz")) {
            $smwSource = $($sourcePath -replace ".lpz",".smw")
            $foundSmw = $false
            if (!(Microsoft.PowerShell.Management\Test-Path $smwSource)) {
                if (!(Microsoft.PowerShell.Management\Test-Path -LiteralPath $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($smwSource))) {
                    Confirm-YesOrNo "$sourcePath is a release artifact, but no corresponding .smw produces it. Are you sure you wish to continue?"
                }
                else {
                    if ($DebugEnable) { Write-Log "Found the SMW using the LiteralPath option." -LogLevel Notice }
                    $foundSmw = $true
                    $smwSource = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($smwSource)
                }
            }
            else {
                $foundSmw = $true
            }

            if ($foundSmw -and $vswrPath) {
                Write-Log "Verifying SIMPL Windows references within $smwSource" -LogLevel Warn
                $smwRefsTestResult = Invoke-VerifySimplWindowsRefs $vswrPath "$smwSource"
                if (!$smwRefsTestResult -ne "0") {
                    throw New-GenerateReleaseException 30 "$smwSource references modules that are not present in this repository. Aborting release."
                }
            }
        }

        try {
            Microsoft.PowerShell.Management\Copy-Item -LiteralPath "$sourcePath" -Destination "$destinationDir"
        }
        catch {
            try {
                Microsoft.PowerShell.Management\Copy-Item "$sourcePath" -Destination "$destinationDir"
            }
            catch {
                $message = "Could not copy $sourcePath into $destinationDir"
                if ((Get-PathLength $sourcePath $destinationDir "Error") -ge 260) {
                    $message = "Could not copy the file because its destination path exceeds 260 characters."
                }

                throw New-GenerateReleaseException 4 $message
            }
        }
    }

    return $resolvedOutputPath
}

function Copy-ReleaseFiles($BuildProperties, $BasePath) {
    $BuildDir = $BuildProperties.StagingPath

    if (!$($BuildProperties.Copy)) {
        throw New-GenerateReleaseException 28 "$PropertiesFile contains an empty 'Copy' array and the resulting release will have no content. Is $PropertiesFile correct?"
    }

    if (Microsoft.PowerShell.Management\Test-Path $BuildDir) {
        Write-Log "Removing previous local releases within $BuildDir"
        if ($PSVersionTable.PSVersion.Major -ge 7) {
            Microsoft.PowerShell.Management\Remove-Item -Recurse $BuildDir -ProgressAction SilentlyContinue
        } else {
            Microsoft.PowerShell.Management\Remove-Item -Recurse $BuildDir
        }
    }

    $PackageName = $(Build-PackageName $BuildProperties)

    $PackagePath = "$BasePath\$BuildDir\$PackageName"

    # Check to see if build directory exists
    if (!$(Remove-DirIfExists $PackagePath)) {
        throw New-GenerateReleaseException 5 "Failed to delete $PackagePath. Manually delete $PackagePath"
    }

    # todo: do something smarter
    $startMessage = "Generating release package $PackageName"
    # Create the build directory
    if ($IncludeGroups) {
        $startMessage = "$startMessage including groups $IncludeGroups"
    }
    if ($ExcludeGroups) {
        $startMessage = "$startMessage excluding groups $ExcludeGroups"
    }
    if ($ExcludeAllGroups) {
        $startMessage = "$startMessage excluding all groups"
    }

    Write-Log "$startMessage in $PackagePath"
    
    Microsoft.PowerShell.Management\New-Item -ItemType directory -Path "$PackagePath" > $null

    $includedCount = $(Set-Included $BuildProperties.Copy)

    if ($includedCount -eq 0) {
        throw New-GenerateReleaseException 25 "The release would contain no files with the -IncludeGroups / -ExcludeGroups / -ExcludeAllGroups provided."
    }

    $verifySimplWindowsRefsPath = ""
    if (!$SkipVerifySmw){
        foreach ($file in $BuildProperties.Copy) {
            if ($file.Included){
                if ($($file.Filename).EndsWith(".lpz")) {
                    Write-Log "This release contains SIMPL Windows programs. Downloading VerifySimplWindowsRefs to ensure program dependencies are within the repository." -LogLevel Warn
                    $verifySimplWindowsRefsPath = $(Get-VerifySimplWindowsRefs)
                    break
                }
            }
        }
    }

    # Copy everything in the copy directive
    Foreach ($File in $BuildProperties.Copy) {
        if ($File.Included) {
            if ($File.OutputDirectory) {
                $destination = "$PackagePath\$($File.OutputDirectory)"
            }
            else {
                $destination = "$PackagePath\$(Microsoft.PowerShell.Management\Split-Path -Path $File.FileName -Parent)"
            }

            $source = $(Build-Path $BasePath $($File.FileName))
            $resolvedOutputPath = Copy-ReleaseFile $source $destination $verifySimplWindowsRefsPath
            Write-Log "Output path: $resolvedOutputPath" -LogLevel Warn
        } 
        else {
            Write-Log "$($File.FileName) is not included" -LogLevel Warn
        }
    }
    
    if ($verifySimplWindowsRefsPath -and $(Microsoft.PowerShell.Management\Test-Path $verifySimplWindowsRefsPath)) {
        $tempPath = $(Microsoft.PowerShell.Management\Split-Path -Parent $verifySimplWindowsRefsPath)
        Write-Log "Removing $tempPath" -LogLevel Warn

        Microsoft.PowerShell.Management\Remove-Item $tempPath -Recurse
    }

    return "$PackagePath"
}

# Get the full length of the $srcFile and $outputPath, used to provide error feedback when copying files.
# Include $enableLogOutputLevel to print the full path and length to Write-Log at the given LogLevel
function Get-PathLength($srcFile, $outputPath, $enableLogOutputLevel) {
    $srcFile = Split-Path $srcFile -Leaf

    # Get the full path and remove any extraneous relative directory elements from the path, e.g. '\.'
    $fullPath = Get-IOPathGetFullPath((Join-Path -Path $outputPath -ChildPath $srcFile))
    if ($enableLogOutputLevel -and ($fullPath.Length -ge 260)) {
        Write-Log "The path is $($fullPath.Length) characters long.  Your system may not support paths >= 260 characters.  Path: $($fullPath)" -LogLevel $enableLogOutputLevel
    }

    return $fullPath.Length
}

function Get-IOPathGetFullPath($path){
    return [IO.Path]::GetFullPath($path)
}

function Merge-ReleasePackage($BuildProperties, $Asset, $tagName, $ReleaseRef) {
    $infoPath = "$($BuildProperties.StagingPath)/avispl-release-info.txt"
    $remote = Get-GitRemoteUrl
    $projectUrl = Build-GitLabProjectUrl $remote
    if (!$(Remove-FileIfExists $infoPath)) {
        throw New-GenerateReleaseException 7 "Failed to delete file $infoPath. Manually delete $infoPath"
    }
    $assetInfos = @{}
    $projectLink = "https://gitlab.avispl.com/$projectUrl"
    Microsoft.PowerShell.Management\Add-Content -path $infoPath -value "Version      : $tagName"
    Microsoft.PowerShell.Management\Add-Content -path $infoPath -value "Tag          : $tagName"
    Microsoft.PowerShell.Management\Add-Content -path $infoPath -value "Commit Sha   : $ReleaseRef"
    Microsoft.PowerShell.Management\Add-Content -path $infoPath -value "Project Link : $projectLink"
    Microsoft.PowerShell.Management\Add-Content -path $infoPath -value "Release Link : $projectLink/-/releases/$tagName"
    Microsoft.PowerShell.Management\Add-Content -path $infoPath -value "Release Path : $PsScriptRoot"
    Microsoft.PowerShell.Management\Add-Content -path $infoPath -value "Release Date : $(get-date -f 'yyyy-MM-dd:hh:mm:ss')"
    Microsoft.PowerShell.Management\Add-Content -path $infoPath -value "Release User : ${env:UserName}/${env:ComputerName}"
    $crestrons = $(Microsoft.PowerShell.Management\Get-ChildItem "$Asset\*.[lc]pz" -Recurse)

    foreach ($file in $crestrons) {
        $file = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($file.FullName)
        $rename = [IO.Path]::ChangeExtension($file, ".zip")
        Microsoft.PowerShell.Management\Move-Item -LiteralPath $file -Destination $rename -ErrorAction Stop
        $assetInfo = [ordered]@{}
        $guid = [System.Guid]::NewGuid()
        $workDir = "$env:TEMP\$guid"
        if ($DebugEnable) { Write-Log "workDir = $workDir" }
        try {
            if ($DebugEnable) { Write-Log "rename = $rename" }
            Microsoft.PowerShell.Archive\Expand-Archive -LiteralPath $rename -destinationpath $workDir -ErrorAction stop
        }
        catch {
            throw New-GenerateReleaseException 8 "Failed to expand $file for modification. Create an issue on https://gitlab.avispl.com/avispl/modules/tools/generate-release/-/issues"
        }
        
        # Get Crestron Program Info (3- & 4- Series)
        if ([IO.Path]::GetExtension($file).EndsWith(".lpz")) {
            Write-Log "Checking the LPZ."
            $codeArchive = $(Microsoft.PowerShell.Management\Get-ChildItem $workDir *_archive.zip)
            if (!$codeArchive) {
                Microsoft.PowerShell.Utility\Write-Host ""
                Write-Log "$file was compiled without enabling code archive embedding" -LogLevel Warn
                Confirm-YesOrNo "Continue"
            }

            # get full path to file being added to the archive
            $infoFile = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($infoPath)

            # Add the release info to the Crestron SIMPL archive.
            if ($codeArchive) {
                $codeArchivePath = Microsoft.PowerShell.Management\Convert-Path -LiteralPath "$($codeArchive.FullName)"
                if ($DebugEnable) { Write-Log "codeArchivePath: $codeArchivePath" }

                $archive = Open-Archive $codeArchivePath

                try {
                    if ($DebugEnable) { Write-Log "Adding infoFile: $infoFile, name: $(Split-Path $infoPath -Leaf)" }
                    
                    # Add the release info to the archive
                    Invoke-CreateEntryFromFile $archive $infoFile $(Microsoft.PowerShell.Management\Split-Path $infoPath -Leaf)
                    
                    Write-Log "Added the release info to $(Split-Path $codeArchive -Leaf)" -LogLevel Notice
                }
                catch {
                    throw New-GenerateReleaseException 9 "Failed to add the release info ($infoPath) to the SIMPL archive ($codeArchive). Create an issue on https://gitlab.avispl.com/avispl/modules/tools/gitlab-utility-scripts/-/issues"
                }
                # release the file handle
                $archive.Dispose()
            }

            # Add the updated release info and/or archive to the lpz
            $archive = Open-Archive $rename
            try {
                if ($DebugEnable) { Write-Log "rename: $($archive.Entries.Name)"}
                # if ($DebugEnable) { Write-Log "Adding codeArchivePath (lpz): $codeArchivePath, filename: $(Split-Path $codeArchive -Leaf) into rename: $rename" }
                
                if ($codeArchive) {
                    # Look for the existing _archive.zip file in the lpz and delete it.  Otherwise we end up with a duplicate entry when updating the file.  Using UTF8 or UTF8NoBom encoding when opening the archive did not obviate the need for this step.
                    foreach ($entry in $archive.Entries)
                    {
                        if ($DebugEnable) { Write-Log "entry: $($entry.FullName)"}
                        if ($entry.FullName.Contains((Split-Path $codeArchive -Leaf))) {
                            if ($DebugEnable) { Write-Log "Trying to remove file $(Split-Path $codeArchive -Leaf) from the archive."}
                            $entry.Delete()
                            if ($DebugEnable) { Write-Log "Removed file $(Split-Path $codeArchive -Leaf) from the archive." -LogLevel Notice }

                            # The enumeration will fail if iteration continues after deleting this entry.
                            break
                        }
                    }

                    Invoke-CreateEntryFromFile $archive $codeArchivePath (Microsoft.PowerShell.Management\Split-Path $codeArchive -Leaf)
                    Write-Log "Added the updated archive into $(Microsoft.PowerShell.Management\Split-Path $codeArchive -Leaf)" -LogLevel Notice
                }

                Invoke-CreateEntryFromFile $archive $infoFile (Microsoft.PowerShell.Management\Split-Path $infoPath -Leaf)
                Write-Log "Added the updated release info to $(Split-Path $file -Leaf)" -LogLevel Notice
            }
            catch {
                Write-Log "Exception updating the lpz." -LogLevel Error
                Write-Log $_ -LogLevel Error
                $archive.Dispose()
                throw New-GenerateReleaseException 10 "Failed to update $rename with $codeArchive and/or $infoPath. Create an issue on https://gitlab.avispl.com/avispl/modules/tools/gitlab-utility-scripts/-/issues"
            }
            # release the file handle
            $archive.Dispose()
        }

        if ([IO.Path]::GetExtension($file).EndsWith(".lpz")) {
            if ($DebugEnable) { Write-Log "Extracting SIMPL Windows program info." }
            $bootBtContents = $(Microsoft.PowerShell.Management\Get-Content "$workDir\boot.bt")
            Add-ToHashNonEmpty $assetInfo "Compiled" $(($bootBtContents | Microsoft.PowerShell.Utility\Select-String "CompileDateTime=") -replace "CompileDateTime=", "")
            Add-ToHashNonEmpty $assetInfo "sha1sum" $(Microsoft.PowerShell.Utility\Get-FileHash -algorithm SHA1 -LiteralPath "$rename").hash.ToLower()
            Add-ToHashNonEmpty $assetInfo "Processor" $($($bootBtContents | Microsoft.PowerShell.Utility\Select-String "RackType=") -replace "RackType=", "")
            Add-ToHashNonEmpty $assetInfo "Device Database Version" $($($bootBtContents | Microsoft.PowerShell.Utility\Select-String "DeviceDBVersion=") -replace "DeviceDBVersion=", "")
            Add-ToHashNonEmpty $assetInfo "Crestron Database Version" $($($bootBtContents | Microsoft.PowerShell.Utility\Select-String "CresDBVersion=") -replace "CresDBVersion=", "")
            Add-ToHashNonEmpty $assetInfo "SIMPL Version" $($($bootBtContents | Microsoft.PowerShell.Utility\Select-String "SourceEnvironmentVersion=") -replace "SourceEnvironmentVersion=", "" -replace "SIMPL Windows ", "")
            Add-ToHashNonEmpty $assetInfo "Compiler Version" $($($bootBtContents | Microsoft.PowerShell.Utility\Select-String "CompilerVersion=") -replace "CompilerVersion=", "")
            if ($DebugEnable) { Write-Log "Extracted SIMPL Windows program info." -LogLevel Notice }
        }
        else {
            if ($DebugEnable) { Write-Log "Extracting the SIMPL# PRO program info." }
            Add-ToHashNonEmpty $assetInfo "Compiled" $($([xml]$(Microsoft.PowerShell.Management\Get-Content "$workDir\ProgramInfo.config")).ProgramInfo.OptionalInfo.CompiledOn)
            Add-ToHashNonEmpty $assetInfo "sha1sum" $(Microsoft.PowerShell.Utility\Get-FileHash -algorithm SHA1 -LiteralPath "$rename").hash.ToLower()
            if ($DebugEnable) { Write-Log "Extracte the SIMPL# PRO program info." -LogLevel Notice }
        }

        if ($DebugEnable) { Write-Log "Restoring the file's extension."  }
        $assetInfos.Add([IO.Path]::GetFileName($file), $assetInfo)
        Microsoft.PowerShell.Management\Move-Item -LiteralPath $rename $file

        if (!$(Remove-DirIfExists $workDir)) {
            throw New-GenerateReleaseException 11 "Failed to delete $workDir. Manually delete $workDir"
        }
    }

    $releaseNotes = $(New-ReleaseNotes $BuildProperties $assetInfos $projectLink)

    Microsoft.PowerShell.Management\New-Item -ItemType File -Path "$Asset\ReleaseNotes.md" -Value $releaseNotes  > $null

    if ($PauseBeforeZipping) {
        Microsoft.PowerShell.Utility\Write-Host ""
        Write-Log "Files have been added to the release folder.  Make modifications now before the script continues." -LogLevel Warn
        Pause
    }

    # execute post-copy here
    if (!$SkipPost){
        foreach($p in $BuildProperties.PostScripts){
            write-log "Executing '$p' before creating package archive"
            $p | invoke-expression
        }
    }

    Microsoft.PowerShell.Archive\Compress-Archive "$Asset/*" -DestinationPath "$Asset.zip" -CompressionLevel Optimal
    Microsoft.PowerShell.Archive\Compress-Archive -path $infoPath -update -destinationpath "$Asset.zip"
    Write-Log "Release package ready in $Asset.zip" -LogLevel Notice
    Write-Log ""
    
    return "$Asset.zip"
}

function Add-ToHashNonEmpty($hash, $key, $value) {
    if ($key -and $value) {
        $hash[$key] = $value
    }
}

function Test-AssetFileSize($FilePath) {
    Write-Log "Ensuring $FilePath does not exceed the maximum upload size" -LogLevel Warn
    $size = [math]::Ceiling($($(Microsoft.PowerShell.Management\Get-Item -LiteralPath "$FilePath").Length/1MB))
    if ($(${size} -gt ${MaxUploadFileSizeMb})) {
        throw New-GenerateReleaseException 29 "$FilePath is ${size}MB, which exceeds the maximum allowed size ${MaxUploadFileSizeMb}MB" "Consider splitting this release into multiple parts."
    }
    else {
        if ($DebugEnable) { Write-Log "$FilePath does not exceed the maximum upload size." -LogLevel Notice }
    }
}

function New-Release($ApiProperties, $BasePath, $Asset, $BuildProperties) {
    # Get the project ID from Gitlab using the git remote info
    if ($DebugEnable) { Write-Log "Entering New-Release(ApiProperties: $ApiProperties,`nBasePath: $BasePath,`nAsset: $Asset,`nBuildProperties: $BuildProperties)" }
    $tagName=$(Build-TagName)
    if (!$LocalOnly) {
        if (!($ApiProperties.access_token)) {
            throw New-GenerateReleaseException 33 "No access_token is defined within $ApiPropFilePath, nor is one provided by the -OverrideApiToken flag. Aborting release."
        }

        try {
            $oldTlsSettings = [Net.ServicePointManager]::SecurityProtocol
            $ReleaseRef = Test-GitReleaseReady $tagName
            $remote = Get-GitRemoteUrl
            $projectUrl = Build-GitLabProjectUrl $remote
            $projectUrlEncoded = [System.Web.HttpUtility]::UrlEncode($projectUrl)
            $AssetZip = Merge-ReleasePackage $BuildProperties $Asset $tagName $ReleaseRef
            if ($DebugEnable) { Write-Log "Asset: $Asset`nAssetZip: $AssetZip" }
            Test-AssetFileSize "$Asset.zip"
            [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12, [Net.SecurityProtocolType]::Tls13
            try {
                $releases = Get-GitLabReleaseInfo "$($ApiProperties.url)/projects/$projectUrlEncoded" "$($ApiProperties.access_token)"
            } catch {
                $errorResponse = $_.Exception.Response
                if ($errorResponse.StatusCode -and $errorResponse.StatusCode -eq 404) {
                    throw New-GenerateReleaseException 31 "GitLab responded with a 404 when trying to query existing releases using $($ApiProperties.url)/projects/$projectUrl" "Has this project been moved? Check the banner on GitLab and update your remote URL using git remote set-url origin <new-url>"
                } elseif ($errorResponse.StatusCode -and $errorResponse.StatusCode -eq 401) {
                    throw New-GenerateReleaseException 33 "GitLab responded with a 401 when trying to query existing releases at $($ApiProperties.url)/projects/$projectUrl" "This typically indicates a problem with the access_token within $env:USERPROFILE\.gitlab_api or provided with -OverrideApiToken. Ensure that the token is correct."
                }
                elseif ($errorResponse.StatusCode) {
                    throw New-GenerateReleaseException 32 "GitLab responded with a $($errorResponse.StatusCode) when trying to query existing releases at $($ApiProperties.url)/projects/$projectUrl" "This error is not typically encountered in normal operation. Create an issue on https://gitlab.avispl.com/avispl/modules/tools/generate-release/-/issues"
                } else {
                    throw New-GenerateReleaseException 34 "Encountered an unknown exception from GitLab when trying to query existing releases at $($ApiProperties.url)/projects/${projectUrl}: $_" "Create an issue on https://gitlab.avispl.com/avispl/modules/tools/generate-release/-/issues with a reproduction procedure."
                }
            }

            if ($releases.tag_name -eq $tagName) {
                throw New-GenerateReleaseException 12 "Found release that already references $tagName." "Please update VersionNumber in $PropertiesFile and try again."
            }
            try {
                if (!$MilestonePromptDisable) {
                    $milestonePages = Get-Milestones "$($ApiProperties.url)/projects/$projectUrlEncoded" "$($ApiProperties.access_token)"
                    $activeMilestones = @()
                    $existingMilestoneTitles=@()
                    if ($milestonePages){
                        foreach($page in $milestonePages) {
                            foreach($milestone in $page) {
                                $existingMilestoneTitles += $milestone.title
                                if ($milestone.state -eq "active") {
                                    $activeMilestones += $milestone
                                }
                            }
                        }
                    }

                    $chosenMilestoneTitles = New-MilestoneMenu -ProjectUrl "$($ApiProperties.url)/projects/$projectUrlEncoded" -Token "$($ApiProperties.access_token)" -ProjectUrlNoEncode $projectUrl -Milestones $activeMilestones -Multiselect -ExistingMilestoneTitles $existingMilestoneTitles -Title "Select a milestone from the menu below."
                    if (!$chosenMilestoneTitles) {
                        Write-Log "You have chosen not to associate this release with any milestone. Make sure to do so manually on GitLab." -LogLevel Warn
                    }
                }

                $releaseResponse = New-GitLabRelease "$($ApiProperties.url)/projects/$projectUrlEncoded" "$($ApiProperties.access_token)" "$tagName" "$ReleaseRef" $chosenMilestoneTitles
                if ($releaseResponse.message -eq 'Release already exists') {
                    throw New-GenerateReleaseException 12 "Release already exists."
                }
            }
            catch [GenerateReleaseException] {
                throw $_
            }
            catch {
                try {
                    $jsonError = $($_.ErrorDetails.Message | Microsoft.PowerShell.Utility\ConvertFrom-Json)
                    if ($jsonError.error_description -and $($jsonError.error_description).StartsWith("Token is expired.")) {
                        throw New-GenerateReleaseException 23 "Your gitlab token is expired. Create a new one on https://gitlab.avispl.com/-/profile/personal_access_tokens and update $env:USERPROFILE\.gitlab_api"
                    }
                    else {
                        throw New-GenerateReleaseException 24 "GitLab responded with the following error information:" $jsonError "Check that your access_token within $env:USERPROFILE\.gitlab_api or provided with -OverrideApiToken is correct."
                    }
                }
                catch [GenerateReleaseException] {
                    throw $_
                }
                catch {
                    throw New-GenerateReleaseException 13 "Unmapped error creating release: $($_.ErrorDetails.Message). Create an issue on https://gitlab.avispl.com/avispl/modules/tools/generate-release/-/issues"
                }
            }
            try {
                $file = Update-ProjectFile "$($ApiProperties.url)/projects/$projectUrlEncoded" $($ApiProperties.access_token) "$Asset.zip"
            }
            catch {
                throw New-GenerateReleaseException 14 "Error uploading files: $($_.ErrorDetails.Message). Create an issue on https://gitlab.avispl.com/avispl/modules/tools/generate-release/-/issues"
            }
            try {
                $finalReleaseNotes = $(Microsoft.PowerShell.Management\Get-Content "$Asset\ReleaseNotes.md" -raw)
                Update-ReleaseDescription "$($ApiProperties.url)/projects/$projectUrlEncoded" $ApiProperties.access_token $tagName "$finalReleaseNotes" > $null
            }
            catch {
                throw New-GenerateReleaseException 15 "Error updating release description: $($_.ErrorDetails.Message). Create an issue on https://gitlab.avispl.com/avispl/modules/tools/generate-release/-/issues"
            }
            try {
                $updateResponse = Update-ReleaseLink "$($ApiProperties.url)/projects/$projectUrlEncoded" $($ApiProperties.access_token) $tagName $(Microsoft.PowerShell.Management\Split-Path $file.full_path -Leaf) "https://gitlab.avispl.com$($file.full_path)" 'package'
                Write-Log "Release assets updated with $($updateResponse.name) at $($updateResponse.url)" 
            }
            catch {
                throw New-GenerateReleaseException 16 "Error updating release links: $($_.ErrorDetails.Message). Create an issue on https://gitlab.avispl.com/avispl/modules/tools/generate-release/-/issues"
            }

            Write-URLFiles $BuildProperties.ReleaseName $projectUrl $tagName

            Write-Log "View this release at https://gitlab.avispl.com/$projectUrl/-/releases/$tagName" -LogLevel Notice

            foreach($d in $BuildProperties.CopyReleaseTo){
                $destination = "$d\$([IO.Path]::GetFileName("$Asset.zip"))"
                Write-Log -LogLevel Notice "Copying release to $destination"

                try {
                    Microsoft.PowerShell.Management\Copy-Item -Path "$Asset.zip" -Destination "$destination"
                } catch {
                    Write-Log -LogLevel Warn "Failed to copy release to $destination. See detailed reason below:"
                    Write-Log -LogLevel Warn $_
                }
            }
        }
        finally {
            [Net.ServicePointManager]::SecurityProtocol = $oldTlsSettings
        }
    }
    else {
        $AssetZip = Merge-ReleasePackage $BuildProperties $Asset $tagName "localBuild"
        Test-AssetFileSize "$Asset.zip"
        Write-Log "Local release packaging complete."
        Write-Log "View this release package at $AssetZip" -LogLevel Notice
        foreach($d in $BuildProperties.CopyReleaseTo){
            $destination = "$d/$([IO.Path]::GetFileName($AssetZip))"
            Write-Log "Would copy the release to $destination"
        }
    }
}

function Write-URLFiles($friendlyName, $projectPath, $version) {
    $projectUrlFileName = Join-Path $BuildProperties.StagingPath "$friendlyName-Project.url"
    $releaseUrlFileName = Join-Path $BuildProperties.StagingPath "$friendlyName-$version-Release.url"
    Microsoft.PowerShell.Management\Remove-Item $projectUrlFileName -ErrorAction Ignore
    Microsoft.PowerShell.Management\Remove-Item $releaseUrlFileName -ErrorAction Ignore

    Write-Log "Generating Project url file: $projectUrlFileName"
    Microsoft.PowerShell.Management\Add-Content -Path $projectUrlFileName -Value "[InternetShortcut]"
    Microsoft.PowerShell.Management\Add-Content -Path $projectUrlFileName -Value "URL=https://gitlab.avispl.com/$projectPath"
    
    Write-Log "Generating Release url file: $releaseUrlFileName"
    Microsoft.PowerShell.Management\Add-Content -Path $releaseUrlFileName -Value "[InternetShortcut]"
    Microsoft.PowerShell.Management\Add-Content -Path $releaseUrlFileName -Value "URL=https://gitlab.avispl.com/$projectPath/-/releases/$version"
}

function Compare-Semver($a, $b) {
    $result = 0
    $result = $a.Major.CompareTo($b.Major)
    if ($result -ne 0) { return $result }

    $result = $a.Minor.CompareTo($b.Minor)
    if ($result -ne 0) { return $result }

    $result = $a.Patch.CompareTo($b.Patch)
    if ($result -ne 0) { return $result }

    $ap = $a.Pre
    $bp = $b.Pre
    if ($ap.Length -eq 0 -and $bp.Length -eq 0) { return 0 }
    if ($ap.Length -eq 0) { return 1 }
    if ($bp.Length -eq 0) { return -1 }

    for ($i = 0; $i -lt [Math]::Min($ap.Length, $bp.Length); $i++) {
        $result = $ap[$i].CompareTo($bp[$i])
        if ($result -ne 0) { return $result }
    }

    return 0
}

function ConvertTo-SemanticVersion($tagName) {
    $matches.Clear()
    if ($tagName -match $TagSemverRegexPattern) {
        $major = [int]$matches[1]
        $minor = [int]$matches[2]
        $patch = [int]$matches[3]
        if ($null -eq $matches[4]) { $pre = @() }
        else { $pre = $matches[4].Split(".") }

        return Microsoft.PowerShell.Utility\New-Object PSObject -property @{
            Major   = $major
            Minor   = $minor
            Patch   = $patch
            Pre     = $pre
            TagName = $tagName
        }
    }
    else {
        return $null
    }
}

function Get-LatestRelease($releases) {
    for ($i = 0; $i -lt $releases.Length; $i++) {
        $rank = 0
        for ($j = 0; $j -lt $releases.Length; $j++) {
            $diff = 0
            $diff = Compare-Semver $releases[$i].SemVer $releases[$j].SemVer
            if ($diff -gt 0) {
                $rank++
            }
        }
        $current = [PsObject]$releases[$i]
        Microsoft.PowerShell.Utility\Add-Member -InputObject $current -MemberType NoteProperty -Name Rank -Value $rank -fo
    }

    return $releases | Microsoft.PowerShell.Utility\Sort-Object -property Rank -Descending | Microsoft.PowerShell.Utility\Select-Object -First 1
}

function Build-SemVerString($semver) {
    $result = "$($semver.Major).$($semver.Minor).$($semver.Patch)"
    if ($semver.Pre.Length -gt 0) {
        $result += "-$($semver.Pre -join '.')"
    }

    return $result
}

function Update-GenerateRelease {
    if ($LocalOnly) {
        Write-Log "Running with -LocalOnly. Skipping script update check" -LogLevel Warn
        return $False
    }

    Write-Log "Looking for updates to GenerateRelease.ps1..."
    try {
        $request = $(Microsoft.PowerShell.Utility\Invoke-WebRequest "$GitlabV4ApiUrl/projects/$GenerateReleaseProjectId/releases")
    }
    catch { }

    if ($request.StatusCode -ne 200) {
        Write-Log "Script was unable to examine new releases. Aborting update attempt..."
        return $False
    }

    $requestJson = $($request.Content | Microsoft.PowerShell.Utility\ConvertFrom-Json)
    $releases = $($requestJson | Microsoft.PowerShell.Core\Where-Object { $_.tag_name -match $TagSemverRegexPattern })
    $releases | Microsoft.PowerShell.Core\ForEach-Object { Microsoft.PowerShell.Utility\Add-Member -inputobject $_ -membertype NoteProperty -name SemVer -Value $(ConvertTo-SemanticVersion($_.tag_name)) -fo }
    
    # we only need to compare to the latest release.
    $latestRelease = $(Get-LatestRelease $releases)

    if ($latestRelease) {
        if ((Compare-Semver $latestRelease.SemVer $ScriptVersion) -gt 0) {
            Write-Log "There is a release v$(Build-SemVerString $latestRelease.SemVer) that is newer than the current v$(Build-SemVerString $ScriptVersion)" -LogLevel Warn
            Write-Log "Release Description:" -LogLevel Warn
            Write-Log "$($latestRelease.description)"
            
            $useNewPrompt = $true
            # check for generaterelease in git repo (will fail if not in repo)
            if (!(Invoke-GitLsFiles("GenerateRelease.ps1")) -or (Invoke-GitStatusPorcelainTrackedOnly)) {
                $prompt = "Do you want to update? [y/n]"
                $useNewPrompt = $false
            } else {
                $prompt = "Do you want to update? y = update and commit, u = update only, n = abort [y/u/n]"
            }

            do {
                Write-Log $prompt -LogLevel Warn
                $userInput = $(Microsoft.PowerShell.Utility\Read-Host)
                if ($useNewPrompt) {
                    $updateAndCommit = $($userInput -like "y*")
                    $updateOnly = $($userInput -like "u*")
                } else {
                    $updateAndCommit = $false
                    $updateOnly = $($userInput -like "y*")
                }
                if ($updateAndCommit -or $updateOnly) {
                    $releaseName = "GenerateRelease-$($latestRelease.tag_name)"
                    $releaseZipName = "$releaseName.zip"
                    Microsoft.PowerShell.Utility\Invoke-WebRequest $latestRelease.assets.links.url -outfile $releaseZipName
                    $guid = [System.Guid]::NewGuid()
                    $workDir = "$env:TEMP\$guid"
                    Microsoft.PowerShell.Archive\Expand-Archive -path $releaseZipName -DestinationPath $workDir
                    $ngrp = "$workDir\GitLab\GenerateRelease.ps1"
                    $ngrpNew = "$workDir\GenerateRelease.ps1"
                    if (Microsoft.PowerShell.Management\Test-Path -LiteralPath $ngrp) {
                        Microsoft.PowerShell.Management\Copy-Item $ngrp ".\GenerateRelease.ps1"
                    } elseif (Microsoft.PowerShell.Management\Test-Path -LiteralPath $ngrpNew) {
                        Microsoft.PowerShell.Management\Copy-Item $ngrpNew ".\GenerateRelease.ps1"
                    } else {
                        Write-Log "The release asset posted for $($latestRelease.tag_name) does not contain GenerateRelease.ps1. Create an issue on https://gitlab.avispl.com/avispl/modules/tools/generate-release/-/issues" -LogLevel Warn
                        return $False
                    }

                    Microsoft.PowerShell.Management\Remove-Item -recurse $workDir
                    Microsoft.PowerShell.Management\Remove-Item $releaseZipName
                    if ($updateAndCommit){
                        Invoke-GitAdd "GenerateRelease.ps1"
                        Invoke-GitCommit "chore: updating GenerateRelease.ps1 to $($latestRelease.tag_name)"
                    }
                    
                    Write-Log "GenerateRelease.ps1 has been updated to $($latestRelease.tag_name)"
                    return $True
                }
            } until($userInput -like 'n*')
           
            Write-Log "User chose not to update. Continuing..."
            return $False
        }
    }

    Write-Log "No script update is required. Continuing..."
    return $False
}

function Get-VerifySimplWindowsRefs {
    Write-Log "Downloading latest VerifySimplWindowsRefs release..."
    try {
        $request = $(Microsoft.PowerShell.Utility\Invoke-WebRequest "$GitlabV4ApiUrl/projects/$VerifySimplWindowsRefsProjectId/releases")
    }
    catch { }

    if ($request.StatusCode -ne 200) {
        Write-Log "Unable to examine VerifySimplWindowsRefs releases. SIMPL Windows program module references cannot be checked." -LogLevel Warn
        return ""
    }

    $requestJson = $($request.Content | Microsoft.PowerShell.Utility\ConvertFrom-Json)
    $releases = $($requestJson | Microsoft.PowerShell.Core\Where-Object { $_.tag_name -match $TagSemverRegexPattern })
    $releases | Microsoft.PowerShell.Core\ForEach-Object { Microsoft.PowerShell.Utility\Add-Member -inputobject $_ -membertype NoteProperty -name SemVer -Value $(ConvertTo-SemanticVersion($_.tag_name)) -fo }

    $latestRelease = $(Get-LatestRelease $releases)

    if ($latestRelease) {
        Write-Log "Downloading VerifySimplWindowsRefs $($latestRelease.tag_name)"
        $releaseName = "VerifySimplWindowsRefs-$($latestRelease.tag_name)"
        $releaseZipName = "$releaseName.zip"
        Microsoft.PowerShell.Utility\Invoke-WebRequest $latestRelease.assets.links.url -outfile $releaseZipName
        $guid = [System.Guid]::NewGuid()
        $workDir = "$env:TEMP\$guid"
        Microsoft.PowerShell.Archive\Expand-Archive -path $releaseZipName -DestinationPath $workDir
        $vswrp = "$workDir\VerifySimplWindowsRefs.ps1"
        if (Microsoft.PowerShell.Management\Test-Path -LiteralPath $vswrp) {
            Write-Log "VerifySimplWindowsRefs.ps1 $($latestRelease.tag_name) has been downloaded to $vswrp"
            (Remove-FileIfExists $releaseZipName | out-null)
            return $vswrp
        }
        else {
            Write-Log "The release asset posted for $($latestRelease.tag_name) does not contain VerifySimplWindowsRefs.ps1. Create an issue on https://gitlab.avispl.com/avispl/modules/tools/generate-release/-/issues" -LogLevel Warn
            return ""
        }
    }

    Write-Log "Could not determine the correct version of VerifySimplWindowsRefs.ps1. SIMPL Windows program module references cannot be checked." -LogLevel Warn
    return ""
}

function Invoke-GitLocalBranchName() {
    return "$(git branch --show-current)"
}

function Invoke-GitRemoteRefs() {
    return $(git fetch --all;$?)
}

function Invoke-GitStatusPorcelainTrackedOnly() {
    return $(git status --porcelain -uno)
}

function Invoke-GitStatusPorcelainUntrackedOnly() {
    return $(git status --porcelain -u | Microsoft.PowerShell.Core\Where-Object { "$_" -match "\?\?" -and $_ -match ".*Code\/.*|.*Configs\/.*|.*Documentation\/.*|.*Symphony\/.*|.*TPUG\/.*|.*UI\/.*" })
}

function Invoke-GitExistingTag(){
    return $(git ls-remote --tags --refs origin $tagName)
}

function Invoke-GitLocalSha($localRef){
    return "$(git rev-parse $localRef)"
}

function Invoke-GitRemoteSha($remoteRef){
    return "$(git rev-parse $remoteRef)"
}

function Invoke-GitLsFiles($file){
    return $(git ls-files $file)
}

function Invoke-GitAdd($path){
    return $(git add $path)
}

function Invoke-GitCommit($message){
    return $(git commit -m "$message")
}

function Invoke-GitCheckIgnore($path){
    return $(git check-ignore $path)
}

function Open-Archive($path) {
    # Use .NET methods to be able to handle filenames with [].
    return [System.IO.Compression.ZipFile]::Open($path, 'update')
}

function Invoke-CreateEntryFromFile($archive, $filePath, $fileName) {
    [System.IO.Compression.ZipFileExtensions]::CreateEntryFromFile($archive, $filePath, $fileName, [System.IO.Compression.CompressionLevel]::Fastest)
}

function Resolve-RelativePath(){
    param(
        [string]$path
    )

    $path = Microsoft.PowerShell.Management\Resolve-Path -LiteralPath $path -ErrorAction SilentlyContinue -ErrorVariable resolvePathError -Relative

    if (-not($path)){
        $path = $resolvePathError[0].TargetObject
    }

    return $path
}

function Test-GitReleaseReady($tagName) {
    Write-Log "Checking git repository state for release readiness..."

    if (Microsoft.PowerShell.Management\Test-Path -LiteralPath "$BasePath/.git/config") {
        $remote = Get-GitRemoteUrl
        if (!$remote) {
            throw New-GenerateReleaseException 17 "The remote push URL for 'origin' is either using the 'https' protocol instead of the 'git@' protocol, or it is malformed." "Ensure that the git repository is cloned or configured using the 'git@' protocol. See the output of 'git remote -v' for more information."
        }
    }
    else {
        throw New-GenerateReleaseException 18 "./.git directory not found. Make sure the script is executed in a git repository at the root level."
    }
    
    $currentBranch = Invoke-GitLocalBranchName

    if (!$currentBranch) {
        throw New-GenerateReleaseException 19 "The local repository HEAD doesn't point to a branch. Aborting release." "Are you in a 'detached HEAD' state? Ensure you have the right branch checked out using 'git branch' and 'git checkout'."
    }

    Write-Log "Fetching remote references to ensure local ref is pushed..."
    if (!(Invoke-GitRemoteRefs)) {
        throw New-GenerateReleaseException 27 "Failed to fetch remote references. Are you connected to the VPN?"
    }

    $localTrackedChanges = Invoke-GitStatusPorcelainTrackedOnly
    if ($localTrackedChanges) {
        throw New-GenerateReleaseException 20 "The local repository contains uncommitted changes in tracked files. Aborting release." $localTrackedChanges "Ensure that you have committed and pushed your local changes using 'git commit' and 'git push', or stash your local changes using 'git stash'."
    }

    $vtpFiles = $(Get-ChildItem "$BasePath" -Filter "*.vtp" -Recurse)
    foreach ($vtp in $vtpFiles) {
        $relativeVtp = Resolve-RelativePath $vtp.FullName
        $expectedVta = $relativeVtp -replace ".vtp",".vta"
        if (!(Invoke-GitLsFiles $expectedVta)) {
            throw New-GenerateReleaseException 40 "The vtp $relativeVtp has no corresponding vta at $expectedVta. Use VT-Pro to export an archive, add it using 'git add', add it to your build.properties, and re-run this command."
        }

        $expectedVtaFile = $(Get-Item $expectedVta)
        if ($vtp.LastWriteTime -gt $expectedVtaFile.LastWriteTime){
            Write-Log "$relativeVtp has a LastWriteTime that is more recent than $expectedVta. Export the archive to update the vta, or enable the 'Automatically Archive on Save' feature in VT-Pro." -LogLevel Warn
            Confirm-YesOrNo "Continue anyway"
        }
    }

    $potentialMissedAdds = Invoke-GitStatusPorcelainUntrackedOnly

    if ($potentialMissedAdds) {
        Write-Log "There are untracked files within the standard job directories:" -LogLevel Warn
        $potentialMissedAdds | Microsoft.PowerShell.Core\ForEach-Object { Write-Log $_ -LogLevel Warn }
        Write-Log "If omitted by accident, stop now and use 'git add', 'git commit', and 'git push' before continuing to generate the release." -LogLevel Warn
        Confirm-YesOrNo "Continue anyway"
    }

    $existingTag = Invoke-GitExistingTag

    if ($existingTag) {
        throw New-GenerateReleaseException 21 "A tag with name $tagName already exists. Aborting release."
    }

    $localSha = "$(Invoke-GitLocalSha $currentBranch)"
    $remoteSha = "$(Invoke-GitRemoteSha "origin/$currentBranch")"

    if ($localSha -ne $remoteSha) {
        throw New-GenerateReleaseException 22 "The local ref branch $currentBranch does not match ref origin/$currentBranch" "Ensure that you have pushed this ref using 'git push' or 'git push origin $currentBranch'"
    }

    return $localSha
}

function Main() {
    Write-Log "GenerateRelease.ps1 v$(Build-SemVerString $ScriptVersion) on PowerShell v$($PSVersionTable.PSVersion)"
    
    if (Update-GenerateRelease) {
        Set-Exit(0)
    }
    
    if ($IncludeGroups -and $ExcludeGroups) {
        Write-Log "Using both -IncludeGroups and -ExcludeGroups is not allowed." -LogLevel Error
        Set-Exit(26)
    }
    
    # If build.properties exists, load it and validate it.
    if (Microsoft.PowerShell.Management\Test-Path $PropertiesFile) {
        $BuildPropFullPath = Microsoft.PowerShell.Management\Resolve-Path -LiteralPath $PropertiesFile
        $BasePath = $BuildPropFullPath | Microsoft.PowerShell.Management\Split-Path -Parent
        try {
            $BuildProperties = Import-BuildProperties "$BuildPropFullPath"
        } catch [GenerateReleaseException] {
            Handle-TerminatingGenerateReleaseException $_.Exception
        }
        if ($UpdateFiles) {
            Write-Log "Updating $PropertiesFile with new files within $BasePath" -LogLevel Warn
            $json = (Update-BuildProperties "$BasePath" $BuildProperties | Microsoft.PowerShell.Utility\ConvertTo-Json -depth 100)
            Microsoft.PowerShell.Management\Set-Content -Path "$PropertiesFile" -Value $json
            Write-Log "Updates to $PropertiesFile complete."
            Set-Exit(0)
        } else {
            if ($BuildProperties.PreScripts -or $BuildProperties.PostScripts) {
                write-log "***************************************************************************************************************" -LogLevel Warn
                write-log "This build.properties currently contains references to execute external scripts. This is potentially dangerous." -LogLevel Warn
                write-log "Please review the PreScripts and PostScripts property in your build.properties." -LogLevel Warn
                write-log "***************************************************************************************************************" -LogLevel Warn
                Confirm-YesOrNo "Are you sure you want to continue"
            }
            if (!$SkipPre) {
                foreach ($s in $BuildProperties.PreScripts){
                    Write-Log "Executing '$s' before copying release files"
                    $s | Invoke-Expression
                }
            }

            try {
                $ReleaseFile = Copy-ReleaseFiles $BuildProperties $BasePath
            } catch [GenerateReleaseException] {
                Handle-TerminatingGenerateReleaseException $_.Exception
            }
        }
    } 
    else {
        Write-Log "Did not find $PropertiesFile in the script directory. Creating a default one" -LogLevel Warn
        Write-Log "Review it, then run this command again." -LogLevel Warn
        $parent = "$PropertiesFile" | Microsoft.PowerShell.Management\Split-Path -Parent
        if (!$parent) {
            $parent = "."
        }
        $json = (New-BuildProperties $parent) | Microsoft.PowerShell.Utility\ConvertTo-Json
        Microsoft.PowerShell.Management\New-Item -ItemType File -Path "$PropertiesFile" -Value $json
        Set-Exit(0)
    }
    
    if ($LocalOnly) {
        Write-Log "Testing local release package creation"
    }
    else {
        Write-Log "Posting release to GitLab"
    }
    # Load API properties
    $ApiPropFilePath = "$env:USERPROFILE\.gitlab_api"
    $ApiProperties = Import-APIProperties $ApiPropFilePath
    
    if ($ApiProperties.Length -eq 0) {
        $ApiProperties = New-APIProperties $ApiPropFilePath $GitlabV4ApiUrl
    }
    
    if ($OverrideApiToken) {
        Write-Log "Overriding access_token within $env:USERPROFILE\.gitlab_api" -LogLevel Warn
        $ApiProperties.access_token = $OverrideApiToken
    }
    
    try {
        if ($DebugEnable) { Write-Log "Calling New-Release" }
        New-Release $ApiProperties $(Microsoft.PowerShell.Management\Resolve-Path -LiteralPath $PropertiesFile | Microsoft.PowerShell.Management\Split-Path -Parent) $ReleaseFile $BuildProperties $branchRef
    } catch [GenerateReleaseException] {
        Handle-TerminatingGenerateReleaseException $_.Exception
    } catch {
        Write-Log "Unexpected exception occurred during New-Release:" -LogLevel Error
        Write-Log $_ -LogLevel Error
        Set-Exit(1)
    }
}

if ($MyInvocation.InvocationName -ne ".") {
    Main
}