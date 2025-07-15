param(
    [Parameter(Mandatory = $true)]
    [string]$ProjectPath,
    
    [Parameter(Mandatory = $true)]
    [string]$ModelPath,
    
    [switch]$ShowInTerminal,
    
    [string]$OutputFile,
    
    [switch]$ShowDebug,
    
    [string]$CompareLabelFile,
    
    [switch]$IncrementalUpdate,
    
    [switch]$ReplaceExisting,
    
    [switch]$AutoCreateLabels,

    [switch]$FindMissingLabels,
    
    [string]$DefaultProjectsPath = "C:\DevOps\Dev\Projects",
    
    [string]$DefaultModelPath = "K:\AosService\PackagesLocalDirectory\STC\STC"
)

#region Funções e Classes Auxiliares

class LabelInfo {
    [string]$LabelId; [string]$ObjectName; [string]$ObjectType; [string]$PropertyType; [string]$FilePath
    LabelInfo([string]$l, [string]$o, [string]$ot, [string]$p, [string]$f) {
        $this.LabelId = $l; $this.ObjectName = $o; $this.ObjectType = $ot; $this.PropertyType = $p; $this.FilePath = $f
    }
}

function Write-DebugInfo {
    param([string]$Message)
    if ($ShowDebug) { Write-Host "[DEBUG] $Message" -ForegroundColor Gray }
}

function Get-ProjectsFromSolution {
    param([string]$SolutionPath)
    $projects = @()
    if (Test-Path $SolutionPath) {
        $content = Get-Content $SolutionPath
        foreach ($line in $content) {
            if ($line -match 'Project\(".*"\) = ".*", "(.*\.rnrproj)"') {
                $projectPath = Join-Path (Split-Path $SolutionPath) $matches[1]
                if (Test-Path $projectPath) { $projects += $projectPath }
            }
        }
    }
    return $projects
}

function Get-ElementsFromProject {
    param([string]$ProjectPath)
    $elements = @()
    if (Test-Path $ProjectPath) {
        [xml]$projectXml = Get-Content $ProjectPath
        foreach ($item in $projectXml.Project.ItemGroup.Content) {
            if ($item.Include) {
                # Usa a pasta como o tipo do elemento, que é mais confiável
                $folder = ($item.Include.Split('\') | Select-Object -First 1)
                $elements += @{ Name = $item.Name; Type = $folder }
            }
        }
    }
    return $elements
}

function Find-ElementInModel {
    param([string]$ModelPath, [string]$ElementName, [string]$Folder)
    
    # Mapeia o nome da pasta para a extensão do arquivo de metadados do D365
    $extensionMapping = @{
        'Classes' = 'Class'; 'Forms' = 'Form'; 'Tables' = 'Table'; 'Views' = 'View';
        'Queries' = 'Query'; 'Reports' = 'Report'; 'Extended Data Types' = 'Edt'; 'Enums' = 'Enum';
        'Enum Extensions' = 'EnumExtension'; 'Form Extensions' = 'FormExtension'; 'Table Extensions' = 'TableExtension';
        'View Extensions' = 'ViewExtension'; 'Menu Items\Display' = 'MenuItemDisplay'; 'Menu Items\Output' = 'MenuItemOutput';
        'Menu Items\Action' = 'MenuItemAction'; 'Label Files' = 'LabelFile'
    }

    $fileExtension = if ($extensionMapping.ContainsKey($Folder)) { ".$($extensionMapping[$Folder])" } else { "" }
    $fileName = "$ElementName$fileExtension.xml"
    
    Write-DebugInfo "Searching for file '$fileName' in folder '$Folder'"

    $searchPath = Join-Path $ModelPath $Folder
    if (Test-Path $searchPath) {
        $xmlFile = Get-ChildItem -Path $searchPath -Filter $fileName -Recurse -ErrorAction SilentlyContinue
        if ($xmlFile) { return $xmlFile.FullName }
    }
    
    # Fallback para procurar com o nome de arquivo antigo (sem a extensão do tipo)
    $xmlFile = Get-ChildItem -Path $ModelPath -Filter "$ElementName.xml" -Recurse -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($xmlFile) { return $xmlFile.FullName }

    return $null
}

function Extract-LabelsFromXpp {
    param([string]$FilePath, [string]$ObjectName, [string]$ObjectType)
    $labels = @()
    if ($FilePath -and (Test-Path $FilePath)) {
        $content = Get-Content $FilePath -Raw -Encoding UTF8
        $labelPatterns = @(
            '<Label>@([^:]+):([^<]+)</Label>', '<Text>@([^:]+):([^<]+)</Text>', '<HelpText>@([^:]+):([^<]+)</HelpText>',
            'Label="@([^:]+):([^"]+)"', 'Text="@([^:]+):([^"]+)"', 'HelpText="@([^:]+):([^"]+)"', 'Caption="@([^:]+):([^"]+)"',
            '@([^:]+):([A-Za-z0-9_]+)', '"@([^:]+):([^"]+)"', "'@([^:]+):([^']+)'"
        )
        foreach ($pattern in $labelPatterns) {
            $matches = [regex]::Matches($content, $pattern, [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
            foreach ($match in $matches) {
                if ($match.Groups[1].Value -eq "STC") {
                    $labels += [LabelInfo]::new($match.Groups[2].Value, $ObjectName, $ObjectType, "Label", $FilePath)
                }
            }
        }
    }
    return $labels
}

function Read-ExistingLabelFile {
    param([string]$FilePath)
    $existingLabels = @{ ById = @{}; ByValue = @{} }
    if (Test-Path $FilePath) {
        Write-DebugInfo "Reading existing label file: $FilePath"
        $content = Get-Content $FilePath -Encoding UTF8
        foreach ($line in $content) {
            if ($line -and $line.Contains('=') -and -not $line.StartsWith(';')) {
                $parts = $line.Split('=', 2)
                if ($parts.Length -eq 2) {
                    $labelId = $parts[0].Trim(); $labelValue = $parts[1].Trim()
                    if (-not [string]::IsNullOrEmpty($labelId)) {
                        $existingLabels.ById[$labelId] = $labelValue
                        if (-not $existingLabels.ByValue.ContainsKey($labelValue)) {
                            $existingLabels.ByValue[$labelValue] = $labelId
                        }
                    }
                }
            }
        }
        Write-Host "[INFO] Comparison file loaded: $($existingLabels.ById.Count) labels" -ForegroundColor Cyan
    }
    else {
        Write-Warning "Comparison file not found: $FilePath. It will be created if new labels are generated."
    }
    return $existingLabels
}

function Convert-And-Replace-HardcodedText {
    param(
        [string]$FilePath, [string]$ObjectName, [string]$ObjectType, [hashtable]$ExistingLabels
    )
    $newLabelsToAdd = [System.Collections.Generic.List[string]]::new()
    if (-not (Test-Path $FilePath)) { 
        Write-Output -NoEnumerate $newLabelsToAdd
        return
    }

    [xml]$xmlContent = Get-Content -Path $FilePath -Encoding UTF8
    $targetTags = "Label", "HelpText", "Text", "Caption"
    $wasModified = $false
    
    # Define o tamanho máximo de caracteres para um texto ser considerado uma label
    $maxLabelLength = 300

    foreach ($tagName in $targetTags) {
        $nodes = $xmlContent.SelectNodes("//*[local-name()='$tagName' and not(ancestor::*[namespace-uri()='http://schemas.microsoft.com/sqlserver/reporting/2016/01/reportdefinition'])]")
        
        foreach ($node in $nodes) {
            $hardcodedText = $node.'#text'.Trim()
            
            # CORREÇÃO: Ignora textos muito longos que provavelmente não são labels (como XMLs inteiros)
            if ($hardcodedText.Length -gt $maxLabelLength) {
                Write-DebugInfo "Skipping text longer than $maxLabelLength characters in '$($node.LocalName)' tag of object '$ObjectName'."
                continue
            }

            if (-not [string]::IsNullOrEmpty($hardcodedText) -and $hardcodedText -notlike "@*") {
                $labelIdToUse = ""
                if ($ExistingLabels.ByValue.ContainsKey($hardcodedText)) {
                    $labelIdToUse = $ExistingLabels.ByValue[$hardcodedText]
                    Write-Host "[AUTO-LABEL] Reusing label for text '$hardcodedText': @STC:$labelIdToUse" -ForegroundColor DarkCyan
                }
                else {
                    $combinedId = "$($ObjectName)$($node.LocalName)$($ObjectType.Replace('Ax', ''))"
                    $baseLabelId = $combinedId -replace '[^a-zA-Z0-9]', ''
                    $newLabelId = $baseLabelId; $counter = 1
                    while ($ExistingLabels.ById.ContainsKey($newLabelId) -or ($newLabelsToAdd | Where-Object { $_.Split('=')[0] -eq $newLabelId })) {
                        $newLabelId = "${baseLabelId}${counter}"; $counter++
                    }
                    $labelIdToUse = $newLabelId
                    $newLabelDefinition = "$labelIdToUse=$hardcodedText"
                    $newLabelsToAdd.Add($newLabelDefinition)
                    $ExistingLabels.ById[$labelIdToUse] = $hardcodedText
                    $ExistingLabels.ByValue[$hardcodedText] = $labelIdToUse
                    Write-Host "[AUTO-LABEL] Text '$hardcodedText' converted to new label: @STC:$labelIdToUse" -ForegroundColor Yellow
                }
                $node.'#text' = "@STC:$labelIdToUse"
                $wasModified = $true
            }
        }
    }

    if ($wasModified) {
        $writerSettings = New-Object System.Xml.XmlWriterSettings
        $writerSettings.Indent = $true; $writerSettings.IndentChars = "    "
        $writerSettings.Encoding = New-Object System.Text.UTF8Encoding($false)
        $writer = [System.Xml.XmlWriter]::Create($FilePath, $writerSettings)
        $xmlContent.Save($writer); $writer.Close()
        Write-Host "[SAVE] XML file '$FilePath' has been updated." -ForegroundColor Green
    }
    
    Write-Output -NoEnumerate $newLabelsToAdd
}

function Show-MissingLabels {
    param([array]$Labels, [hashtable]$ExistingLabels)
    Write-Host "`n[DIAGNOSTIC] Checking referenced labels that don't exist in the file..." -ForegroundColor Cyan
    $missingLabels = @()
    $uniqueLabelReferences = $Labels | Group-Object LabelId
    foreach ($group in $uniqueLabelReferences) {
        if (-not $ExistingLabels.ById.ContainsKey($group.Name)) {
            $missingLabels += $group
        }
    }
    if ($missingLabels.Count -gt 0) {
        Write-Host "`n[WARNING] $($missingLabels.Count) LABELS NOT FOUND IN LABEL FILE:" -ForegroundColor Yellow
        Write-Host "==========================================================" -ForegroundColor Yellow
        foreach ($group in $missingLabels) {
            Write-Host "`n- Missing Label: @STC:$($group.Name)" -ForegroundColor Yellow
            Write-Host "  Referenced in:" -ForegroundColor Gray
            foreach ($item in $group.Group) {
                Write-Host "    - Object: $($item.ObjectName) ($($item.ObjectType))"
                Write-Host "      File: $($item.FilePath)" -ForegroundColor DarkGray
            }
        }
    }
    else {
        Write-Host "[SUCCESS] All referenced labels were found in the label file." -ForegroundColor Green
    }
    return $missingLabels.Count
}

#endregion

# --- Execução Principal ---
try {
    Write-Host "D365 F&O Label Extractor" -ForegroundColor Green; Write-Host "=========================" -ForegroundColor Green

    if ($AutoCreateLabels -and ([string]::IsNullOrEmpty($CompareLabelFile) -or -not (Test-Path (Split-Path $CompareLabelFile -Parent) -PathType Container))) {
        throw "The -AutoCreateLabels parameter requires -CompareLabelFile to be specified with a path where the parent directory exists."
    }
    if ($FindMissingLabels -and [string]::IsNullOrEmpty($CompareLabelFile)) {
        throw "The -FindMissingLabels parameter requires -CompareLabelFile to be specified."
    }
    
    if (-not ($ProjectPath.EndsWith(".sln") -or $ProjectPath.EndsWith(".rnrproj")) -and $ProjectPath -notlike "*\*" -and $ProjectPath -notlike "*/*") {
        Write-Host "[INFO] Simple project name detected. Resolving path for '$ProjectPath'..." -ForegroundColor Cyan
        $potentialSlnPath = Join-Path -Path $DefaultProjectsPath -ChildPath $ProjectPath | Join-Path -ChildPath "$ProjectPath.sln"
        $potentialProjPath = Join-Path -Path $DefaultProjectsPath -ChildPath $ProjectPath | Join-Path -ChildPath "$ProjectPath.rnrproj"
        if (Test-Path $potentialSlnPath -PathType Leaf) { $ProjectPath = $potentialSlnPath; Write-Host "[SUCCESS] Solution found: $ProjectPath" -ForegroundColor Green } 
        elseif (Test-Path $potentialProjPath -PathType Leaf) { $ProjectPath = $potentialProjPath; Write-Host "[SUCCESS] Project file found: $ProjectPath" -ForegroundColor Green } 
        else { throw "Could not find '$($ProjectPath).sln' or '$($ProjectPath).rnrproj' in the default directory. Check the name and path." }
    }
    if (-not (Test-Path $ProjectPath -PathType Leaf)) { throw "The final project path was not found: '$ProjectPath'" }

    $existingLabels = Read-ExistingLabelFile -FilePath $CompareLabelFile
    $allLabels = @()
    $allNewLabelsToAppend = [System.Collections.Generic.List[string]]::new()
    $missingLabelCount = 0

    $projectFiles = if ($ProjectPath.EndsWith(".sln")) { Get-ProjectsFromSolution -SolutionPath $ProjectPath } else { @($ProjectPath) }
    
    foreach ($projectFile in $projectFiles) {
        Write-Host "`nProcessing project: $projectFile" -ForegroundColor Cyan
        $elements = Get-ElementsFromProject -ProjectPath $projectFile
        foreach ($element in $elements) {
            Write-DebugInfo "Processing element: $($element.Name) - Type (Folder): $($element.Type)"
            $modelFilePath = Find-ElementInModel -ModelPath $ModelPath -ElementName $element.Name -Folder $element.Type
            
            if ($modelFilePath) {
                if ($AutoCreateLabels) {
                    $newlyCreated = Convert-And-Replace-HardcodedText -FilePath $modelFilePath -ObjectName $element.Name -ObjectType $element.Type -ExistingLabels $existingLabels
                    if ($null -ne $newlyCreated) {
                        $allNewLabelsToAppend.AddRange($newlyCreated)
                    }
                }

                $labelsInFile = Extract-LabelsFromXpp -FilePath $modelFilePath -ObjectName $element.Name -ObjectType $element.Type
                if ($null -ne $labelsInFile) {
                    $allLabels += $labelsInFile
                }
            }
            else {
                Write-Warning "File not found for element: $($element.Name) of type $($element.Type)"
            }
        }
    }

    Write-Host "`nTotal label references found: $($allLabels.Count)" -ForegroundColor Green
    if ($allLabels.Count -eq 0 -and $allNewLabelsToAppend.Count -eq 0) {
        Write-Warning "No labels were found or created. Use -ShowDebug for more information."
        return
    }

    if ($FindMissingLabels) {
        $missingLabelCount = Show-MissingLabels -Labels $allLabels -ExistingLabels $existingLabels
    }

    if ($AutoCreateLabels -and $allNewLabelsToAppend.Count -gt 0) {
        $uniqueNewLabels = $allNewLabelsToAppend | Sort-Object | Get-Unique
        Write-Host "`n[APPEND] Adding $($uniqueNewLabels.Count) new labels to file '$CompareLabelFile'..." -ForegroundColor Magenta
        if (Test-Path $CompareLabelFile) {
            $rawContent = Get-Content $CompareLabelFile -Raw -Encoding UTF8 -ErrorAction SilentlyContinue
            if (-not [string]::IsNullOrEmpty($rawContent) -and $rawContent[-1] -ne "`n") {
                Add-Content -Path $CompareLabelFile -Value "" -Encoding UTF8
            }
        }
        $uniqueNewLabels | Add-Content -Path $CompareLabelFile -Encoding UTF8
        Write-Host "[SUCCESS] Label file has been updated." -ForegroundColor Green
    }

    if ($ShowInTerminal) {
        Write-Host "`nExtracted labels (after conversion):" -ForegroundColor Green; Write-Host "=================================" -ForegroundColor Green
        # Combina labels existentes e novas, e então exibe
        $combinedLabels = ($allLabels | Select-Object -ExpandProperty LabelId) + ($allNewLabelsToAppend | ForEach-Object { ($_ -split '=')[0] })
        $combinedLabels | Sort-Object -Unique | ForEach-Object { Write-Host "@STC:$_" }
    }

    Write-Host "`n[FINAL SUMMARY]:" -ForegroundColor Magenta; Write-Host "==============================" -ForegroundColor Magenta
    Write-Host "- Label references found: $($allLabels.Count)" -ForegroundColor White
    if ($AutoCreateLabels) { Write-Host "- New labels created/added: $($allNewLabelsToAppend.Count)" -ForegroundColor Yellow }
    if ($FindMissingLabels) { Write-Host "- Referenced labels not found: $missingLabelCount" -ForegroundColor Red }
}
catch {
    Write-Error "Error during execution: $($_.Exception.Message)"; if ($ShowDebug) { Write-Host $_.Exception.StackTrace -ForegroundColor Red }; exit 1
}