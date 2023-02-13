[CmdletBinding()]
param (
    [Parameter()]
    [string]$SourcePath
)

#Default 7-Zip install path
$7zipPath = "C:\Program Files\7-Zip\7z.exe"
Set-Alias 7zip $7zipPath

$DesktopPath = [Environment]::GetFolderPath("Desktop")
$CSVPath = "$($DesktopPath)\Output.csv"

#Check if TempDir already exists and delete if so
$TempDirectory = "$($DesktopPath)\TempDir"
if (Test-Path $TempDirectory) {
    Remove-Item $TempDirectory -Recurse -Force
}
New-Item -Path $TempDirectory -ItemType Directory

#Create list of PSObjects for output
$Output = [System.Collections.Generic.List[PSOBJECT]]::new()

function SetOutput ($Name) {
    $Output.Add([PSCustomObject]@{
            'Filename' = $Name
        })
}

function GetLeaf ($Path){
    $Leaf = Split-Path $Path -Leaf
    $LeafArr = $Leaf.Split(".")
    $LeafArr[0]
}

#Start Recursive Function
function GetDirContents ($DirPath){
    $files = Get-ChildItem -Path $DirPath -Recurse -Force

    foreach ($file in $files) {

        if ($file.Extension -like "*.zip") {
            Expand-Archive -Path $File.FullName -DestinationPath $TempDirectory -Force
            SetOutput -Name $File.name
    
            GetDirContents -DirPath "$($TempDirectory)\$($File.BaseName)"
    
        }
        elseif ($file.Extension -like "*.7z") {
            7zip x -y -o"$TempDirectory" $File.FullName
            SetOutput -Name $File.name
                
            GetDirContents -DirPath "$($TempDirectory)\$($File.BaseName)"
    
        }
        else {
            SetOutput -Name $File.name
        }
        
    } 
}
#End Recursive Function

#Start Input Type Check
if ($SourcePath -like "*.zip") {
    Expand-Archive -Path $SourcePath -DestinationPath $TempDirectory 
    if (Test-Path "$TempDirectory\$(GetLeaf -Path $SourcePath)") {
        GetDirContents -DirPath "$TempDirectory\$(GetLeaf -Path $SourcePath)"
    }
    else {
        GetDirContents -DirPath "$TempDirectory"
    }

}
elseif ($SourcePath -like "*.7z") {
    7zip x -y -o"$TempDirectory" "$SourcePath"
    if (Test-Path "$TempDirectory\$(GetLeaf -Path $SourcePath)") {
        GetDirContents -DirPath "$TempDirectory\$(GetLeaf -Path $SourcePath)"
    }
    else {
        GetDirContents -DirPath "$TempDirectory"
    }

}
else {
    #Just a folder
    GetDirContents -DirPath $SourcePath
}
#End Input Type Check

#Get Rid of duplicates
$SortedOutput = $Output | Sort-Object -Unique -Property "Filename"

$SortedOutput | Export-Csv $CSVPath -NoTypeInformation
Invoke-Item $CSVPath
Stop-Process -Id $PID