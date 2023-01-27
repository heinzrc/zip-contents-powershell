[CmdletBinding()]
param (
    [Parameter()]
    [string]$SrcPath
)

#Default 7-Zip install path
$7zipPath = "C:\Program Files\7-Zip\7z.exe"
Set-Alias sz $7zipPath

$DesktopPath = [Environment]::GetFolderPath("Desktop")
$CSVPath = "$($DesktopPath)\Output.csv"

#Check if TempDir already exists and delete if so
$TempPath = "$($DesktopPath)\TempDir"
if (Test-Path $TempPath) {
    Remove-Item "$($DesktopPath)\TempDir" -Recurse -Force
}
New-Item -Path $TempPath -ItemType Directory

#Create list of PSObjects for output
$Output = [System.Collections.Generic.List[PSOBJECT]]::new()

function SetOutput {
    param (
        $Name
    )
    $Output.Add([PSCustomObject]@{
        'Filename' = $Name
    })
}

function GetLeaf {
    param (
        $Str
    )
    $Leaf = Split-Path $Str -Leaf
    $LeafArr = $Leaf.Split(".")
    $LeafArr[0]
}

function GetDirContents {
    param (
        $DirPath
    )
    $files = Get-ChildItem -Path $DirPath -Recurse -Force
    foreach ($file in $files) {
        
        $FileShortName = $File.Fullname.Replace($DirPath,"")
        $FileShortName = $FileShortName.Split(".")
        $ArchivePath = "$($DirPath)$($FileShortName[0]).$($FileShortName[1])"

        if ($file.Extension -like "*.zip") {
            Expand-Archive -Path $ArchivePath -DestinationPath $TempPath -Force
            SetOutput -Name $File.name
            
            GetDirContents -DirPath "$($TempPath)$($FileShortName[0])"

        } elseif ($file.Extension -like "*.7z") {
            sz x -o"$TempPath" $ArchivePath
            SetOutput -Name $File.name
            
            GetDirContents -DirPath "$($TempPath)$($FileShortName[0])"

        } elseif ($file.Attributes -eq "Directory") {
            Copy-Item -Path $ArchivePath -Destination $TempPath -Recurse -Force

            SetOutput -Name $File.name

            GetDirContents -DirPath "$($TempPath)$($FileShortName[0])"

        } else {

            SetOutput -Name $File.name
        }
    } 
}

if ($SrcPath -like "*.zip") {
    Expand-Archive -Path $SrcPath -DestinationPath $TempPath
    GetDirContents -DirPath "$TempPath\$(GetLeaf -Str $SrcPath)"

} elseif ($SrcPath -like "*.7z") {
    sz x -o"$TempPath" "$SrcPath"
    GetDirContents -DirPath "$TempPath\$(GetLeaf -Str $SrcPath)"

} else {

    GetDirContents -DirPath $SrcPath
}

#Get Rid of duplicates
$SortedOutput = $Output | Sort-Object -Unique -Property "Filename"

$SortedOutput | Export-Csv $CSVPath -NoTypeInformation
Invoke-Item $CSVPath
Stop-Process -Id $PID