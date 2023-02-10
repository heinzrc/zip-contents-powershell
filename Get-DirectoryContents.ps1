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

#Start Recursive Function
function GetDirContents {
    param (
        $DirPath
    )
    "Test: $DirPath"
    $files = Get-ChildItem -Path $DirPath -Recurse -Force

    foreach ($file in $files) {
        
        $FileShortName = $File.Fullname.Replace($DirPath, "")
        $FileShortName = $FileShortName -Split '\.', 2 #Accounts for files with multiple file extensions
        $ArchivePath = "$($DirPath)$($FileShortName[0]).$($FileShortName[1])"
        $FileShortName[0]
        if (Test-Path $ArchivePath) {
            if ($file.Extension -like "*.zip") {
                Expand-Archive -Path $ArchivePath -DestinationPath $TempPath -Force
                SetOutput -Name $File.name
    
                GetDirContents -DirPath "$($TempPath)$($FileShortName[0])"
    
            }
            elseif ($file.Extension -like "*.7z") {
                sz x -y -o"$TempPath" $ArchivePath 
                SetOutput -Name $File.name
                
                GetDirContents -DirPath "$($TempPath)$($FileShortName[0])"
    
            }
            elseif ($file.Attributes -eq "Directory") {
    
                SetOutput -Name $File.name
    
                GetDirContents -DirPath "$($DirPath)$($FileShortName[0])"
    
            }
            else {
    
                SetOutput -Name $File.name
            }
        } 
    } 
}
#End Recursive Function

#Start Input Type Check
if ($SrcPath -like "*.zip") {
    Expand-Archive -Path $SrcPath -DestinationPath $TempPath 
    if (Test-Path "$TempPath\$(GetLeaf -Str $SrcPath)") {
        GetDirContents -DirPath "$TempPath\$(GetLeaf -Str $SrcPath)"
    }
    else {
        GetDirContents -DirPath "$TempPath"
    }

}
elseif ($SrcPath -like "*.7z") {
    sz x -y -o"$TempPath" "$SrcPath"
    if (Test-Path "$TempPath\$(GetLeaf -Str $SrcPath)") {
        GetDirContents -DirPath "$TempPath\$(GetLeaf -Str $SrcPath)"
    }
    else {
        GetDirContents -DirPath "$TempPath"
    }

}
else {
    #Just a folder

    GetDirContents -DirPath $SrcPath
}
#End Input Type Check

#Get Rid of duplicates
$SortedOutput = $Output | Sort-Object -Unique -Property "Filename"

$SortedOutput | Export-Csv $CSVPath -NoTypeInformation
Invoke-Item $CSVPath
Stop-Process -Id $PID