# This script will give the total size of the contents each folder (not including sub folders)
# It will display the results in tabular form.
Param(
    [Parameter(Mandatory=$true)][string]$directoryPath
)

$dirRoot = $directoryPath
#$dirRoot = Read-Host -Prompt 'Please enter the full path '
$filter = '*' # Change this as desired, e.g. *.log or *.txt

if (-Not((get-item $dirRoot).psIsContainer)){
    write-host 'Not a valid directory... Exiting'
    Exit 
}

function GetFolderSize($path){
    $total = (Get-ChildItem $path -ErrorAction SilentlyContinue -filter $filter | Measure-Object -Property length -Sum -ErrorAction SilentlyContinue).Sum
    if (-not($total)) { $total = 0 }
    $total
    } # end function GetFolderSize

# Entry point into script
$results = @()

$dirs = Get-ChildItem $dirRoot -Recurse -ErrorAction SilentlyContinue | Where-Object { $_.psIsContainer}

$result = New-Object psobject -Property @{Folder = (Convert-Path $dirRoot)
                                            TotalSize =  "{0:N2}" -f ((GetFolderSize($dirRoot)) / 1KB)}
$results += $result

foreach ($dir in $dirs) {
   
#    $childFiles = @(Get-ChildItem $dir.pspath -ErrorAction SilentlyContinue -filter $filter| Where-Object{ -not($_.psIsContainer)})
#    if ($childFiles) { $filecount = ($childFiles.count)}
#    else                     { $filecount = 0                  }

    $childDirs = @(Get-ChildItem $dir.pspath -ErrorAction SilentlyContinue | Where-Object{ $_.psIsContainer})
    if ($childDirs ){ $dircount = ($childDirs.count)}
    else                    { $dircount = 0                 }
    
#    $result = New-Object psobject -Property @{Folder = (Split-Path $dir.pspath -NoQualifier)
#                                              TotalSize = (GetFolderSize($dir.pspath))
#                                              FileCount = $filecount; SubDirs = $dircount}
    $result = New-Object psobject -Property @{Folder = (Split-Path $dir.pspath -NoQualifier)
                                              TotalSize =  "{0:N2}" -f ((GetFolderSize($dir.pspath)) / 1KB)}

    
    $results += $result
} # end foreach

$results | Select-Object TotalSize, Folder | Sort-Object TotalSize -Descending | Format-Table -autosize -wrap 

$results | Export-Csv -Path "du_output.csv" -notypeinformation 

Write-Host "Total of $($dirs.count) directories"
