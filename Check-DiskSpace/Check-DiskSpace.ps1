#Global Variables
[int] $global:IdTrackerNum = 0;
$global:directorySizeInfos = New-Object PSObject[] 1
#functions
function Get-DirectorySize
{
    Param(
    [Parameter()]
    [System.IO.DirectoryInfo] $dInfo, 
        
    [Parameter()]
    [int] $parentId
    )
        [long] $totalSize = 0;
        [long] $fileCount = 0;
            

        $mainFolderTracker = New-IdTracker -Path $dInfo.FullName -ParentId $parentId

        foreach ($file in $dInfo.EnumerateFiles())
        {
            $fileTracker = $null;

            try
            {
                $totalSize += $file.Length;
                [long] $sizeCompare = 1048576 * 100;

                if ($file.Length -ge ($sizeCompare))
                {
                    $fileTracker = New-IdTracker -Path $file.FullName -ParentId $mainFolderTracker.Id

                    Record-Data -Id $fileTracker.Id -ParentId $mainFolderTracker.Id -Name $file.Name -Path $file.FullName -Type "File" -TotalSize $file.Length -DirectorySize 0 -FileCount 0 -AccessAllowed $true
                }
            }
            catch
            {
                try
                {
                    if ($fileTracker -eq $null)
                    {
                        $fileTracker = New-IdTracker -Path $file.FullName -ParentId $mainFolderTracker.Id
                    }
                    Record-Data -Id $fileTracker.Id -ParentId $mainFolderTracker.Id -Name $file.Name -Path $file.FullName -Type "File" -TotalSize 0 -DirectorySize 0 -FileCount 0 -AccessAllowed $false
                }
                catch
                {

                }
            }

            $fileCount += 1;
        }

        $tmpTotalSize = $totalSize;
        [long] $subDirectoriesSize = 0;

        foreach ($directory in $dInfo.EnumerateDirectories())
        {
            try
            {
                $tmpDirectorySize = Get-DirectorySize -dInfo $directory -parentId $mainFolderTracker.Id

                $subDirectoriesSize += $tmpDirectorySize;
                $totalSize += $tmpDirectorySize;
            }
            catch
            {
                $folderTracker = New-IdTracker -ParentId $mainFolderTracker.Id -Path $directory.FullName

                Record-Data -Id $folderTracker.Id -ParentId $folderTracker.ParentId -Name $directory.Name -Path $directory.FullName -Type "Folder" -TotalSize 0 -DirectorySize 0 -FileCount 0 -AccessAllowed $false
            }
        }
        Record-Data -Id $mainFolderTracker.Id -ParentId $mainFolderTracker.ParentId -Name $dInfo.Name -Path $dInfo.FullName -Type "Folder" -TotalSize $tmpTotalSize -DirectorySize $subDirectoriesSize -FileCount $fileCount -AccessAllowed $true

        return $totalSize;
}

function Record-Data
{
    Param(

    [Parameter()]
    [int] $Id,

    [Parameter()]
    [int] $ParentId,

    [Parameter()]
    [string] $Name,

    [Parameter()]
    [string] $Path,

    [Parameter()]
    [string] $Type,

    [Parameter()]
    [long] $TotalSize,

    [Parameter()]
    [long] $DirectorySize,

    [Parameter()]
    [long] $FileCount,

    [Parameter()]
    [bool] $AccessAllowed

    )
    Process
    {
        $dirSizeInfo = New-Object PSObject -Property @{Id = $Id;
                ParentId = $ParentId;
                Name = $Name;
                Path = $Path;
                Type = $Type;
                FileSize = $TotalSize;
                FileCount = $FileCount;
                DirectorySize = $DirectorySize;
                AccessAllowed = $AccessAllowed}

        if($global:directorySizeInfos -ne $null){
            $global:directorySizeInfos += $dirSizeInfo;
        }else{
            $global:directorySizeInfos[0] = $dirSizeInfo; 
        }
    }
}

function New-IdTracker
{
    Param(

    [Parameter()]
    [int] $ParentId,

    [Parameter()]
    [string] $Path

    )

    Process
    {
        $global:IdTrackerNum += 1;
        return New-Object PSObject -Property @{ Id=$global:IdTrackerNum; ParentId=$ParentId; Path=$Path }; 
    }
}

function Write-Data
{
    Process
    {
        $drvSpace = Get-TotalFreeSpace -DriveName "C:\\"
        try{
            $loc = Get-Location
            $locPath = $loc.Path;
            $sw = New-Object System.IO.StreamWriter -ArgumentList "$locPath\FolderData.txt"
            $drvEntry = $global:directorySizeInfos | ? {([string]($_.Name)).ToUpper() -eq "C:\" };

            $sw.WriteLine("Id\tParentId\tName\tType\tPath\tFileCount\tFilesSize\tSubDirectorySize\tTotalSize\tFreeSpace\tAccessAllowed");

            if ($drvEntry -ne $null)
            {
                $sw.WriteLine("$($drvEntry.Id)\t$($drvEntry.ParentId)\t$($drvEntry.Name)\tDrive\t$($drvEntry.Path)\t$($drvEntry.FileCount)\t$($drvEntry.FileSize)\t$($drvEntry.DirectorySize)\t$($drvSpace.TotalSize)\t$($drvSpace.TotalFreeSpace)\ttrue");
            }
            $drvEntrys = $global:directorySizeInfos | ? {$_.Name.ToUpper() -ne "C:\" };
            foreach ($d in $drvEntrys)
            {
                $sw.WriteLine("$($d.Id)\t$($d.ParentId)\t$($d.Name)\t$($d.Type)\t$($d.Path)\t$($d.FileCount)\t$($d.FileSize)\t$($d.DirectorySize)\t$($d.FileSize + $d.DirectorySize)\t0\t$($d.AccessAllowed)");
                [System.Threading.Thread]::Sleep(1);
            }

            $sw.Flush();
            $sw.Close();
        }
        finally{
            $sw.Dispose();
        }
    }
}

function Get-TotalFreeSpace
{

    Param(

    [Parameter()]
    [string] $DriveName

    )

    Process
    {
        foreach($drive in [System.IO.DriveInfo]::GetDrives())
        {
            if($drive.IsReady -and ($drive.Name -eq $DriveName))
            {
                return $drive;
            }
        }
        return $null;
    }

}

$dInfo = New-Object System.IO.DirectoryInfo "C:\"
$sizeOfDir = Get-DirectorySize -dInfo $dInfo -parentId 0;

Write-Data;