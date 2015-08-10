<#
.SYNOPSIS
Checks the space of a disk storing the results in a file

.PARAMETER DirectoryPath
Path of the Directory space you would like to measure. Defaults to C:\

.PARAMETER ResultFilePath
Path of the file you would like to store the results in. Defaults to Public\Documents\FolderData.csv

.Example
./Check-DiskSpace.ps1
Checks the disk space of C drive and stores the result in Public\Documents\FolderData.csv


#>
[CmdletBinding()]
Param(
    [Parameter()]
    [String] $DirectoryPath = "C:\",

    [Parameter()]
    [String] $ResultFilePath = "$env:PUBLIC\Documents\FolderData.csv"
)

Begin{
$assemblies = ('System', 'mscorlib', 'System.IO');
$sourceCode = @'
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace DiskSpaceChecker
{
    public class DiskChecker
    {

        public long DirectorySize(DirectoryInfo dInfo, int parentId)
        {
            // Enumerate all the files
            long totalSize = 0;
            long fileCount = 0;
            

            var mainFolderTracker = new IdTracking()
            {
                Path = dInfo.FullName,
                ParentId = parentId
            };

            foreach (var file in dInfo.EnumerateFiles())
            {
                IdTracking fileTracker = null;

                try
                {
                    totalSize += file.Length;
                    const long sizeCompare = 1048576 * 100;

                    if (file.Length >= (sizeCompare))
                    {
                        fileTracker = new IdTracking()
                        {
                            Path = file.FullName,
                            ParentId = mainFolderTracker.Id
                        };

                        RecordData(fileTracker.Id, mainFolderTracker.Id, file.Name, file.FullName, "File", file.Length, 0, 0, true);
                    }
                }
                catch (Exception)
                {
                    try
                    {
                        if (fileTracker == null)
                        {
                            fileTracker = new IdTracking()
                            {
                                Path = file.FullName,
                                ParentId = mainFolderTracker.Id
                            };
                        }

                        RecordData(fileTracker.Id, mainFolderTracker.Id, file.Name, file.FullName, "File", 0, 0, 0, false);
                    }
                    catch (Exception)
                    {

                    }
                }

                fileCount += 1;
            }

            var tmpTotalSize = totalSize;
            long subDirectoriesSize = 0;

            foreach (var directory in dInfo.EnumerateDirectories())
            {
                try
                {
                    var tmpDirectorySize = DirectorySize(directory, mainFolderTracker.Id);

                    subDirectoriesSize += tmpDirectorySize;
                    totalSize += tmpDirectorySize;
                }
                catch (Exception)
                {
                    var folderTracker = new IdTracking()
                    {
                        Path = directory.FullName,
                        ParentId = mainFolderTracker.Id
                    };

                    RecordData(folderTracker.Id, folderTracker.ParentId, directory.Name, directory.FullName, "Folder", 0, 0, 0, false);
                }
            }

            RecordData(mainFolderTracker.Id, mainFolderTracker.ParentId, dInfo.Name, dInfo.FullName, "Folder", tmpTotalSize, subDirectoriesSize, fileCount, true);

            return totalSize;
        }

        private void RecordData(int id, int? parentId, string name, string path, string type, long totalSize, long directorySize, long fileCount, bool accessAllowed)
        {
            var dirSizeInfo = new DirectorySizeInfo()
            {
                Id = id,
                ParentId = parentId,
                Name = name,
                Path = path,
                Type = type,
                FileSize = totalSize,
                FileCount = fileCount,
                TotalSize = totalSize,
                DirectorySize = directorySize,
                FreeSpace = 0,
                AccessAllowed = accessAllowed
            };

            DirectorySizeInfos.Add(dirSizeInfo);
        }

        public DriveInfo GetTotalFreeSpace(string driveName)
        {
            foreach (var drive in DriveInfo.GetDrives())
            {
                if (drive.IsReady && drive.Name == driveName)
                {
                    return drive;
                }
            }
            return null;
        }

        public void Reset()
        {
            DirectorySizeInfos = new List<DirectorySizeInfo>();
            IdTracker = 0;
            IdTrackers = new List<IdTracking>();
        }

        public List<DirectorySizeInfo> DirectorySizeInfos = new List<DirectorySizeInfo>();

        
        public class IdTracking
        {
            public IdTracking()
            {
                IdTracker++;
                Id = IdTracker;
            }

            public string Path { get; set; }

            public int Id { get; set; }

            public int ParentId { get; set; }
        }

        public static int IdTracker = 0;

        public List<IdTracking> IdTrackers = new List<IdTracking>();
    }


    public class DriveSpace
    {
        public long TotalSpace { get; set; }

        public long FreeSpace { get; set; }
    }

    public class DirectorySizeInfo
    {
        public int Id { get; set; }

        public int? ParentId { get; set; }

        public bool AccessAllowed { get; set; }

        public string Name { get; set; }

        public string Type { get; set; }

        public string Path { get; set; }

        public long FileSize { get; set; }

        public long TotalSize { get; set; }

        public long DirectorySize { get; set; }

        public long FreeSpace { get; set; }

        public long FileCount { get; set; }
    }

    
}
'@
}

Process{
    Write-Host "     Creating DiskChecker Object"
	Add-Type -TypeDefinition $sourceCode -ReferencedAssemblies $assemblies -ErrorAction STOP;
    $checker = New-Object DiskSpaceChecker.DiskChecker
    Write-Host "     Getting Directory Info"
    $dInfo = New-Object System.IO.DirectoryInfo $DirectoryPath
    Write-Host "     Checking Used Disk Space. This may take several minutes..."
    $checker.DirectorySize($dInfo, 0) | Out-Null
    Write-Host "     Getting Free Space"
    $DirectoryDrive = $checker.DirectorySizeInfos | ? Name -EQ $DirectoryPath
    $Directoryinfo = $checker.GetTotalFreeSpace($DirectoryPath)
    if($Directoryinfo -ne $null){
        $DirectoryDrive.TotalSize = $Directoryinfo.TotalSize
        $DirectoryDrive.FreeSpace = $Directoryinfo.TotalFreeSpace
    }
    Write-Host "     Outputting to $ResultFilePath"
    $checker.DirectorySizeInfos | Export-Csv $ResultFilePath -NoTypeInformation
    $checker.Reset();
    Write-Host "     Process Complete"
}