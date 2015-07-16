<#
.SYNOPSIS
Checks the space of a disk storing the results in a file

.PARAMETER DirectoryPath
Path of the Directory space you would like to measure. Defaults to C:\

.PARAMETER ResultFilePath
Path of the file you would like to store the results in. Defaults to CurrentDirectory\FolderData.txt

.Example
./Check-DiskSpace.ps1
Checks the disk space of C drive and stores the result in CurrentDirectory\FolderData.txt


#>
[CmdletBinding()]
Param(
    [Parameter()]
    [String] $DirectoryPath = "C:\",

    [Parameter()]
    [String] $ResultFilePath = "$((Get-Location).Path)\FolderData.txt"
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
                DirectorySize = directorySize,
                AccessAllowed = accessAllowed
            };

            DirectorySizeInfos.Add(dirSizeInfo);
        }

        public void WriteData(string fullFilePath)
        {
            var drvSpace = GetTotalFreeSpace("C:\\");

            using (var sw = new StreamWriter(fullFilePath))
            {
                var drvEntrys = DirectorySizeInfos.Where(d => d.Name.ToUpper() == "C:\\");
                var drvEntry = drvEntrys.FirstOrDefault();

                sw.WriteLine("Id\tParentId\tName\tType\tPath\tFileCount\tFilesSize\tSubDirectorySize\tTotalSize\tFreeSpace\tAccessAllowed");

                if (drvEntry != null)
                {
                    sw.WriteLine(drvEntry.Id + "\t" + drvEntry.ParentId + "\t" + drvEntry.Name + "\t" + "Drive" + "\t" +
                                 drvEntry.Path + "\t" +
                                 drvEntry.FileCount + "\t" + drvEntry.FileSize + "\t" + drvEntry.DirectorySize + "\t" +
                                 drvSpace.TotalSize + "\t" + drvSpace.TotalFreeSpace + "\t" + "true");
                }

                foreach (var d in DirectorySizeInfos.Where(d => d.Name.ToUpper() != "C:\\"))
                {
                    sw.WriteLine(d.Id + "\t" + d.ParentId + "\t" + d.Name + "\t" + d.Type + "\t" + d.Path + "\t" +
                        d.FileCount + "\t" + d.FileSize + "\t" + d.DirectorySize + "\t" + (d.FileSize + d.DirectorySize) + "\t0\t" + d.AccessAllowed);
                }

                sw.Flush();
                sw.Close();
            }
        }

        private DriveInfo GetTotalFreeSpace(string driveName)
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

        public long FileCount { get; set; }
    }

    
}
'@
}

Process{

	Add-Type -TypeDefinition $sourceCode -ReferencedAssemblies $assemblies -ErrorAction STOP;
    $checker = New-Object DiskSpaceChecker.DiskChecker
    $dInfo = New-Object System.IO.DirectoryInfo $DirectoryPath
    $checker.DirectorySize($dInfo, 0);
    $checker.WriteData($ResultFilePath);
}