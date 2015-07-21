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
    [String] $ResultFilePath = "$env:PUBLIC\Documents\FolderData.xlsx",

    [Parameter()]
    [String] $ExcelSourcePath = "$env:PUBLIC\Documents\ExcelTemplate.xlsx"
)

Begin{
$assemblies = ('System', 'mscorlib', 'System.IO', 'Microsoft.Office.Interop.Excel');
$sourceCode = @'
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;

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

        public void WriteData(string DestinationFilePath, string SourceFilePath)
        {
            var drvSpace = GetTotalFreeSpace("C:\\");
            var xlApp = new Microsoft.Office.Interop.Excel.Application();
            var xlWorkBook = xlApp.Workbooks.Open(SourceFilePath);
            var dataSheet = (Microsoft.Office.Interop.Excel.Worksheet)(xlWorkBook.Worksheets.get_Item(4));
            dataSheet.Cells.Clear();
            var drvEntrys = DirectorySizeInfos.Where(d => d.Name.ToUpper() == "C:\\");
            var drvEntry = drvEntrys.FirstOrDefault();
            //first line entry
            dataSheet.Cells[1, 1] = "Id";
            dataSheet.Cells[1, 2] = "ParentId";
            dataSheet.Cells[1, 3] = "Name";
            dataSheet.Cells[1, 4] = "Type";
            dataSheet.Cells[1, 5] = "Path";
            dataSheet.Cells[1, 6] = "FileCount";
            dataSheet.Cells[1, 7] = "FileSize";
            dataSheet.Cells[1, 8] = "SubDirectorySize";
            dataSheet.Cells[1, 9] = "TotalSize";
            dataSheet.Cells[1, 10] = "FreeSpace";
            dataSheet.Cells[1, 11] = "AccessAllowed";
            if (drvEntry != null)
            {
                dataSheet.Cells[2, 1] = drvEntry.Id;
                dataSheet.Cells[2, 2] = drvEntry.ParentId;
                dataSheet.Cells[2, 3] = drvEntry.Name;
                dataSheet.Cells[2, 4] = "Drive";
                dataSheet.Cells[2, 5] = drvEntry.Path;
                dataSheet.Cells[2, 6] = drvEntry.FileCount;
                dataSheet.Cells[2, 7] = drvEntry.FileSize;
                dataSheet.Cells[2, 8] = drvEntry.DirectorySize;
                dataSheet.Cells[2, 9] = drvSpace.TotalSize;
                dataSheet.Cells[2, 10] = drvSpace.TotalFreeSpace;
                dataSheet.Cells[2, 11] = "true";
            }

            long row = 3;
            foreach (var d in DirectorySizeInfos.Where(d => d.Name.ToUpper() != "C:\\"))
            {
                dataSheet.Cells[row, 1] = d.Id;
                dataSheet.Cells[row, 2] = d.ParentId;
                dataSheet.Cells[row, 3] = d.Name;
                dataSheet.Cells[row, 4] = d.Type;
                dataSheet.Cells[row, 5] = d.Path;
                dataSheet.Cells[row, 6] = d.FileCount;
                dataSheet.Cells[row, 7] = d.FileSize;
                dataSheet.Cells[row, 8] = d.DirectorySize;
                dataSheet.Cells[row, 9] = (d.FileSize + d.DirectorySize);
                dataSheet.Cells[row, 10] = 0;
                dataSheet.Cells[row, 11] = d.AccessAllowed;
                row++;
            }
            xlWorkBook.SaveAs(DestinationFilePath);
            xlWorkBook.Close();
            xlApp.Quit();
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
    $checker.WriteData($ResultFilePath, $ExcelSourcePath);
}