#Check Local Disk Space

###Pre-requisites

Excel 2016

###Examples

1. Download Excel 2016 New Charts.docx and Excel Template.xlsx. 

          The Excel 2016 New Charts.docx will explain the process of creating a Treemap or Sunburst chart in Excel.
          
          The Excel Template.xlsx is the template containing the PivotTable.  

2. Open a PowerShell console.

          From the Run dialog type PowerShell
  
3. Change the directory to the location where the PowerShell Script is saved.

          Example: cd C:\PowerShellScripts
  
4. Run the Script.

          Type . .\Check-DiskSpace.ps1

          By including the additional period before the relative script path you are 'Dot-Sourcing' 
          the PowerShell function in the script into your PowerShell session which will allow you to 
          run the function 'Get-ModernOfficeApps' from the console
           
5. A folder called FolderData.csv containing the results will be copied to Public\Documents.

6. Open Excel Template.xlsx. Click the Data Tab and under Connections choose Refresh All.

7. Follow the instructions in Excel 2016 New Charts.docx to insert the data into a Treemap or Sunburst chart.
           
