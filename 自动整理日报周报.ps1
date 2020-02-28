
#下面定义class
Add-Type @'
public class DaliyReportItem    
{
    public int Index     = 0;   
    public string Name      = "";   
    public string Date  = ""; 
    public string department      = "";  
    public string position     = "";
    public string Project     = "";
    public string report    = "";
    public string other    = "";
}
'@

#下面定义class
Add-Type @"
public class WeekyReportItem    
{
    public int    Index                 = 0; 
    public string Name                  = "";   
    public string position              = "";
    public string Project               = "";
    public string Date                  = ""; 
    public string projectTimeRatio      = "";
    public string thisWeekPlan          = "";
    public string thisWeekReport        = "";
    public string needFeedbackProblem   = "";
    public string needHelp              = "";
    public string nextWeekPlan          = "";
    public string other                 = "";
}
"@

#输入目录路径
$ExcelFilesFolderDir = Read-Host "输入要整理的excel表格的文件夹完整路径"
# $ExcelFilesFolderDir = "C:\Users\jeffWu\Documents\yanghaimei\dailyReport\2.26"

# $CurrentPslPath = Split-Path -Parent $MyInvocation.MyCommand.Definition  # 应该和$PWD一样的吧
$CurrentPslPath=$PWD

$dateStr= ($ExcelFilesFolderDir -split "\\")[-1] #取文件夹的名称，代表是那天的日报

$logFile="$CurrentPslPath`\\log_$dateStr.md"

write-output "#脚本信息日志`n" > $logFile
#判断是否为有效的目录路径
while((Test-Path -Path $ExcelFilesFolderDir -PathType Container) -eq $false)
{
    $ExcelFilesFolderDir = Read-Host "Please enter a valid directory path"
}
$Files = Get-ChildItem -Path $ExcelFilesFolderDir -Filter *.xls?
Write-Host "总共有"$Files.Count"个excel表。"

$ReportType = Read-Host "请问统计的是周报还是日报，日报输入1，周报输入2"
if($ReportType -eq 1)
{
    $Data = @() #定义Data为数组
    # Create an Object Excel.Application using Com interface
    $objExcel = New-Object -ComObject Excel.Application
    # Disable the 'visible' property so the document won't open in excel
    $objExcel.Visible = $false
    $i=0
    foreach ($File in $Files)
    {
        Write-Host $File
        #Specify the path of the excel file
        $FilePath = $File.FullName
        # Open the Excel file(ReadOnly mode) and save it in $WorkBook
        $WorkBook = $objExcel.Workbooks.Open($FilePath,$true)
        # $WorkBook.FullName|Write-Host
        if($null -eq $WorkBook.FullName)
        {
            write-output "打不开$File" >> $logFile
        }
        else
        {
            # Load the WorkSheet
            $i=$i+1
            $WorkSheet = $WorkBook.Sheets.Item(1)
            $DataItem = New-Object DaliyReportItem
            $DataItem.Index=$i
            $DataItem.Name =        $WorkSheet.Range("B6").Text
            $DataItem.Date =        $WorkSheet.Range("A6").Text
            $DataItem.department =  $WorkSheet.Range("C6").Text
            $DataItem.position =    $WorkSheet.Range("D6").Text
            $DataItem.Project =     $WorkSheet.Range("E6").Text
            $DataItem.report =      $WorkSheet.Range("F6").Text
            $DataItem.other =       $WorkSheet.Range("G6").Text
            $Data += $DataItem
            $WorkBook.Close()
        }
    }
}
elseif($ReportType -eq 2)
{
    $Data = @() #定义Data为数组
    # Create an Object Excel.Application using Com interface
    $objExcel = New-Object -ComObject Excel.Application
    # Disable the 'visible' property so the document won't open in excel
    $objExcel.Visible = $false
    $i=0
    foreach($File in $Files)
    {
        Write-Host $File
        #Specify the path of the excel file
        $FilePath = $File.FullName
        # Open the Excel file(ReadOnly mode) and save it in $WorkBook
        $WorkBook = $objExcel.Workbooks.Open($FilePath, $true)
        # $WorkBook.FullName|Write-Host
        if($null -eq $WorkBook.FullName)
        {
            write-output "打不开$File" >> "$CurrentPslPath`\\message.txt"
        }
        else
        {
            # Load the WorkSheet
            $i=$i+1
            $WorkSheet = $WorkBook.Sheets.Item(1)
            $DataItem                       = New-Object WeekyReportItem
            $DataItem.Index=$i
            $DataItem.Name                  = $WorkSheet.Range("C2").Text
            $DataItem.position              = $WorkSheet.Range("E2").Text
            $DataItem.Project               = $WorkSheet.Range("C3").Text
            $DataItem.Date                  = $WorkSheet.Range("E4").Text
            $DataItem.projectTimeRatio      = $WorkSheet.Range("E3").Text
            $DataItem.thisWeekPlan          = $WorkSheet.Range("C5").Text
            $DataItem.thisWeekReport        = $WorkSheet.Range("C6").Text
            $DataItem.needFeedbackProblem   = $WorkSheet.Range("C7").Text
            $DataItem.needHelp              = $WorkSheet.Range("C8").Text
            $DataItem.nextWeekPlan          = $WorkSheet.Range("C9").Text
            $DataItem.other                 = $WorkSheet.Range("C10").Text
            $Data += $DataItem
            $WorkBook.Close()
        }
    }  
}
$objExcel.Quit()
Stop-Process -Name excel

$ExportFilePath = Join-Path -Path $CurrentPslPath -ChildPath "$dateStr`文件夹的统计结果"
$Data | Export-Csv "$ExportFilePath.csv" -NoTypeInformation -Encoding UTF8

if(Test-Path "$ExportFilePath.xlsx")
{
    Remove-Item "$ExportFilePath.xlsx"  # 输出原来的文件
}
# load into Excel
$excel = New-Object -ComObject Excel.Application 
# $excel.Visible = $true
# $excel.Workbooks.Open("$ExportFilePath.csv").SaveAs("$ExportFilePath.xlsx",51, [Type]::Missing, [Type]::Missing, $false, $false, 1, 2)
$excel.Workbooks.Open("$ExportFilePath.csv").SaveAs("$ExportFilePath.xlsx",51)

Write-Host "正在统计没交日报的人...."
write-output "`统计没交日报的人名单" >> $logFile
$MemberListExcelFile = Read-Host "输入需要提交日报或周报的人员名单excel文件"
$NameWorkBook=$excel.Workbooks.open($MemberListExcelFile)
$NameWorkSheet=$NameWorkBook.Sheets.Item(1)
for($i=2;;$i++)
{
    $name=$NameWorkSheet.Range("B$i").Text
    if("" -ne $name)
    {
        if($Data.Name -notcontains $name)
        {
            write-output "$name  的日报没有交!" >> $logFile
        }
    }
    else
    {
        break;    
    }
}

$excel.Quit()
Write-Host "统计完成！！！！"
exit
# explorer.exe "/Select,$ExportFilePath.xlsx"
