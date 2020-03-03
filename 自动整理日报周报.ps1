# 从config.json文件读取配置信息
$CONF = (Get-Content "config.json") | ConvertFrom-Json
$ExcelFilesFolderDir = $CONF.readFloderPath
$dailyReportListExcelFile = $CONF.dailyReportNameExcel
$weeklyReportListExcelFile = $CONF.weeklyReportNameExcel
$dailyReportPosition= $CONF.dailyReport
$weeklyReportPosition=$CONF.weeklyReport


#下面定义class
Add-Type @'
public class DaliyReportItem    
{
    public int    Index     = 0;   
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
# $ExcelFilesFolderDir = Read-Host "输入要整理的excel表格的文件夹完整路径"
# $ExcelFilesFolderDir = "C:\Users\jeffWu\Documents\yanghaimei\dailyReport\2.26"

# $CurrentPslPath = Split-Path -Parent $MyInvocation.MyCommand.Definition  # 应该和$PWD一样的吧
$CurrentPslPath=$PWD

$dateStr= ($ExcelFilesFolderDir -split "\\")[-1] #取文件夹的名称，代表是那天的日报

$logFile="$CurrentPslPath`\\log_$dateStr.md"

write-output "# 脚本信息日志`n" > $logFile
#判断是否为有效的目录路径
while((Test-Path -Path $ExcelFilesFolderDir -PathType Container) -eq $false)
{
    $ExcelFilesFolderDir = Read-Host "Please enter a valid directory path"
}
$Files = Get-ChildItem -Path $ExcelFilesFolderDir -Filter *.xls?
Write-Host "总共有"$Files.Count"个excel表。"

$ReportType = Read-Host "请问统计的是周报还是日报，日报输入1，周报输入2"
write-output "## 统计错误信息  " >> $logFile
$errorNumber=1
if($ReportType -eq 1)
{
    $reportTypeName="日报"
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
            write-output "$errorNumber`:打不开$File" >> $logFile
            $errorNumber = $errorNumber+1
        }
        else
        {
            # Load the WorkSheet
            $i=$i+1
            $WorkSheet = $WorkBook.Sheets.Item($dailyReportPosition.sheetNumber)
            $DataItem = New-Object DaliyReportItem
            $DataItem.Index=$i
            $DataItem.Name =        $WorkSheet.Range($dailyReportPosition.Name).Text
            $DataItem.Date =        $WorkSheet.Range($dailyReportPosition.Date).Text
            $DataItem.department =  $WorkSheet.Range($dailyReportPosition.department).Text
            $DataItem.position =    $WorkSheet.Range($dailyReportPosition.position).Text
            $DataItem.Project =     $WorkSheet.Range($dailyReportPosition.Project).Text
            $DataItem.report =      $WorkSheet.Range($dailyReportPosition.report).Text
            $DataItem.other =       $WorkSheet.Range($dailyReportPosition.other).Text
            $Data += $DataItem
            $WorkBook.Close()
        }
    }
}
elseif($ReportType -eq 2)
{
    $reportTypeName="周报"
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
            write-output "$errorNumber`:打不开$File" >> $logFile
            $errorNumber = $errorNumber+1
        }
        else
        {
            # Load the WorkSheet
            $i=$i+1
            $WorkSheet = $WorkBook.Sheets.Item($weeklyReportPosition.sheetNumber)
            $DataItem                       = New-Object WeekyReportItem
            $DataItem.Index=$i
            $DataItem.Name                  = $WorkSheet.Range($weeklyReportPosition.Name).Text
            $DataItem.position              = $WorkSheet.Range($weeklyReportPosition.position).Text
            $DataItem.Project               = $WorkSheet.Range($weeklyReportPosition.Project).Text
            $DataItem.Date                  = $WorkSheet.Range($weeklyReportPosition.Date).Text
            $DataItem.projectTimeRatio      = $WorkSheet.Range($weeklyReportPosition.projectTimeRatio).Text
            $DataItem.thisWeekPlan          = $WorkSheet.Range($weeklyReportPosition.thisWeekPlan).Text
            $DataItem.thisWeekReport        = $WorkSheet.Range($weeklyReportPosition.thisWeekReport).Text
            $DataItem.needFeedbackProblem   = $WorkSheet.Range($weeklyReportPosition.needFeedbackProblem).Text
            $DataItem.needHelp              = $WorkSheet.Range($weeklyReportPosition.needHelp).Text
            $DataItem.nextWeekPlan          = $WorkSheet.Range($weeklyReportPosition.nextWeekPlan).Text
            $DataItem.other                 = $WorkSheet.Range($weeklyReportPosition.other).Text
            $Data += $DataItem
            $WorkBook.Close()
        }
    }  
}
if($errorNumber -eq 1)
{
    write-output "没有错误" `n >> $logFile
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

Write-Host "正在统计没交$reportTypeName`的人...."
write-output "## 统计没交$reportTypeName`的人名单  " >> $logFile
# $MemberListExcelFile = Read-Host "输入需要提交日报或周报的人员名单excel文件"
if($ReportType -eq 1)
{
    $MemberListExcelFile=$dailyReportListExcelFile
}
elseif($ReportType -eq 2)
{
    $MemberListExcelFile=$weeklyReportListExcelFile
}
$NameWorkBook=$excel.Workbooks.open($MemberListExcelFile)
$NameWorkSheet=$NameWorkBook.Sheets.Item(1)
$notReportPersonNumber=1
for($i=2;;$i++)
{
    $name=$NameWorkSheet.Range("B$i").Text
    if("" -ne $name)
    {
        if($Data.Name -notcontains $name)
        {
            write-output "$notReportPersonNumber：$name  $reportTypeName`没有交!  " >> $logFile
            $notReportPersonNumber=$notReportPersonNumber+1
        }
    }
    else
    {
        break;    
    }
}
if($notReportPersonNumber -eq 1)
{
    write-output "所有人的$reportType`都交了!  " >> $logFile
}

$excel.Quit()
Write-Host "统计完成！！！！"
exit
# explorer.exe "/Select,$ExportFilePath.xlsx"
