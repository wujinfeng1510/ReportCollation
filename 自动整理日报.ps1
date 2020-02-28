#输入目录路径
$WorkingDir = Read-Host "Please enter a valid directory path"
# $WorkingDir = "C:\Users\jeffWu\Documents\yanghaimei\dailyReport\2.26"
$CurrentPslPath = Split-Path -Parent $MyInvocation.MyCommand.Definition
$dateStr= ($WorkingDir -split "\\")[-1]

write-output "错误信息" > "$CurrentPslPath`\\message.txt"
#判断是否为有效的目录路径
while((Test-Path -Path $WorkingDir -PathType Container) -eq $false)
{
    $WorkingDir = Read-Host "Please enter a valid directory path"
}
$Files = Get-ChildItem -Path $WorkingDir -Filter *.xls?
Write-Host "总共有"$Files.Count"个excel表。"

$Data = @() #定义Data为数组

#下面定义class
Add-Type @'
public class ProcessDataItem    
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

        $DataItem = New-Object ProcessDataItem
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
$objExcel.Quit()
Stop-Process -Name excel

# $ExportFilePath = Join-Path -Path $WorkingDir -ChildPath "result.csv"

$ExportFilePath = Join-Path -Path $CurrentPslPath -ChildPath "$dateStr`日报的统计结果"
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
write-output "`统计没交日报的人名单" >> "$CurrentPslPath`\\message.txt"
$NameWorkBook=$excel.Workbooks.open("C:\Users\wujf_nuc\yanghaimei\日报名单.xlsx")
$NameWorkSheet=$NameWorkBook.Sheets.Item(1)
for($i=2;;$i++)
{
    $name=$NameWorkSheet.Range("B$i").Text
    if("" -ne $name)
    {
        if($Data.Name -notcontains $name)
        {
            write-output "$name  的日报没有交!" >> "$CurrentPslPath`\\message.txt"
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
