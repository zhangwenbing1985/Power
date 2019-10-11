if((Get-PSSnapin -Name VeeamPSSnapIn -ErrorAction SilentlyContinue) -eq $null){
    Add-PSSnapin "VeeamPSSnapIn"
}
Disconnect-VBRServer
Connect-VBRServer -Server "172.30.0.111" -User "ap\cn11-dataprotector" -Password ""

#<#
#调用EXCEL
$excel = new-object -comobject excel.application
$workbook = $excel.workbooks.add()
$workbook.worksheets.item(1).name = "products"
$sheet = $workbook.worksheets.item("products")
$excel.visible = $true


#定义枚举类型
$linestyle = "microsoft.office.interop.excel.xllinestyle" -as [type]
$colorstyle = "microsoft.office.interop.excel.xlcolorstyle" -as [type]
$borderWeight = "microsoft.office.interop.excel.xlBorderWeight" -as [type]
$chartstyle = "microsoft.office.interop.excel.xlchartstyle" -as [type]

for($b = 1 ; $b -le 5 ; $b++)
    {
     $sheet.cells.item(1,$b).font.bold = $true
     $sheet.cells.item(1,$b).borders.LineStyle = $lineStyle::xlDashDot
     $sheet.cells.item(1,$b).borders.ColorIndex = $colorstyle::xlColorIndexAutomatic
     $sheet.cells.item(1,$b).borders.weight = $borderWeight::xlMedium
    }
    
$sheet.cells.item(1,1) = "Jobname"
$sheet.cells.item(1,2) = "VMname"
$sheet.cells.item(1,3) = "Size"
$sheet.cells.item(1,4) = "Algorithm"
$sheet.cells.item(1,5) = "RetainCycles"
$sheet.cells.item(1,6) = "Proxy"
$x=2
#>

$JobInfo = @()
foreach($vbrJob in (Get-VBRJob | Sort Name)){
    foreach($Object in (Get-VBRJobObject -Job $vbrJob | Sort Name)){
       $Details = "" | Select JobName,VMname,ObjectSizeinGB,Algorithm,RetainCycles,SourceProxyAutoDetect
       $JobSize = ([math]::Round($Object.Info.ApproxSize / 1GB))
    
       $Details.JobName = $vbrJob.Name

       $Details.VMName = $object.Name
    
       $Details.ObjectSizeinGB = $JobSize

       $JobnameObject = Get-VBRJob | Where {$_.Name -eq $vbrJob.Name}
       $Options = $JobnameObject.GetOptions()

       # Determine what kinda job this is
       if ($Options.Options.RootNode.Algorithm -eq "Syntethic"){
          $Algorithm = "Reversed Incremental"
       }
       if ($Options.Options.RootNode.Algorithm -eq "Increment") {
          if ($Options.Options.RootNode.TransformFullToSyntethic -eq "True" -And $Options.Options.RootNode.TransformIncrementsToSyntethic -eq "True") {
             $Algorithm = "Incremental ( Synthetic full enabled, Transform previous full backup chains into rollbacks )"
          }
          if ($Options.Options.RootNode.TransformFullToSyntethic -eq "True" -And $Options.Options.RootNode.TransformIncrementsToSyntethic -eq "False") {
             $Algorithm = "Incremental ( Synthetic full enabled )"
          }
          if ($Options.Options.RootNode.TransformFullToSyntethic -eq "False" -And $Options.Options.RootNode.TransformIncrementsToSyntethic -eq "False") {
             $Algorithm = "Incremental ( Synthetic full disabled, Active full backups )"
          }
       }


       # Detect if SourceProxyAutoDetect is Automatically Selected
       if ($Options.Options.RootNode.SourceProxyAutoDetect -eq "True") {
          $backupproxy = "Automatic selection"
       } else { $backupproxy = $Options.Options.RootNode.SourceProxyAutoDetect }


       #$RetainDays          = $Options.Options.RootNode.RetainDays
       $Details.Algorithm       = $Algorithm
       $Details.RetainCycles    = $Options.Options.RootNode.RetainCycles
       $Details.SourceProxyAutoDetect = $backupproxy
    }

    
    #$details.jobname

#<#
   $sheet.cells.item($x,1) = $Details.JobName
   $sheet.cells.item($x,2) = $Details.VMName
   $sheet.cells.item($x,3) = $Details.ObjectSizeinGB
   $sheet.cells.item($x,4) = $Details.Algorithm
   $sheet.cells.item($x,5) = $Details.RetainCycles
   $sheet.cells.item($x,6) = $Details.SourceProxyAutoDetect
   $x++
 #>

}
Disconnect-VBRServer
Exit
