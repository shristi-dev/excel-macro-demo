param($Data,$OutputFile)

Import-Module ImportExcel

$configuration = Import-LocalizedData -BaseDirectory $PSScriptRoot -FileName configuration.psd1
$excelPackage = $Data | Export-Excel -PassThru
$excelPackage.Workbook.CreateVBAProject()
$configuration.Macro.Modules.Keys |ForEach-Object {
    $module = $excelPackage.Workbook.VbaProject.Modules.AddModule($_)
    $module.Code = (Get-Content (Join-Path $configuration.Macro.BasePath $configuration.Macro.Modules[$_])) -join "`n"
}

$excelPackage.SaveAs($OutputFile)