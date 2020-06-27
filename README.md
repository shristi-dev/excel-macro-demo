# A workflow/build system for excel macro development

To get started, install `ImportExcel` PowerShell module. Refer to [ImportExcel Project](https://github.com/dfinke/ImportExcel) for more details.
This code can be used in conjunction with the original CodeProject article published [here](https://www.codeproject.com/script/Articles/ArticleVersion.aspx?waid=4236203&aid=5272220)

**Repo structure**
- `build.ps1` Build script
- `vba\*` VBA files
- `configuraion.psd1` Configuration document containing path to the macro files and their module names

Example:
```
.\build.ps1 -data (get-service|select -first 10) -outputfile "C:\test.xlsx"
```