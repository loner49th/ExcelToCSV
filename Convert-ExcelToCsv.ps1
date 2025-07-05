param(
    [Parameter(Mandatory=$true)]
    [string]$ExcelFilePath,
    
    [Parameter(Mandatory=$false)]
    [string]$OutputPath = "",
    
    [Parameter(Mandatory=$false)]
    [string]$WorksheetName = "",
    
    [Parameter(Mandatory=$false)]
    [switch]$AllWorksheets
)

function Convert-ExcelToCsv {
    param(
        [string]$ExcelPath,
        [string]$CsvPath,
        [string]$Worksheet = "",
        [bool]$ConvertAllSheets = $false
    )
    
    # Convert relative path to absolute path
    try {
        $ExcelPath = Resolve-Path $ExcelPath -ErrorAction Stop
    } catch {
        Write-Error "Excel file not found: $ExcelPath"
        return
    }
    
    try {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        
        $workbook = $excel.Workbooks.Open($ExcelPath)
        
        if ($ConvertAllSheets) {
            foreach ($sheet in $workbook.Worksheets) {
                $sheetName = $sheet.Name
                $csvFileName = [System.IO.Path]::GetFileNameWithoutExtension($ExcelPath) + "_" + $sheetName + ".csv"
                
                if ($CsvPath -eq "") {
                    $csvFullPath = [System.IO.Path]::Combine([System.IO.Path]::GetDirectoryName($ExcelPath), $csvFileName)
                } else {
                    $csvFullPath = [System.IO.Path]::Combine($CsvPath, $csvFileName)
                }
                
                $sheet.SaveAs($csvFullPath, 6)
                Write-Host "Conversion completed: $csvFullPath"
            }
        } else {
            $targetSheet = $null
            
            if ($Worksheet -ne "") {
                $targetSheet = $workbook.Worksheets | Where-Object { $_.Name -eq $Worksheet }
                if ($targetSheet -eq $null) {
                    Write-Error "Specified worksheet '$Worksheet' not found"
                    return
                }
            } else {
                $targetSheet = $workbook.Worksheets[1]
            }
            
            $csvFileName = [System.IO.Path]::GetFileNameWithoutExtension($ExcelPath) + ".csv"
            
            if ($CsvPath -eq "") {
                $csvFullPath = [System.IO.Path]::Combine([System.IO.Path]::GetDirectoryName($ExcelPath), $csvFileName)
            } else {
                if (Test-Path $CsvPath -PathType Container) {
                    $csvFullPath = [System.IO.Path]::Combine($CsvPath, $csvFileName)
                } else {
                    $csvFullPath = $CsvPath
                }
            }
            
            $targetSheet.SaveAs($csvFullPath, 6)
            Write-Host "Conversion completed: $csvFullPath"
        }
        
        $workbook.Close($false)
        $excel.Quit()
        
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
        
    } catch {
        Write-Error "Conversion error: $($_.Exception.Message)"
    }
}

# Main process
if (-not $ExcelFilePath) {
    Write-Host "Usage:"
    Write-Host "  .\Convert-ExcelToCsv.ps1 -ExcelFilePath 'file.xlsx'"
    Write-Host "  .\Convert-ExcelToCsv.ps1 -ExcelFilePath '.\data\file.xlsx' -OutputPath '.\output\'"
    Write-Host "  .\Convert-ExcelToCsv.ps1 -ExcelFilePath 'file.xlsx' -WorksheetName 'Sheet1'"
    Write-Host "  .\Convert-ExcelToCsv.ps1 -ExcelFilePath 'file.xlsx' -AllWorksheets"
    exit 1
}

Convert-ExcelToCsv -ExcelPath $ExcelFilePath -CsvPath $OutputPath -Worksheet $WorksheetName -ConvertAllSheets $AllWorksheets.IsPresent