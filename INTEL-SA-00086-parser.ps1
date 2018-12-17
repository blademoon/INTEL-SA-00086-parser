#Получаем список файлов XML из папки с результатами.
$ResultDirectory = Read-Host -Prompt "Введите путь к папке в которой находятся результаты тестирования "
$ResultFiles = Get-ChildItem $ResultDirectory -Filter *.xml



#Выводим результаты поиска файлов в папке с XML
Write-Host ("__________________________________________________________________________________________________________")
Write-Host ("В указанной папке рекурсивно найдено файлов *.xml: " +$ResultFiles.Count)
Write-Host ("__________________________________________________________________________________________________________")

# Открываем Excel
$excel = New-Object -comobject Excel.Application
# Блокируем запросы на подтверждение выполнения операции, скрываем окно Excel, отключаем обновление окна Excel (повышаем быстродействие).
#$Excel.DisplayAlerts = $false
#$Excel.ScreenUpdating = $false
$Excel.Visible = $True
#$UpdateLinks = $False
#$ReadOnly = $True
$CurrentLocation = [string](Get-Location)
$ExcelFilePath = $CurrentLocation + "\Result_table.xlsx"
$WorkBook = $excel.workbooks.Open($ExcelFilePath)
$WorkSheetName = "INTEL-SA-00086 Detection Tool"
$WorkSheet = $WorkBook.Worksheets.Item($WorkSheetName)
$Cells=$WorkSheet.Cells

#Стартовое значение строки с которой начинаем вставку
$i = 2


foreach($file in $ResultFiles) {
    $FullPath = $file.FullName

    # Получаем информацию из текущего XML файла
    [xml]$XmlDocument = Get-Content $FullPath
    
    $Cells.item($i,1) = $XmlDocument.System.Computer_Name
    $Cells.item($i,2) = $XmlDocument.System.Hardware_Inventory.Computer_Manufacturer
    $Cells.item($i,3) = $XmlDocument.System.Hardware_Inventory.Computer_Model
    $Cells.item($i,4) = $XmlDocument.System.Hardware_Inventory.OperatingSystem
    $Cells.item($i,5) = $XmlDocument.System.Hardware_Inventory.Processor
    $Cells.item($i,6) = $XmlDocument.System.ME_Firmware_Information.Driver_Installed
    $Cells.item($i,7) = $XmlDocument.System.ME_Firmware_Information.FW_Version
    $Cells.item($i,8) = $XmlDocument.System.ME_Firmware_Information.Platform
    $Cells.item($i,9) = $XmlDocument.System.System_Status.System_Risk
    $Cells.item($i,10) = $XmlDocument.System.System_Status.System_Risk_Value
    $Cells.item($i,11) = $XmlDocument.System.Scan_Date
    $Cells.item($i,12) = $XmlDocument.System.Application_Name
    $Cells.item($i,13) = $XmlDocument.System.Application_Version
    $i++
}

# Форматирование листа, сохранение и завершение работы Excel
$usedRange = $WorkSheet.UsedRange                                                                                              
$usedRange.EntireColumn.AutoFit() | Out-Null
$WorkBook.RefreshAll()
Read-Host -Prompt "Press Enter to continue"
