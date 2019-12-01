#some documentation
#https://docs.microsoft.com/en-us/dotnet/api/microsoft.office.interop.excel.xlfileformat?view=excel-pia
#https://stackoverflow.com/questions/39376896/vba-vbproject-vbcomponents-itemthisworkbook-codemodule-addfromstring-isnt
#https://docs.microsoft.com/en-us/office/vba/api/excel.application.automationsecurity
#https://docs.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa443946(v=vs.60)

function Open-File {
    Add-Type -AssemblyName System.Windows.Forms
    #use your DesktopFolder or w/e
    $initDir = [System.Environment]::GetFolderPath('Desktop')
    $FileBrowser = New-Object System.Windows.Forms.OpenFileDialog -Property @{InitialDirectory = $initDir}
    #if | Out-Null is not used here, functions returns array with OK and FileName
    $FileBrowser.ShowDialog() | Out-Null
    $filePath = $FileBrowser.FileName
    return $filePath
}
function Release-Ref($ref) {
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject([System.__ComObject]$ref) | Out-Null
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}

$excel = New-Object -ComObject Excel.Application
$excel.Application.EnableEvents = $false
$excel.DisplayAlerts = $false
$excel.Application.AutomationSecurity = 3

$FilePath = Open-File
if(-not (Test-Path -Path $FilePath)) {
    return
}

$code = @"
Sub test()
'
End Sub
"@


New-ItemProperty -Path "HKCU:\Software\Microsoft\Office\$($excel.Version)\excel\Security" -Name AccessVBOM -Value 1 -Force | Out-Null
New-ItemProperty -Path "HKCU:\Software\Microsoft\Office\$($excel.Version)\excel\Security" -Name VBAWarnings -Value 1 -Force | Out-Null


$workbook = $excel.Workbooks.Add($FilePath)
$ListOfModules = @($workbook.VBProject.VBComponents)
foreach ($module in $ListOfModules) {
    if($module.Name -eq 'Test') {
        $TestModuleFound = $true
    } else {
        $TestModuleFound = $false
    }
}
#if module with name 'test' exists we delete 1st 4lines and add $code again which will cause recompile and excel file will work
if($TestModuleFound) {
    $xlmodule = $workbook.VBProject.VBComponents.item('Test')
    $xlmodule.CodeModule.DeleteLines(1,3)
    $xlmodule.CodeModule.AddFromString($code)
} else {
    #if module test does not exists we crete one and add some lines of dummy code to it which causes project to recompile and file will work
    # .Add(1) for module
    # .Add(2) for class I believe
    # .Add(3) for for UserForm
    $workbook.VBProject.VBComponents.Add(1).Name = "Test"
    $xlmodule = $workbook.VBProject.VBcomponents.item('Test')
    $xlmodule.CodeModule.AddFromString($code);
}
#save as 52 so it is saved as macro enabled file
#TODO: Excel application is not closing automatically, that should be fixed so we don't have to call stop-process

$workbook.SaveAs($FilePath,52)
$excel.Application.EnableEvents = $true
$excel.DisplayAlerts = $true
$excel.Application.AutomationSecurity = 1
$excel.Workbooks.Close()
$excel.Application.Quit()
$excel.Quit()
#release
Release-Ref($excel)
#quit excel as $excel.Quit() does not do the job or I'm stupid
Get-Process excel | Stop-Process
Write-Output "Zavrseno."