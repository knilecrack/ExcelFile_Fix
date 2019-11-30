#some documentation
#https://docs.microsoft.com/en-us/dotnet/api/microsoft.office.interop.excel.xlfileformat?view=excel-pia
#https://stackoverflow.com/questions/39376896/vba-vbproject-vbcomponents-itemthisworkbook-codemodule-addfromstring-isnt
#https://docs.microsoft.com/en-us/office/vba/api/excel.application.automationsecurity
#https://docs.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa443946(v=vs.60)

function Open-File {
    param (
        $filePath
    )
    Add-Type -AssemblyName System.Windows.Forms
    #use your DesktopFolder or w/e
    $initDir = [System.Environment]::GetFolderPath('Desktop')
    $FileBrowser = New-Object System.Windows.Forms.OpenFileDialog -Property @{InitialDirectory = $initDir}
    $FileBrowser.ShowDialog()
    return $FileBrowser.FileName
}


$excel = New-Object -ComObject Excel.Application
$excel.Application.EnableEvents = $false
$excel.Application.AutomationSecurity = 3
#$FilePath = "\\172.24.48.2\Šabloni\Zajednički folder\Zajednički folder\Knjigovodstvo TEST\Knile_Temp_Folder\Knjigovodstvo Fullhand 011.xlsm"
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

    #if module test does not exists we crete one and add some lines of dummy code to it which cases project to recompile and file will work
    # Add(1) for module
    # Add(2) for class I believe
    # add(3) for for UserForm
    $workbook.VBProject.VBComponents.Add(1).Name = "Test"
    $xlmodule = $workbook.VBProject.VBcomponents.item('Test')
    $xlmodule.CodeModule.AddFromString($code);
}
#save as 52 so it can be saved as macro enabled file
$workbook.SaveAs($FilePath,52)
$excel.Application.EnableEvents = $true
$excel.Application.AutomationSecurity = 1
$excel.Workbooks.Close()
$excel.Application.Quit()
get-process excel | stop-process