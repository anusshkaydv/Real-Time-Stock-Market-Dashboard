Option Explicit

Sub RefreshDashboard()

    Dim PythonExe As String
    Dim PythonScript As String
    Dim WshShell As Object
    Dim Exec As Object

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    ' Python executable
    PythonExe = "C:\Users\anush\AppData\Local\Python\pythoncore-3.14-64\python.exe"

    ' Python script
    PythonScript = "C:\Users\anush\Desktop\Real-Time Stock Market Dashboard\Stock_Data_Fetcher.py"

    Set WshShell = CreateObject("WScript.Shell")

    ' Run Python and WAIT until it finishes
    Set Exec = WshShell.Exec("""" & PythonExe & """ """ & PythonScript & """")

    Do While Exec.Status = 0
        DoEvents
    Loop

    ' Refresh Power Query
    ThisWorkbook.RefreshAll

    ' Wait for refresh to complete
    Application.CalculateUntilAsyncQueriesDone

    ' Refresh all Pivot Tables
    Dim ws As Worksheet
    Dim pt As PivotTable

    For Each ws In ThisWorkbook.Worksheets
        For Each pt In ws.PivotTables
            pt.RefreshTable
        Next pt
    Next ws

    Application.CalculateFull

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

    MsgBox "Dashboard Updated Successfully!", vbInformation

End Sub
