'Imports Microsoft.Office.Interop.Excel

Module Module1
    Public Pwr_Err_msg() As String = {"電源電壓正常", "電源電壓 過小", "電源電壓 過大", "電源電壓異常"}

    Public Output_Err_msg() As String = {"出力正常", "出力瞬間過電流", "出力斷線", "出力電流檢出異常"}
    Public Loader_Err_msg() As String = {"負荷正常", "電壓限制警報", "馬達電流限制警報", "軟體過電流-警報", "過負荷-異常", "GAS-警報", "入力過電流"}
    Public Comm_Err_msg() As String = {"通信正常", "通信異常"}
    Public Start_Err_msg() As String = {"起動成功", "起動失敗", "起動 -- ", "起動 系統異常"}
    Public CtlV_Err_msg() As String = {"制御電源-正常", "制御電源-", "制御電源-異常", "制御電源-"}
    Public STB_Err_msg() As String = {"STB-正常", "STB-警報"}
    Public STB_Sign_msg() As String = {"STB-OFF", "STB-ON"}
    Public OVH_Err_msg() As String = {"正常", "過熱L", "過熱S", "Sensor異常"}
    Public LOH_Sign_msg() As String = {"低溫-正常", "低溫-警報"}




    'Sub Creat_excel()
    '    Dim fileName As String = "D:\test\testExcel.xlsx"
    '    Dim xlApp As New Application()
    '    If xlApp IsNot Nothing Then
    '        xlApp.Visible = True
    '        Dim wb As Workbook = xlApp.Workbooks.Open(fileName)
    '        CType(wb.Sheets(1), Worksheet).Select()
    '        Dim aRange As Range = xlApp.Range("A1")
    '        If aRange IsNot Nothing Then
    '            Console.WriteLine(If(aRange.Value2, String.Empty))
    '            aRange.Value2 = "コード レシピ Excel 編"
    '            Console.WriteLine(aRange.Value2)
    '        End If
    '        xlApp.ActiveWorkbook.Close(True)
    '        xlApp.Quit()
    '    End If
    'End Sub
End Module
