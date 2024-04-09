' CONECTIONS TO DB


Option Explicit
' EXCEL VARIABLES
Public cnnExcel As New ADODB.Connection
Public rsExcel As New ADODB.Recordset
Public strSQL As String

' MADM VARIABLES
Public MADMConn As New ADODB.Connection
Public MADMrs As New ADODB.Recordset
Public MADMCommand As New ADODB.Command

' CPS VARIABLES
Public CPSConn As New ADODB.Connection
Public CPSrs As New ADODB.Recordset
Public CPSCommand As New ADODB.Command

'AS400 TTS VARIABLES
Public AS400_TTSConn As New ADODB.Connection
Public AS400_TTSRs As New ADODB.Recordset
Public sQuery As String
Public TTSUser As String
Public TTSPass As String






' Set the connection to madm.
Public Function setConnSQLMadm() As Boolean
    On Error GoTo ErrorControl
    Dim connected As Boolean
    connected = False
    
    If MADMConn.State = adStateOpen Then
        MADMConn.Close
    End If

    ' Define the connection string by provider driver and database details.
    Dim sConnString As String
        sConnString = "Provider=SQLOLEDB; " & _
        "Data Source=WSANETDWV; " & _
        "Initial Catalog=Manufacturing;" & _
        "Integrated Security=SSPI"
        MADMConn.ConnectionString = sConnString
        MADMConn.Open         ' Now, open the connection.
        connected = True
done:
     setConnSQLMadm = connected
     Exit Function
ErrorControl:
    'MsgBox "The following error ocurred to try connect MADM: " & Err.Description
    Call SendEmail("Error connect to MADM", "The following error ocurred to try connect MADM: " & Err.Description)
End Function




' Set the connection to TTS.
Public Function setConnAS400_TTS() As Boolean
    On Error GoTo ErrorControl
    Dim connected As Boolean
    connected = False
    
    If AS400_TTSConn.State = adStateOpen Then
        AS400_TTSConn.Close
    End If

    ' Define the connection string by provider driver and database details.
    Dim sConnString As String
    If TTSUser <> "" Or TTSPass <> "" Then
        sConnString = "Provider=IBMDA400;Data Source=HQ400B;User Id=" & TTSUser & ";Password=" & TTSPass & ";"
        AS400_TTSConn.ConnectionString = sConnString
        AS400_TTSConn.Open         ' Now, open the connection.
        connected = True
    Else
        connected = False
    End If
done:
     setConnAS400_TTS = connected
     Exit Function
ErrorControl:
    'MsgBox "The following error ocurred to try connect TTS: " & Err.Description
    Call SendEmail("Error connect to TTS", "The following error ocurred to try connect TTS: " & Err.Description)
End Function





' Set the connection to CPS.
Public Function setConnSQLCPS() As Boolean
    
    On Error GoTo ErrorControl
    Dim connected As Boolean
    connected = False
    
    If CPSConn.State = adStateOpen Then
        CPSConn.Close
    End If

    ' Define the connection string by provider driver and database details.
    Dim sConnString As String
        sConnString = "Provider=SQLOLEDB; " & _
        "Data Source=SQLP1BUS12\P1BUS12; " & _
        "Initial Catalog=CPSMaster;" & _
        "Integrated Security=SSPI"
        CPSConn.ConnectionString = sConnString
        CPSConn.Open         ' Now, open the connection.
        connected = True
done:
     setConnSQLCPS = connected
     Exit Function
ErrorControl:
    'MsgBox "The following error ocurred to try connect CPS: " & Err.Description
    Call SendEmail("Error connect to CPS", "The following error ocurred to try connect CPS: " & Err.Description)
End Function






' Set the connection to excel.
Public Sub OpenExcelDB()
    If cnnExcel.State = adStateOpen Then cnnExcel.Close
    cnnExcel.ConnectionString = "Driver={Microsoft Excel Driver (*.xls, *.xlsx, *.xlsm, *.xlsb)};DBQ=" & _
    ActiveWorkbook.Path & Application.PathSeparator & ActiveWorkbook.Name
    cnnExcel.Open
End Sub




' close the record set
Public Sub closeExcelRS()
    If rsExcel.State = adStateOpen Then rsExcel.Close
    rsExcel.CursorLocation = adUseClient
End Sub








