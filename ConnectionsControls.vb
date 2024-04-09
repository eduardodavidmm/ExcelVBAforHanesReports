' CONECTION CONTROLS FUNCTIONS

Option Explicit


' =====VARIABLES=====
Public rng As Range
Public cell As Range
Public BucketRow As Integer
Public initialDate As String
Public endDate As String
Public InitialDateTTS As String
Public EndDateTTS As String
Public rowCount As Integer
Public f As Long
Public initialTransaction As String
Public endTransaction As String
Public bucketId As String
Public bucketSource As String
Public startProcess As Date
Public endProcess As Date
Public totalProcess As Integer
Public transmittedOrders As String
Public totalTransmittedOrder As Integer
Public InitialTransCondition As String
Public EndTransCondition As String
Public plants As String
Public CutDueDateAnalysis As Boolean
Public SewDueDateAnalysis As Boolean
Public DCDueDateAnalysis As Boolean
Public BucketDescription As String
Public BucketGoalInDays As Integer
Public plantCategory As String
Public plantName As String
Public SupplyChainHub As String
Public WeekSelected As String
Public valueA As String
Public valueB As String
Public reportType As String






'CALL THE FORM TO ENTER THE TTS CREDENTIALS.
Public Sub callForm()
    Load TTS_Credentials_Form
    TTS_Credentials_Form.Show
End Sub


'READ THE EXCEL CONFIG FILE
Public Sub readValues()
    ' Open Workbook A with specific location
    Dim src As Workbook
    Set src = Workbooks.Open("\\v3v4hfps01\SSIS\GeneralConfigFile\LeadTimeCalculationReport_ConfigFile.xlsx", True, True)

    valueA = src.Worksheets("sheet1").Cells(1, 1)
    valueB = src.Worksheets("sheet1").Cells(2, 1)

    ' Close Workbooks A
    src.Close False
    Set src = Nothing
End Sub





'FUNCTION TO VERIFY IF THE USER CAN ACCESS TO MADM DATABASE SOURCE
Public Sub VerifyConnectionToMADM()
    If setConnSQLMadm = True Then
        MsgBox "You have access to MADM, you can run this report."
    Else
        MsgBox "You don't have access to MADM, you can't run this report."
    End If
End Sub




'FUNCTION TO VERIFY IF THE USER CAN ACCESS TO MADM DATABASE SOURCE
Public Sub VerifyConnectionToTTS()
    callForm
     If setConnAS400_TTS = True Then
        MsgBox "You have access to TTS, you can run this report."
    Else
        MsgBox "You don't have access to TTS, you can't run this report."
    End If
End Sub



'FUNCTION TO VERIFY IF THE USER CAN ACCESS TO MADM DATABASE SOURCE
Public Sub VerifyConnectionToCPS()
  If setConnSQLCPS = True Then
        MsgBox "You have access to CPS, you can run this report."
    Else
        MsgBox "You don't have access to CPS, you can't run this report."
    End If
End Sub



'FUNCTION TO VERIFY IF THE USER CAN ACCESS TO DIFFERENT DATABASE SOURCE
Public Function verifyConections() As Boolean
    Dim connectionsCorrect As Boolean
    connectionsCorrect = True
    
    ' CONNECTION VERIFICATION TO MADM
     If setConnSQLMadm = False Then
        connectionsCorrect = False
        verifyConections = connectionsCorrect
        Exit Function
     End If
          
    ' CONNECTION VERIFICATION TO CPS
     If setConnSQLCPS = False Then
        connectionsCorrect = False
        verifyConections = connectionsCorrect
        Exit Function
     End If
         
     'CALL THE FUNCTION TO READ THE FILE
     readValues
     TTSUser = Encryption("Hanes", valueA, False)
     TTSPass = Encryption("Hanes", valueB, False)

      ' CONNECTION VERIFICATION TO TTS
     If setConnAS400_TTS = False Then
        connectionsCorrect = False
        verifyConections = connectionsCorrect
        Exit Function
     End If
     verifyConections = connectionsCorrect
End Function





Private Function StrToPsd(ByVal Txt As String) As Long
'UpdatebyKutoolsforExcel20151225
    Dim xVal As Long
    Dim xCh As Long
    Dim xSft1 As Long
    Dim xSft2 As Long
    Dim I As Integer
    Dim xLen As Integer
    xLen = Len(Txt)
    For I = 1 To xLen
        xCh = Asc(Mid$(Txt, I, 1))
        xVal = xVal Xor (xCh * 2 ^ xSft1)
        xVal = xVal Xor (xCh * 2 ^ xSft2)
        xSft1 = (xSft1 + 7) Mod 19
        xSft2 = (xSft2 + 13) Mod 23
    Next I
    StrToPsd = xVal
End Function




Private Function Encryption(ByVal Psd As String, ByVal InTxt As String, Optional ByVal Enc As Boolean = True) As String
    Dim xOffset As Long
    Dim xLen As Integer
    Dim I As Integer
    Dim xCh As Integer
    Dim xOutTxt As String
    xOffset = StrToPsd(Psd)
    Rnd -1
    Randomize xOffset
    xLen = Len(InTxt)
    For I = 1 To xLen
        xCh = Asc(Mid$(InTxt, I, 1))
        If xCh >= 32 And xCh <= 126 Then
            xCh = xCh - 32
            xOffset = Int((96) * Rnd)
            If Enc Then
                xCh = ((xCh + xOffset) Mod 95)
            Else
                xCh = ((xCh - xOffset) Mod 95)
                If xCh < 0 Then xCh = xCh + 95
            End If
            xCh = xCh + 32
            xOutTxt = xOutTxt & Chr$(xCh)
        End If
    Next I
    Encryption = xOutTxt
End Function


