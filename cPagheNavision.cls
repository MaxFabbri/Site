VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cPutPay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit

Implements IPutPay

' dati standard

Private mFromDate               As Date
Private mToDate                 As Date

Private DebugOperation          As String

Private mNominativo             As String
Private mID                     As Integer

Private mbHasGone               As Boolean

Private Type LayoutDay
    ORE         As String
    Causale     As String
End Type

Private Type Layout
    Cartellino      As String
    Data            As String
    Day(1 To 5)     As LayoutDay
    crlf            As String
End Type

Dim navisionLayout  As Layout

Dim personCode      As String
Dim payCodes()      As String

Dim formatHour      As String
Dim getHourValue    As Boolean

Dim formatDate      As String

Dim FileName        As String

Dim hFile           As Integer

Dim outList()       As String

Private Function IPutPay_AfterInitializeDip(CustomMessage As String) As Boolean
    IPutPay_AfterInitializeDip = True
End Function

Private Sub IPutPay_BuildOutPutFile(WebCaller As Boolean)

    On Error GoTo Pay_ERROR
    
    DebugOperation = "Apertura file " & FileName
    DebugOut DebugOperation

    hFile = FreeFile
    Open FileName For Output As hFile

Pay_EXIT:

    Exit Sub
    
Pay_ERROR:

    Select Case CatchErr(Err.Description, Err.Number, App.EXEName & ".IPutPay_BuildOutPutFile")
    Case vbAbort
        Resume Pay_EXIT
    Case vbRetry
        Resume
    Case vbIgnore
        Resume Next
    End Select

End Sub

Private Function IPutPay_ChangeDay(oday As CDatoGiornaliero) As ePrctInfoStatus
            
    Dim hour                As cOra

    Dim quantity            As Double

    Dim Index               As Integer

    Dim dailyCode           As Boolean

    On Error GoTo Pay_ERROR

    DebugOut "Function IPutPay_ChangeDay"
    
    Debug.Print oday.Giorno
    
    ResetNavisionLayout navisionLayout
    
    navisionLayout.Data = Format$(oday.Giorno, formatDate)
    
    ' calcola le ore lavorate le ore festive e le ore di straordinario non autorizzate
    For Each hour In oday.ORE.Group.Sort
        With hour
            If IsInArray(payCodes, .GetPayCode()) Then
                If Not (oday.Profili(1) Is Nothing) Then
                    Index = Index + 1
                    If (Index <= 5) Then
                        quantity = .ResolvePayQuantity(TimeToMinute(oday.Profili(1).OreMassime), dailyCode)
                        If (Not .Centesimi And Not dailyCode) Then
                            ' è in sessantesimi lo inserisce in base al formato
                            navisionLayout.Day(Index).ORE = MinuteToTime(quantity, , formatHour, getHourValue)
                        Else
                            navisionLayout.Day(Index).ORE = quantity
                        End If
                        navisionLayout.Day(Index).Causale = hour.GetPayCode()
                        IPutPay_ChangeDay = Successful
                    End If
                End If
            End If
        End With
    Next
    
    If (IPutPay_ChangeDay = Successful) Then
        Print #hFile, GetNavisionLayout(navisionLayout)
    End If
    
Pay_EXIT:

    Exit Function
    
Pay_ERROR:

    Select Case CatchErr(Err.Description, Err.Number, App.EXEName & ".IPutPay_ChangeDay" & vbCrLf & "(" & DebugOperation & ")" & vbCrLf & mNominativo)
    Case vbAbort
        Resume Pay_EXIT
    Case vbRetry
        Resume
    Case vbIgnore
        Resume Next
    End Select

End Function

Private Function IPutPay_CheckPeriod() As Boolean

    DebugOut "Function IPutPay_CheckPeriod"
    
    IPutPay_CheckPeriod = oDip.isActive
        
End Function

Private Function IPutPay_DipSelection(ByVal dFromDate As Date, ByVal dToDate As Date, ByVal Name As String) As ADODB.Recordset

On Error GoTo Pay_ERROR

    DebugOut "Function IPutPay_DipSelection"
    
    DebugOperation = "Selezionamento dipendenti"
    Set IPutPay_DipSelection = SelectDip(True, dFromDate, dToDate, "Select " & Name, _
                            "Dipendenti.Cognome,Dipendenti.Nome")


Pay_EXIT:
    Exit Function

Pay_ERROR:

    Select Case CatchErr(Err.Description, Err.Number, App.EXEName & ".IPutPay_DipSelection" & vbCrLf & "(" & DebugOperation & ")" & vbCrLf)
    Case vbAbort
        Resume Pay_EXIT
    Case vbRetry
        Resume
    Case vbIgnore
        Resume Next
    End Select

End Function

Private Property Let IPutPay_FromDate(ByVal RHS As Date)
    mFromDate = RHS
End Property

Private Property Get IPutPay_FromDate() As Date
    IPutPay_FromDate = mFromDate
End Property

Private Property Get IPutPay_FullCalendar() As Boolean
    IPutPay_FullCalendar = True
End Property

Private Function IPutPay_GetContext() As String
    
    ' lasciare vuoto
    
End Function

Private Property Get IPutPay_GetFileList() As String()

    ' ritornare l'elenco dei files da scrivere
    IPutPay_GetFileList = outList

End Property

Private Function IPutPay_GetPassword() As String
    If (InDesign) Then
        IPutPay_GetPassword = "ARTECHADMIN"
    End If
End Function

Private Function IPutPay_GetUser() As String
    If (InDesign) Then
        IPutPay_GetUser = "ARTECH"
    End If
End Function

Private Function IPutPay_Initialize(WebCaller As Boolean) As Boolean

    On Error GoTo Pay_ERROR
    
    DebugOut "Sub IPutPay_Initialize"
    
    With oDip
    
        .ReadsDailyCalendar = True
        .ReadsDailyProfiles = True
        .ReadsHours = True
        .FromDate = mFromDate
        .ToDate = mToDate

    End With
    
    ' l'operazione di cancellazione del file viene effettauta dal main
    Dim location As eINIWhere

    ' la location viene definita sul server o in locale
    location = eINIWhere.LocalINI
    If WebCaller Then
        location = GlobalINI
    End If
    
    ' mettere un percorso univoco per ogni CLIENTE
    FileName = RINI("Files", "Path", location)
    
    'RenFileProgr FileName, 99, frmHide.ListApp, ""
    
    ' fare il push dei vari file aperti
    Push outList(), FileName
    
    navisionLayout.crlf = vbCrLf
    
    mbHasGone = True
    
    IPutPay_Initialize = True
    
Pay_EXIT:
    Exit Function

Pay_ERROR:

    Select Case CatchErr(Err.Description, Err.Number, App.EXEName & ".IPutPay_Initialize" & vbCrLf & "(" & DebugOperation & ")" & vbCrLf)
    Case vbAbort
        Resume Pay_EXIT
    Case vbRetry
        Resume
    Case vbIgnore
        Resume Next
    End Select

End Function

Private Function IPutPay_InitializeDip(CustomMessage As String) As Boolean

    On Error GoTo Pay_ERROR

    DebugOut "Sub IPutPay_InitializeDip"
    
    navisionLayout.Cartellino = oDip.GetAnagData(personCode, mFromDate)
    
    mNominativo = oDip.Nominativo
    mID = oDip.DipID
    
    IPutPay_InitializeDip = True
    
Pay_EXIT:
    Exit Function
    
Pay_ERROR:

    Select Case CatchErr(Err.Description, Err.Number, App.EXEName & ".IPutPay_InitializeDip" & vbCrLf & "(" & DebugOperation & ")" & vbCrLf & mNominativo)
    Case vbAbort
        Resume Pay_EXIT
    Case vbRetry
        Resume
    Case vbIgnore
        Resume Next
    End Select
    
End Function


Private Sub IPutPay_InitializePerc(oPerc As IPercentageInfo)

    DebugOut "Sub IPutPay_InitializePerc"
    
    With oPerc
        .DefaultCalculatedValue = Information
        ' non c'è...
        If (RINI("Opzioni", "Numero secondi di attesa al termine") = "") Then
            .WaitSec = 60
        Else
            .WaitSec = Val(RINI("Opzioni", "Numero secondi di attesa al termine"))
        End If
    End With

End Sub


Private Sub IPutPay_MakeSectionsKeys()

    Dim oS As CSection

    oSects.IniServerName = App.EXEName & ".ini"
    oSects.DescrizioneGenerale = App.FileDescription & " " & App.Major & "." & App.Minor & "." & App.Revision
    
    Set oS = oSects.Add("Parametri")
    oS.AddKey "Codice Persona In Paga", StringText, "Inserire il codice che identifica la persona alle paghe"
    oS.AddKey "Elenco Causali In Paga", ListPayCode, "Codici da passare in paga"
    oS.AddKey "Formato Orario In Sessagesimi", StringText, "Inserire il formato orario solo se in sessagesimi es: hh:nn"
    oS.AddKey "Preleva Valore In Sessagesimi", ComboBox, "Preleva solamente il valore eliminando gli zeri non significativi", "0,1", "No,Sì"
    oS.AddKey "Formato Data", StringText, "Inserire obbligatoriamente il formato della data es: dd/MM/yyyy"

End Sub

Private Property Get IPutPay_MemoLogin() As Boolean

End Property

Private Function IPutPay_PreInitialize() As Boolean
    IPutPay_PreInitialize = True
End Function

Private Function IPutPay_PreInitializeDip(CustomMessage As String) As Boolean
    IPutPay_PreInitializeDip = True
End Function

Private Sub IPutPay_ReadINIParameters()

    personCode = RINI("Parametri", "Codice Persona In Paga")
    payCodes = Split(RINI("Parametri", "Elenco Causali In Paga"), ",")
    
    formatHour = RINI("Parametri", "Formato Orario In Sessagesimi")
    getHourValue = (Val(RINI("Parametri", "Preleva Valore In Sessagesimi")) <> 0)
    
    formatDate = RINI("Parametri", "Formato Data")
    
    
    If (formatHour = "") Then
        formatHour = "hh:nn"
    End If
    
    If (formatDate = "") Then
        formatDate = "dd/MM/yyyy"
    End If
    
    If (personCode = "") Then
        personCode = "codpaghe"
    End If

End Sub

Private Property Get IPutPay_SilentMode() As Boolean

End Property

Private Sub IPutPay_Terminate(WebCaller As Boolean)

    On Error GoTo Pay_ERROR

    DebugOut "Sub IPutPay_Terminate"
    
    Dim ofile   As New cFilePath
    
    If mbHasGone Then
        
        If (hFile > 0) Then
            Close hFile
            'DoEvents
            ofile.AddFile FileName, True
        End If
        
    End If
    
    If ofile.IsReady Then
        ofile.Show
    End If
    
    Set ofile = Nothing
    
Pay_EXIT:
    Exit Sub

Pay_ERROR:

    Select Case CatchErr(Err.Description, Err.Number, App.EXEName & ".IPutPay_Terminate" & vbCrLf & "(" & DebugOperation & ")" & vbCrLf & mNominativo)
    Case vbAbort
        Resume Pay_EXIT
    Case vbRetry
        Resume
    Case vbIgnore
        Resume Next
    End Select

End Sub


Private Function IPutPay_TerminateDip(eState As ePrctInfoStatus, CustomMessage As String) As ePrctInfoStatus

    On Error GoTo Pay_ERROR
    
    DebugOut "Function IPutPay_TerminateDip()"
    
    IPutPay_TerminateDip = eState
    
Pay_EXIT:
    Exit Function

Pay_ERROR:

    Select Case CatchErr(Err.Description, Err.Number, App.EXEName & ".IPutPay_TerminateDip" & vbCrLf & "(" & DebugOperation & ")" & vbCrLf & mNominativo)
    Case vbAbort
        Resume Pay_EXIT
    Case vbRetry
        Resume
    Case vbIgnore
        Resume Next
    End Select
    
End Function

Private Property Let IPutPay_ToDate(ByVal RHS As Date)
    mToDate = RHS
End Property

Private Property Get IPutPay_ToDate() As Date
    IPutPay_ToDate = mToDate
End Property

Private Function IPutPay_ViewCalendar(dFromDate As Date, dToDate As Date, ByVal Name As String) As Boolean

End Function

Private Function IPutPay_ViewFullCalendar(StartDate As Date, EndDate As Date, ByVal Name As String) As ADODB.Recordset

    On Error GoTo Pay_ERROR

    DebugOut "Function IPutPay_ViewFullCalendar()"

    DebugOperation = "View calendar"

    Dim oC  As New cCalendar
    With oC
        .SelectorCount = 6
        .OnlyPath = False
        .CalendarMode = PreviousMonth
        .Caption = "Periodo " & Name
        Set .Icon = LoadResPicture(101, 1)
        .NumberFileOutPut = 1
        .FileName(1) = "Nome file paghe"
        .Fullmonth = True
    End With
    
    Set IPutPay_ViewFullCalendar = oC.Run()
    
    StartDate = oC.StartDate
    EndDate = oC.EndDate
    
    mFromDate = oC.StartDate
    mToDate = oC.EndDate
    
    Set oC = Nothing

Pay_EXIT:
    Exit Function

Pay_ERROR:

    Select Case CatchErr(Err.Description, Err.Number, App.EXEName & ".IPutPay_ViewFullCalendar" & vbCrLf & "(" & DebugOperation & ")" & vbCrLf & mNominativo)
    Case vbAbort
        Resume Pay_EXIT
    Case vbRetry
        Resume
    Case vbIgnore
        Resume Next
    End Select


    
End Function

Private Property Get IPutPay_ViewLOG(LogList As String, Optional WaitSec As Integer = 15) As Boolean

End Property

Private Property Get IPutPay_ViewPercentage() As Boolean
    IPutPay_ViewPercentage = True
End Property

Private Function GetFormatValue(D As Double) As String

    ' se c'è il punto o la virgola li toglie
    If InStr(D, ".") > 0 Then ' toglie il punto
        GetFormatValue = Replace(Format(D, "0000000.00000"), ".", "")
    ElseIf InStr(D, ",") > 0 Then ' toglie la virgola
        GetFormatValue = Replace(Format(D, "0000000.00000"), ",", "")
    Else
        GetFormatValue = Format(D, "0000000") & "00000"
    End If

End Function

Private Function GetNavisionLayout(navisionLayout As Layout) As String

    Dim Index   As Integer
    Dim ret     As String
    
    If Len(navisionLayout.Cartellino) > 0 Then
        If (Right$(navisionLayout.Cartellino, 1) <> ";") Then
            navisionLayout.Cartellino = navisionLayout.Cartellino & ";"
        End If
    End If
    
    navisionLayout.Data = navisionLayout.Data & ";"
    
    GetNavisionLayout = navisionLayout.Cartellino & navisionLayout.Data
    
    For Index = 1 To 5
        GetNavisionLayout = GetNavisionLayout & navisionLayout.Day(Index).ORE & ";"
        GetNavisionLayout = GetNavisionLayout & navisionLayout.Day(Index).Causale & ";"
    Next
    
End Function

Private Sub ResetNavisionLayout(navisionLayout As Layout)

    Dim Index   As Integer
    
    navisionLayout.Data = ""
    For Index = 1 To 5
        navisionLayout.Day(Index).Causale = ""
        navisionLayout.Day(Index).ORE = ""
    Next

End Sub
