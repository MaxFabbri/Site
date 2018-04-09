VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"

Option Explicit

Implements IMain

Private DebugOperation      As String
Private DebugDipValue       As String

Private mCurrDate           As Date

Private LogTrack            As Boolean

Private profileList()       As String
Private delayCode           As String

Private Function IMain_Calculate() As ElabSrvr.CalcResultEnum

' post-elaborazione
'
' ricava i dati supplementari del profilo orario con determinata causale
' assieme alla lista dei turni coinvolti

Dim daily                           As CDailyData

Dim stackMessage                    As IStackMessage

Dim somethingHappened               As Boolean

Dim newHour                         As cHour
Dim mark                            As cMark

Dim slotMarks                       As cAdvTimeSlotCodes
Dim workedSlot                      As cAdvTimeSlotCodes
Dim delaySlot                       As cAdvTimeSlotCodes

Dim slot                            As cAdvTimeSlotCode
Dim slotCover                       As cAdvTimeSlotCode

Dim calcSlot                        As New cCalcSlot

Dim delayMinutes                    As Integer

    On Error GoTo Main_ERROR
    
    IMain_Calculate = wpECalcOK
    
    DebugOperation = "Inizializzazione oggetto organize"
    With oDip
    
        .Change DipID
        
        ' dati relativi ai profili
        .ReadsDailyProfiles = True
        .ReadsDetailsProfile = True
        
        ' legge le opzioni avanzate del profilo orario
        .ReadsAdvancedBandsCodes = True
        
        ' dati relativi alle ore
        .ReadsHours = True
        .ReadsMarks = True
        
        .OrganizeData DipID
        
    End With
    
    Set stackMessage = New cStackMessages
    
    For mCurrDate = FromDate To ToDate
    
        DebugOperation = "ci sono dati il " & mCurrDate & " ?"
        Set daily = oDip.DailyDatas(mCurrDate)
        If (daily Is Nothing) Then
            GoTo SkipDay
        End If
         
        DebugDipValue = "data " & mCurrDate & " dipendente " & oDip.Nominativo
        
        With daily
        
            stackMessage.Push String(80, "_")
            stackMessage.Push DebugDipValue
            
            If (.profiles(1) Is Nothing) Then
               GoTo SkipDay
            End If
            
            If (Not IsInArray(profileList, .profiles(1).Code)) Then
                GoTo SkipDay
            End If
            
            If (.profiles(1).GetAdvTimeSlotCodes Is Nothing) Then
                GoTo SkipDay
            End If

            ' marcature a fasce orarie
            Set slotMarks = calcSlot.MakeSlotMarks(.marks, False)
            Set workedSlot = New cAdvTimeSlotCodes
            
            For Each slot In .profiles(1).GetAdvTimeSlotCodes
                ' la causale inserita nelle fasce supplementari del profilo orario deve corrispondere al ritardo
                If (slot.Causale = delayCode) Then
                    For Each slotCover In calcSlot.GetCoveredSlot(.Day + slot.DaOra, .Day + slot.AOra, slotMarks)
                        ' deve esserci una marcatura all'interno della fascia oraria configurata come ritardo
                        For Each mark In .marks
                            ' ok la fascia è coperta ma la marcatura deve essere all'interno della fascia
                            If .IsMarkInRange(mark, slotCover.DaOra, slotCover.AOra) Then
                                ' abbiamo il lavorato all'interno della fascia
                                calcSlot.CheckSlot workedSlot, slotCover.DaOra, slotCover.AOra, slot.Causale
                            End If
                        Next
                    Next
                End If
            Next
            
            If (workedSlot.Count = 0) Then
                GoTo SkipDay
            End If
            
            Set delaySlot = New cAdvTimeSlotCodes
            ' il ritardo è esattamente l'opposto del lavorato
            Set workedSlot = calcSlot.ReverseSlot(workedSlot)
            
            For Each slot In .profiles(1).GetAdvTimeSlotCodes
                ' la causale inserita nelle fasce supplementari del profilo orario deve corrispondere al ritardo
                If (slot.Causale = delayCode) Then
                    For Each slotCover In calcSlot.GetCoveredSlot(.Day + slot.DaOra, .Day + slot.AOra, workedSlot)
                        ' deve esserci una marcatura all'interno della fascia oraria configurata come ritardo
                        For Each mark In .marks
                            ' ok la fascia è coperta ma la marcatura deve essere all'interno della fascia
                            If .IsMarkInRange(mark, slotCover.DaOra, slotCover.AOra) Then
                                ' abbiamo il lavorato all'interno della fascia
                                calcSlot.CheckSlot delaySlot, slotCover.DaOra, slotCover.AOra, slot.Causale
                            End If
                        Next
                    Next
                End If
            Next
            
            If (delaySlot.Count = 0) Then
                GoTo SkipDay
            End If

            ' in delaySlot abbiamo tutte le fasce orarie
            ' da arrotondare ai 15' in eccesso
            For Each slot In delaySlot
            
                Set newHour = New cHour
                
                delayMinutes = ((TimeToMinute(TimeValue(slot.AOra) - TimeValue(slot.DaOra)) \ 15) + 1) * 15
                
                With newHour
                    .Code = slot.Causale
                    .Minutes = TimeSerial(0, delayMinutes, 0)
                    .StartDate = slot.DaOra
                    .EndDate = slot.AOra
                    .Origin = 3
                    .Note = "Inserita dalla post " & App.EXEName
                End With
                
                oDip.IWriteHour.InsertObjHour mCurrDate, newHour, True
                
                DebugView "inserito codice di ritardo " & newHour.Code & " in data " & mCurrDate
                
                Set newHour = Nothing
                
            Next
            
            
        End With
        
SkipDay:

    Next mCurrDate
    
    If oDip.IWriteHour.ToWrite Then
        somethingHappened = True
        DebugView "scrive in data " & mCurrDate
        oDip.IWriteHour.WriteObjHours
    End If
    
    If LogTrack And somethingHappened Then
        ' scrive il log
        Do While (stackMessage.Count > 0)
            WriteLog stackMessage.Pop
        Loop
    End If
    
    Set stackMessage = Nothing
   
    IMain_Calculate = wpECalcOK
    
Main_EXIT:

    Exit Function
    
Main_ERROR:

    IMain_Calculate = wpECalcError
    
    DebugView "errore"
    
    If InDesign() Then
        Debug.Print Err.Description
        Debug.Assert False
        Resume
    End If
    
    Select Case CatchErr(Err.Description, Err.Number, App.EXEName & ".Main(" & DebugOperation & ")", VbMsgBoxStyle.vbAbortRetryIgnore Or VbMsgBoxStyle.vbQuestion, True)
    Case VbMsgBoxResult.vbRetry
        Resume
    Case VbMsgBoxResult.vbIgnore
        Resume Next
    Case VbMsgBoxResult.vbCancel, VbMsgBoxResult.vbAbort
        Resume Main_EXIT
    End Select

End Function

Private Function IMain_Initialize() As ElabSrvr.CalcResultEnum

    DebugOperation = "Lettura file ini"
    
    profileList = Split(RINI("Parametri", "Lista Profili"), ",")
    delayCode = RINI("Parametri", "Codice Ritardo")
    
    LogTrack = (Val(RINI("Parametri", "Log")) <> 0)
    
    
    IMain_Initialize = wpECalcOK
    
Initialize_EXIT:

    Exit Function
    
Initialize_ERROR:

    Select Case CatchErr(Err.Description, Err.Number, App.EXEName & ".Initialize(" & DebugOperation & ")", VbMsgBoxStyle.vbAbortRetryIgnore Or VbMsgBoxStyle.vbQuestion, True)
    Case vbRetry
        Resume
    Case vbCancel, vbAbort
        IMain_Initialize = False
        Resume Initialize_EXIT
    End Select

End Function

Private Property Let IMain_Phase(ByVal RHS As ElabSrvr.CustElabPhases)

End Property

Private Property Get IMain_Phase() As ElabSrvr.CustElabPhases

End Property

Private Function IMain_Terminate() As ElabSrvr.CalcResultEnum
    IMain_Terminate = wpECalcOK
End Function

