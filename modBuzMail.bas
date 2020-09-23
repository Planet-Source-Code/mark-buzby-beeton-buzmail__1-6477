Attribute VB_Name = "modBuzMail"
Type MailRecip
    Address As String
    Name As String
    Type As Integer
End Type

Type MailAttach
    Name As String
    FileName As String
End Type
    
Type Mail
    ID As String
    RecipCount As Integer
    Recips() As MailRecip
    AttachCount As Integer
    Attach() As MailAttach
    Subject As String
    Text As String
    From As MailRecip
    DateReceived As Date
    Unread As Boolean
End Type

Public Sub MailAddress(ThisMsg As Mail)
'uses the MS MAPI address book to add recipients to ThisMsg

On Error GoTo errorcatch

With frmBuzMail.BuzMessages
    .Compose
    If ThisMsg.RecipCount > 0 Then
        For dummy = 1 To ThisMsg.RecipCount
            .RecipIndex = dummy - 1
            .RecipDisplayName = ThisMsg.Recips(dummy).Name
            .RecipType = ThisMsg.Recips(dummy).Type
            .RecipAddress = ThisMsg.Recips(dummy).Address
        Next
    End If
    .Show
    
    While ThisMsg.RecipCount > 0
        MailRemoveRecip ThisMsg, 1
    Wend
    
    If .RecipCount > 0 Then
        For recadd = 0 To .RecipCount - 1
            .RecipIndex = recadd
            .ResolveName
            MailAddRecip ThisMsg, .RecipDisplayName, .RecipAddress, .RecipType
        Next
    End If

End With

Exit Sub
errorcatch:
Select Case Err
Case 32001
    'user cancelled
    Exit Sub
Case Else
    MsgBox Err.Description, vbExclamation, "Error " & Err
End Select
End Sub


Public Sub MailConfig()
On Error GoTo errorcatch
Debug.Print "MailConfig: Running Applet"
X = Shell("control mlcfg32.cpl", 1)
Exit Sub
errorcatch:
MsgBox Err.Description, vbExclamation, "Error " & Err
End Sub



Public Function MailCount(UnreadOnly As Boolean) As Long
On Error GoTo errorcatch
'this function checks the mail and returns the number of items found
frmBuzMail.BuzMessages.FetchUnreadOnly = UnreadOnly
Debug.Print "MailCount: Fetching Mail: UnreadOnly = " + CStr(UnreadOnly)
frmBuzMail.BuzMessages.Fetch
MailCount = frmBuzMail.BuzMessages.MsgCount
Exit Function
errorcatch:
MsgBox Err.Description, vbExclamation, "Error " & Err

End Function

Public Sub MailAddAttach(ByRef ThisMail As Mail, DispName As String, FileName As String)
On Error GoTo errorcatch
CurrentAttachCount = ThisMail.AttachCount
CurrentAttachCount = CurrentAttachCount + 1
ThisMail.AttachCount = CurrentAttachCount
ReDim Preserve ThisMail.Attach(CurrentAttachCount)
ThisMail.Attach(CurrentAttachCount).Name = DispName
ThisMail.Attach(CurrentAttachCount).FileName = FileName
Debug.Print "MailAddAttach: Added " + DispName
Exit Sub
errorcatch:
MsgBox Err.Description, vbExclamation, "Error " & Err
End Sub

Public Sub MailAddRecip(ByRef ThisMail As Mail, DispName As String, Address As String, RecipType As Integer)
On Error GoTo errorcatch
CurrentRecipCount = ThisMail.RecipCount
CurrentRecipCount = CurrentRecipCount + 1
ThisMail.RecipCount = CurrentRecipCount
ReDim Preserve ThisMail.Recips(CurrentRecipCount)
ThisMail.Recips(CurrentRecipCount).Name = DispName
ThisMail.Recips(CurrentRecipCount).Address = Address
ThisMail.Recips(CurrentRecipCount).Type = RecipType
Debug.Print "MailAddRecip: Added " + DispName
Exit Sub
errorcatch:
MsgBox Err.Description, vbExclamation, "Error " & Err
End Sub






Public Sub MailContents(ByRef ThisMail As Mail, Subject As String, Message As String)
On Error GoTo errorcatch
ThisMail.Subject = Subject
ThisMail.Text = Message
Debug.Print "MailContents: Contents Set"
Exit Sub
errorcatch:
MsgBox Err.Description, vbExclamation, "Error " & Err
End Sub

Public Function MailDelete(MailNum As Long) As Boolean
On Error GoTo errorcatch
'deletes a message from the mailbox - using the index of it.
If MailNum > frmBuzMail.BuzMessages.MsgCount Then
    MailDelete = False
    Debug.Print "MailDelete: No such mail"
    Exit Function
End If

frmBuzMail.BuzMessages.MsgIndex = MailNum - 1
frmBuzMail.BuzMessages.Delete 0
Debug.Print "MailDelete: Mail Deleted"

MailDelete = True
Exit Function
errorcatch:
MsgBox Err.Description, vbExclamation, "Error " & Err

End Function

Public Function MailDeleteID(MailID As String) As Boolean
On Error GoTo errorcatch
thisid = 0
foundit = False

With frmBuzMail.BuzMessages

While foundit = False And thisid < .MsgCount
    .MsgIndex = thisid
    If .MsgID = MailID Then
        foundit = True
    Else
        thisid = thisid + 1
    End If
Wend

End With

If foundit = True Then
    MailDeleteID = MailDelete(CLng(thisid) + 1)
Else
    Debug.Print "MailDeleteID: No such mail"
    MailDeleteID = False
End If
Exit Function
errorcatch:
MsgBox Err.Description, vbExclamation, "Error " & Err

End Function


Public Function MailDetach(ThisMsg As Mail, FileNum As Long) As String
On Error GoTo errorcatch
'returns the filename of the attached file
If FileNum > ThisMsg.AttachCount Then
    MailDetach = ""
    Debug.Print "MailDetach: No such attachment"
    Exit Function
End If

Dim dummail As Mail
dummail = MailGetID(ThisMsg.ID)

If dummail.ID = "" Then
    MailDetach = ""
    Debug.Print "MailDetach: No such mail"
    Exit Function
End If

If FileNum > dummail.AttachCount Then
    MailDetach = ""
    Debug.Print "MailDetach: No such attachment"
    Exit Function
End If

MailDetach = dummail.Attach(FileNum).FileName
Debug.Print "MailDetach: Detached " + CStr(MailDetach)
Exit Function
errorcatch:
MsgBox Err.Description, vbExclamation, "Error " & Err

End Function



Public Function MailGet(MsgNum As Long) As Mail
'reads an item of mail and returns the mail object for it

On Error GoTo errorcatch
If MsgNum > frmBuzMail.BuzMessages.MsgCount Then
    Debug.Print "MailGet: No such mail"
    Exit Function
End If

frmBuzMail.BuzMessages.MsgIndex = MsgNum - 1

With frmBuzMail.BuzMessages

    MailGet.Unread = .MsgRead = False
    MailGet.ID = .MsgID
    MailGet.Subject = .MsgSubject
    MailGet.Text = .MsgNoteText
    MailGet.From.Name = .MsgOrigDisplayName
    MailGet.From.Address = .MsgOrigAddress
    MailGet.From.Type = 0
    MailGet.DateReceived = .MsgDateReceived
    MailGet.AttachCount = .AttachmentCount

    If .AttachmentCount > 0 Then
        ReDim Preserve MailGet.Attach(.AttachmentCount)
        
        For getattach = 0 To .AttachmentCount - 1
            .AttachmentIndex = getattach
            MailGet.Attach(getattach + 1).Name = .AttachmentName
            MailGet.Attach(getattach + 1).FileName = .AttachmentPathName
        Next
    End If
    
End With
Exit Function
errorcatch:
MsgBox Err.Description, vbExclamation, "Error " & Err
End Function

Public Function MailGetID(ByRef MailID As String) As Mail
'find the mail with a specific ID
On Error GoTo errorcatch
thisid = 0
foundit = False

With frmBuzMail.BuzMessages

While foundit = False And thisid < .MsgCount
    .MsgIndex = thisid
    If .MsgID = MailID Then
        foundit = True
    Else
        thisid = thisid + 1
    End If
Wend

End With

If foundit = True Then
    MailGetID = MailGet(CLng(thisid) + 1)
Else
    Debug.Print "MailGetID: No such mail"
End If
Exit Function
errorcatch:
MsgBox Err.Description, vbExclamation, "Error " & Err

End Function



Public Function MailLogoff() As Boolean
On Error GoTo errorcatch

MailLogoff = False
frmBuzMail.BuzSession.SignOff
Unload frmBuzMail

ok:
MailLogoff = True
Debug.Print "MailLogoff: Success"
fail:
Exit Function

errorcatch:
Select Case Err
Case 32053
    'not logged on
    Resume ok
Case Else
    MsgBox Err.Description, vbExclamation, "Error " & Err
End Select

End Function

Public Function MailLogon(Username As String, Password As String) As Boolean
On Error GoTo errorcatch

MailLogon = False
Load frmBuzMail
frmBuzMail.Hide

frmBuzMail.BuzSession.Username = Username
frmBuzMail.BuzSession.Password = Password
frmBuzMail.BuzSession.SignOn

frmBuzMail.BuzMessages.SessionID = frmBuzMail.BuzSession.SessionID

ok:
MailLogon = True
Debug.Print "MailLogon: Success"
fail:
Exit Function

errorcatch:
Select Case Err
Case 32050
    'already signed on
    Resume ok
Case Else
    MsgBox Err.Description, vbExclamation, "Error " & Err
End Select

End Function

Public Sub MailRemoveAttach(ByRef ThisMail As Mail, AttachNum As Integer)
On Error GoTo errorcatch
If ThisMail.AttachCount = 0 Then
    Debug.Print "MailRemoveAttach: ERROR: No Attachments"
    Exit Sub
End If
If AttachNum > ThisMail.AttachCount Then
    Debug.Print "MailRemoveAttach: ERROR: No such attachment"
    Exit Sub
End If

If AttachNum = ThisMail.AttachCount Then
    ThisMail.AttachCount = ThisMail.AttachCount - 1

Else
    For squash = AttachNum To ThisMail.AttachCount - 1
    
        ThisMail.Attach(squash).Name = ThisMail.Attach(squash + 1).Name
        ThisMail.Attach(squash).FileName = ThisMail.Attach(squash + 1).FileName
    Next
    ThisMail.AttachCount = ThisMail.AttachCount - 1
    
End If
Debug.Print "MailRemoveAttach: Success"
Exit Sub
errorcatch:
MsgBox Err.Description, vbExclamation, "Error " & Err

End Sub

Public Sub MailRemoveRecip(ByRef ThisMail As Mail, RecipNum As Integer)
On Error GoTo errorcatch
If ThisMail.RecipCount = 0 Then
    Debug.Print "MailRemoveRecip: ERROR: No Recipients"
    Exit Sub
End If
If RecipNum > ThisMail.RecipCount Then
    Debug.Print "MailRemoveRecip: ERROR: No such recipient"
    Exit Sub
End If

If RecipNum = ThisMail.RecipCount Then
    ThisMail.RecipCount = ThisMail.RecipCount - 1

Else
    For squash = RecipNum To ThisMail.RecipCount - 1
    
        ThisMail.Recips(squash).Name = ThisMail.Recips(squash + 1).Name
        ThisMail.Recips(squash).Address = ThisMail.Recips(squash + 1).Address
        ThisMail.Recips(squash).Type = ThisMail.Recips(squash + 1).Type
    Next
    ThisMail.RecipCount = ThisMail.RecipCount - 1
End If
Debug.Print "MailRemoveRecip: Success"
Exit Sub
errorcatch:
MsgBox Err.Description, vbExclamation, "Error " & Err

End Sub

Public Function MailReply(ThisMsg As Mail) As Boolean
On Error GoTo errorcatch
'takes the 'from' name and places it in the recipient index
If ThisMsg.From.Address = "" Then
    MailReply = False
    Debug.Print "MailReply: FAILED"
    Exit Function
End If

MailAddRecip ThisMsg, ThisMsg.From.Name, ThisMsg.From.Address, 1
MailReply = True

Exit Function
errorcatch:
MsgBox Err.Description, vbExclamation, "Error " & Err

End Function

Public Function MailSend(ByRef ThisMail As Mail) As Integer
'returns  0: OK, -1: No recipients, -2: Not logged on, >0 : can't resolve recipient x

On Error GoTo errorcatch

MailSend = 0

If ThisMail.RecipCount = 0 Then
    MailSend = -1
    Debug.Print "MailSend: ERROR: No Recipients"
    Exit Function
End If

frmBuzMail.BuzMessages.Compose
frmBuzMail.BuzMessages.MsgNoteText = ThisMail.Text
frmBuzMail.BuzMessages.MsgSubject = ThisMail.Subject

For RecipCount = 1 To ThisMail.RecipCount

    frmBuzMail.BuzMessages.RecipIndex = RecipCount - 1
    frmBuzMail.BuzMessages.RecipDisplayName = ThisMail.Recips(RecipCount).Name
    frmBuzMail.BuzMessages.RecipAddress = ThisMail.Recips(RecipCount).Address
    frmBuzMail.BuzMessages.ResolveName
    frmBuzMail.BuzMessages.RecipType = ThisMail.Recips(RecipCount).Type
Next
Debug.Print "MailSend: Added Recipients"

If ThisMail.AttachCount <> 0 Then

    frmBuzMail.BuzMessages.MsgNoteText = frmBuzMail.BuzMessages.MsgNoteText + String(ThisMail.AttachCount + 2, " ")
    For AttachCount = 1 To ThisMail.AttachCount
    
        frmBuzMail.BuzMessages.AttachmentIndex = AttachCount - 1
        frmBuzMail.BuzMessages.AttachmentName = ThisMail.Attach(AttachCount).Name
        frmBuzMail.BuzMessages.AttachmentPathName = ThisMail.Attach(AttachCount).FileName
        frmBuzMail.BuzMessages.AttachmentPosition = CLng(Len(ThisMail.Text) + AttachCount)
        Next
    Debug.Print "MailSend: Added Attachments"
Else
    Debug.Print "MailSend: No Attachments"
End If

'frmBuzMail.BuzMessages.MsgIndex = -1
ThisMail.ID = frmBuzMail.BuzMessages.MsgID
ThisMail.From.Name = frmBuzMail.BuzMessages.MsgOrigDisplayName
ThisMail.From.Address = frmBuzMail.BuzMessages.MsgOrigAddress
ThisMail.From.Type = 0

frmBuzMail.BuzMessages.Send

Debug.Print "MailSend: Mail Sent"

Exit Function

errorcatch:
Select Case Err
Case 32014, 32001
    'can't resolve name
    MailSend = RecipCount
    Debug.Print "MailSend: ERROR: Can't resolve name '" + frmBuzMail.BuzMessages.RecipDisplayName + "'"
Case 32011
    'can't find attachment
    MailSend = 1000 + AttachCount
    Debug.Print "MailSend: ERROR: Can't find attachment '" + frmBuzMail.BuzMessages.AttachmentPathName + "'"
Case 32053
    MailSend = -2
    Debug.Print "MailSend: Mail not logged on"
Case 380
    MailSend = -3
    Debug.Print "MailSend: Possible invalid RecipType (should be >0)"
Case Else
    MsgBox Err.Description, vbExclamation, "Error " & Err
End Select
End Function



