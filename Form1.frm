VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "API Error Description Locator 3.0a"
   ClientHeight    =   5010
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   7755
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5010
   ScaleWidth      =   7755
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraError 
      Caption         =   " Error Number Only "
      Height          =   675
      Left            =   60
      TabIndex        =   12
      Top             =   120
      Width           =   7635
      Begin VB.CommandButton cmdAction 
         Caption         =   "Stop Search"
         Height          =   315
         Index           =   3
         Left            =   6420
         TabIndex        =   14
         Top             =   240
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdAction 
         Caption         =   "Show It"
         Default         =   -1  'True
         Height          =   315
         Index           =   0
         Left            =   6420
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox txtInput 
         Height          =   315
         Index           =   0
         Left            =   60
         TabIndex        =   13
         Text            =   "Error Number"
         Top             =   240
         Width           =   3555
      End
   End
   Begin VB.Frame fraRange 
      Caption         =   " Range "
      Height          =   4095
      Left            =   60
      TabIndex        =   3
      Top             =   840
      Width           =   7635
      Begin MSComctlLib.ProgressBar pbar 
         Height          =   195
         Left            =   60
         TabIndex        =   10
         Top             =   3840
         Visible         =   0   'False
         Width           =   5835
         _ExtentX        =   10292
         _ExtentY        =   344
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VB.CommandButton cmdAction 
         Caption         =   "Filter"
         Height          =   315
         Index           =   2
         Left            =   6420
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   780
         Width           =   1095
      End
      Begin VB.TextBox txtInput 
         Height          =   285
         Index           =   3
         Left            =   60
         TabIndex        =   2
         Top             =   840
         Width           =   3555
      End
      Begin VB.CommandButton cmdAction 
         Caption         =   "List Them"
         Height          =   315
         Index           =   1
         Left            =   6420
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox txtInput 
         Height          =   285
         Index           =   2
         Left            =   2520
         TabIndex        =   1
         Text            =   "10000"
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox txtInput 
         Height          =   285
         Index           =   1
         Left            =   60
         TabIndex        =   0
         Text            =   "0"
         Top             =   240
         Width           =   1095
      End
      Begin MSComctlLib.ListView lvwRange 
         Height          =   2655
         Left            =   60
         TabIndex        =   11
         Top             =   1140
         Width           =   7455
         _ExtentX        =   13150
         _ExtentY        =   4683
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Error"
            Object.Width           =   1270
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Description"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Count:"
         Height          =   195
         Left            =   5460
         TabIndex        =   8
         Top             =   3840
         Width           =   1035
      End
      Begin VB.Label lblCount 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         Height          =   195
         Left            =   6600
         TabIndex        =   7
         Top             =   3840
         Width           =   915
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "To"
         Height          =   255
         Index           =   1
         Left            =   2520
         TabIndex        =   6
         Top             =   540
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "From"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   5
         Top             =   540
         Width           =   1095
      End
   End
   Begin VB.Menu mbrFile 
      Caption         =   "&File"
      Begin VB.Menu miOnTop 
         Caption         =   "&Always On Top"
         Visible         =   0   'False
      End
      Begin VB.Menu miBreak1 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu miExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mbrHelp 
      Caption         =   "&Help"
      Begin VB.Menu miAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const TXT_ERROR_NUMBER = 0
Private Const TXT_SEARCH_FROM = 1
Private Const TXT_SEARCH_TO = 2
Private Const TXT_FILTER = 3

Private Const CMD_SHOW_IT = 0
Private Const CMD_LIST_THEM = 1
Private Const CMD_FILTER = 2
Private Const CMD_STOP_SEARCH = 3

Private bStop As Boolean

Private Const FORMAT_MESSAGE_FROM_SYSTEM = 4096

Private Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" (ByVal dwFlags As Long, lpSource As Any, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Arguments As Long) As Long
Private Declare Function SetUnhandledExceptionFilter& Lib "kernel32" (ByVal lpTopLevelExceptionFilter As Long)
Private Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long



Function GetAPIErrorDescription(ErrorCode As Long) As String
On Error GoTo HandleError:

    Dim lRet As Long                ' return value
    Dim sAPIError As String         ' buffer
    Dim lErrCode As Long            ' number or code of last error that occurred
    
    ' Get error code (number) of last error
    '  to occur
    '
    lErrCode = ErrorCode 'GetLastError
    
    ' Pre-allocate the buffer
    '
    sAPIError = String$(255, " ")
    
    ' Get the formatted message
    '
    lRet = FormatMessage( _
            FORMAT_MESSAGE_FROM_SYSTEM, _
            ByVal 0&, _
            lErrCode, _
            GetSystemDefaultLCID, _
            sAPIError, _
            Len(sAPIError) - 1, _
            0 _
        )
    
    ' Re-format error string
    '
    sAPIError = Left$(sAPIError, lRet)
    
    ' Return error string
    '
    GetAPIErrorDescription = sAPIError
    
Xit:
    Exit Function

HandleError:
    Resume Next
    
End Function

Private Sub FilterSearch()
    '
    ' Filter is used to limit the results to the matching keyword... wildcards are
    '  not support nor is any form of boolean operation (yet)...
    '
    Dim Message As String
    Dim I As Long
    Dim K As Long
    
    I = 1
    
    ' Update the progress indicator
    '
    pbar.Min = I
    pbar.Max = lvwRange.ListItems.Count
    K = pbar.Max
    
    LockWindowUpdate lvwRange.hwnd
    
    ' Iterate through each item in the list
    '
    Do While (I <= lvwRange.ListItems.Count)
        '
        ' Set the string
        '
        Message = lvwRange.ListItems(I).SubItems(1)
        
        ' Find a match... or not
        '
        If Not SearchStr(CStr(txtInput(TXT_FILTER)), Message) Then
            ' Match not found, remove current item
            '
            lvwRange.ListItems.Remove (I)
            'lvwRange.Visible = False
            lblCount = CStr((CLng(lblCount) - 1))
            
            ' Decrement I by 1 because it will be added again, we want to seach
            '  the current item, which was the next item, but since the previous
            '  was removed, the next becomes the current... understand?
            '
            I = (I - 1)
        End If
        
        ' Allow things to happen
        '
        UpdateProgress K, 10
        'Process I, 250
        
        I = (I + 1)
        K = (K - 1)
    Loop
    
    LockWindowUpdate 0&
    pbar.Value = 1
End Sub

Private Sub QueryMessages()
On Error GoTo HandleError:

    '
    ' If the searchable range is not pure numeric, then we can't do anything,
    '  therefore, we do nothing...
    '
    If (Not IsNumeric(txtInput(TXT_SEARCH_FROM)) Or Not IsNumeric(txtInput(TXT_SEARCH_TO))) Then
        Exit Sub
    End If
    
    ' Clear the list
    '
    lvwRange.ListItems.Clear
    
    Dim Item As ListItem                ' The List Item we'll be working with
    Dim Message As String               ' The actual message Windows returns
    Dim Temp As String                  ' Our Temp storage space
    Dim Count As Long                   ' Number of items returned
    Dim I As Long                       ' Incrementor
    Dim X As Long                       ' Incrementor X
    
    pbar.Visible = True
    
    pbar.Min = CLng(txtInput(TXT_SEARCH_FROM))
    pbar.Max = CLng(txtInput(TXT_SEARCH_TO))
    
    LockWindowUpdate lvwRange.hwnd
    ' cmdStop is there so that the user can stop the lengthy list seach...
    '  that's all
    '
    cmdAction(CMD_STOP_SEARCH).Visible = True
    bStop = False
    
    For I = CLng(txtInput(TXT_SEARCH_FROM)) To CLng(txtInput(TXT_SEARCH_TO))
        '
        ' Get the message... please note, that there are some error codes
        '  which actually raise an "Access Violation"... unfortunately, VB
        '  does not handle those... fortunately I wrote some API which does,
        '  if you want to know more, look up basErrorFilter and basAPI, so
        '  it will catch the error and then we handle here as a "resume next"
        '  If that error gets raised, we just keep running as normal and let
        '  Message = NULL
        '
        Message = GetAPIErrorDescription(I)
        
        If (Message <> "") Then
            '
            ' Check to see if user opted to stop the search
            '
            If bStop Then
                Exit For
            End If
            
            '
            ' A Message was returned
            '
            Temp = (Right$("00000000" & CStr(I), 8))
            
            ' There is information in the Filter field, so we will filter the
            '  current result for a match, and make a boolean decision to display
            '  the message or no
            '
            If (Len(txtInput(TXT_FILTER))) > 0 Then
                If SearchStr(CStr(txtInput(TXT_FILTER)), Message) Then
                    ' Match
                    '
                    Set Item = lvwRange.ListItems.Add(, , Temp)
                        Item.SubItems(1) = ConvertStr(Message)
                        
                    Count = (Count + 1)
                End If
            Else
                ' No need to compare, Filter is empty, display it anyway
                '
                Set Item = lvwRange.ListItems.Add(, , Temp)
                    Item.SubItems(1) = ConvertStr(Message)
                    
                Count = (Count + 1)
            End If
            
            ' Show how many matches were found
            '
            lblCount = Count
        End If
        
        ' Do some more processing (overkill, ya'think?)
        '
        UpdateProgress I, 50
        Process I, 250
    Next
    
Xit:
    cmdAction(CMD_STOP_SEARCH).Visible = False
    LockWindowUpdate 0&
    Exit Sub
    
HandleError:
    If Err.Number = 10000 Then
        Resume Next
    Else
        Resume Next
        'MsgBox "Error: " & Err.Number & ":  " & Err.Description
    End If
    
End Sub

Private Sub Process(I As Long, Value As Long)
    If (I Mod Value) = 0 Then
        '
        ' It really takes a long time to insert each item, there is flicker and
        '  there is lag... so, we even the odds a little... instead of freezing
        '  the application while we search, and instead of letting all of the
        '  messages process, we wait until every value'th iteration and then we
        '  unlock the listview updates, and let the system process messages for
        '  one cycle and then lock the listview again... this allows the list
        '  to refresh, and the user to click the button to stop the search, or
        '  exit the program or whatever
        '
        LockWindowUpdate 0&
        DoEvents
        LockWindowUpdate lvwRange.hwnd
    End If
End Sub

Private Sub UpdateProgress(I As Long, Value As Long)
    If (I Mod Value) = 0 Then
        '
        ' Much of our delay in searching isn't in the inneficiency of VB or my
        '  algorithms, much rather in the constant updating of the progress
        '  indicator ever iteration... so, we wait until every value'th iteration
        '  and then update the progress indicator... 250 is good for searching
        '  and 25-50 is good for filtering... this way the progress bar is still
        '  updated frequently, but doesn't really get in the way of our performance
        '  (too much, anyway)
        '
        pbar.Value = I
    End If
End Sub

Private Sub cmdStop_Click()
    bStop = True
    DoEvents
End Sub

Private Sub ShowMessageDescription()
    If Not IsNumeric(txtInput(TXT_ERROR_NUMBER)) Then
        Exit Sub
    End If
    
    ' Show the message
    '
    MsgBox GetAPIErrorDescription( _
            CLng(txtInput(TXT_ERROR_NUMBER))), _
            Title:="Error: " & CLng(txtInput(TXT_ERROR_NUMBER))
    
End Sub

Private Sub cmdAction_Click(Index As Integer)
    Select Case Index
        Case CMD_SHOW_IT
            ShowMessageDescription
        Case CMD_LIST_THEM
            QueryMessages
        Case CMD_FILTER
            If lvwRange.ListItems.Count = 0 Then
                QueryMessages
            Else
                FilterSearch
            End If
        Case CMD_STOP_SEARCH
            DoEvents
            bStop = True
        Case Else
            ' Do nothing
    
    End Select
End Sub

Private Sub Form_Load()
   
    Dim X As Single
    X = Screen.TwipsPerPixelX
    
    With lvwRange
        .ColumnHeaders(1).Width = (96 * X)
        .ColumnHeaders(2).Width = (.Width - (.ColumnHeaders(1).Width) - (22 * X))
    End With
    
    Call Main
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Call ExitApp
End Sub

Private Sub lvwRange_DblClick()
On Error Resume Next

    '
    ' A listitem was double-clicked, so we pop it into a messagebox, using the
    '  same technique as the single error number lookup, in this case, it's
    '  code reuse at it's best (somewhat)
    '
    Dim J As Long
    Dim K As String
    
    Let J = lvwRange.SelectedItem.Index
    Let K = lvwRange.ListItems(J).SubItems(1)
    
    J = CLng(lvwRange.ListItems(J))
    
    MsgBox K, Title:="Error: " & CLng(J)
    
End Sub

Private Function ConvertStr(Str As String) As String
    Dim Temp As String
    
    '
    ' The windows error description will actually place a vbCrLf at the end of
    '  the string, this is where we remove it.  Actually, not only does this
    '  program pickup error descriptions, but string tables, also... the actual
    '  string tables do not have vbCrLf, I meant to create some way to differentiate
    '  them, but I have no need, if you get an error message, chances are, it's
    '  not a normal string table for use by Windows... other than message purposes,
    '  you'll only pickup a non-error related string if you are doing a range
    '  search...
    '
    Temp = Left$(Str, (Len(Str)))
    
    If Right$(Temp, 2) = vbCrLf Then
        Temp = (Left$(Str, (Len(Str) - 2)))
    End If
    
    ConvertStr = Temp

End Function

Private Function SearchStr(Str As String, Message As String) As Boolean
    '
    ' InStr() wasn't working, so I created my own, at least this works
    '
    Dim I As Long
    
    If (Len(Str) > Len(Message)) Then
        SearchStr = False
        Exit Function
    End If
    
    I = 1
    
    Do While (I < (Len(Message)))
        If (Mid$(LCase$(Message), I, Len(Str)) = LCase$(Str)) Then
            SearchStr = True
            Exit Do
        End If
        
        I = (I + 1)
    Loop
    
End Function

Private Sub lvwRange_GotFocus()
    lvwRange.BackColor = RGB(255, 255, 191)
End Sub

Private Sub lvwRange_LostFocus()
    lvwRange.BackColor = vbWhite
End Sub

Private Sub miAbout_Click()
    MsgBox "Copywrite 2000 Shawn Bullock.  All Rights Reserved", vbInformation, "API Error Description Locator"
End Sub

Private Sub miExit_Click()
    ExitApp
End Sub

Private Sub miOnTop_Click()
    miOnTop.Checked = Not miOnTop.Checked
    
    If miOnTop.Checked = True Then
        '
    Else
        '
    End If
    
End Sub

Private Sub txtInput_Change(Index As Integer)
    '
    ' Do not allow any value greater than 32 bits to be entered
    '
    Select Case Index
        Case TXT_ERROR_NUMBER To TXT_SEARCH_TO
            If IsNumeric(txtInput(Index)) Then
                If txtInput(Index) > ((2 ^ 31) - 1) Then
                    If Len(txtInput(Index)) > 1 Then
                        With txtInput(Index)
                            .SelStart = txtInput(Index).SelStart - 1
                            .SelLength = 1
                            .SelText = ""
                        End With
                        Beep
                    End If
                End If
            End If
    End Select
End Sub

Private Sub txtInput_GotFocus(Index As Integer)
    With txtInput(Index)
        .SelStart = 0
        .SelLength = Len(.Text)
        
        .BackColor = RGB(255, 255, 191)
    End With
    
    Select Case Index
        Case TXT_ERROR_NUMBER
            cmdAction(CMD_SHOW_IT).Default = True
        Case TXT_SEARCH_FROM
            cmdAction(CMD_LIST_THEM).Default = True
        Case TXT_SEARCH_TO
            cmdAction(CMD_LIST_THEM).Default = True
        Case TXT_FILTER
            cmdAction(CMD_FILTER).Default = True
        Case Else
            ' Do Nothing
    End Select
End Sub

Private Sub txtInput_LostFocus(Index As Integer)
    With txtInput(Index)
        .BackColor = vbWhite
    End With
End Sub
