VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Message Box Generator- By Lewis"
   ClientHeight    =   7155
   ClientLeft      =   3615
   ClientTop       =   810
   ClientWidth     =   5025
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7155
   ScaleWidth      =   5025
   Begin VB.Frame fraButtons 
      Height          =   735
      Left            =   945
      TabIndex        =   38
      Top             =   3555
      Width           =   4020
      Begin VB.OptionButton optButtonType 
         Caption         =   "OkOnly"
         Height          =   240
         Index           =   0
         Left            =   90
         TabIndex        =   44
         Top             =   180
         Value           =   -1  'True
         Width           =   915
      End
      Begin VB.OptionButton optButtonType 
         Caption         =   "OkCancel"
         Height          =   240
         Index           =   1
         Left            =   1035
         TabIndex        =   43
         Top             =   180
         Width           =   1050
      End
      Begin VB.OptionButton optButtonType 
         Caption         =   "AbortRetryIgnore"
         Height          =   240
         Index           =   2
         Left            =   2385
         TabIndex        =   42
         Top             =   180
         Width           =   1545
      End
      Begin VB.OptionButton optButtonType 
         Caption         =   "YesNoCancel"
         Height          =   240
         Index           =   3
         Left            =   1035
         TabIndex        =   41
         Top             =   450
         Width           =   1320
      End
      Begin VB.OptionButton optButtonType 
         Caption         =   "YesNo"
         Height          =   240
         Index           =   4
         Left            =   90
         TabIndex        =   40
         Top             =   450
         Width           =   870
      End
      Begin VB.OptionButton optButtonType 
         Caption         =   "RetryCancel"
         Height          =   240
         Index           =   5
         Left            =   2385
         TabIndex        =   39
         Top             =   450
         Width           =   1185
      End
   End
   Begin VB.OptionButton optImgOption 
      Height          =   600
      Index           =   1
      Left            =   945
      Picture         =   "Form1.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   37
      Top             =   2925
      Width           =   645
   End
   Begin VB.OptionButton optImgOption 
      Height          =   600
      Index           =   2
      Left            =   1575
      Picture         =   "Form1.frx":0884
      Style           =   1  'Graphical
      TabIndex        =   36
      Top             =   2925
      Width           =   645
   End
   Begin VB.OptionButton optImgOption 
      Height          =   600
      Index           =   3
      Left            =   2205
      Picture         =   "Form1.frx":0CC6
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   2925
      Width           =   645
   End
   Begin VB.OptionButton optImgOption 
      Height          =   600
      Index           =   4
      Left            =   2835
      Picture         =   "Form1.frx":1108
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   2925
      Width           =   645
   End
   Begin VB.Frame fraModal 
      Height          =   465
      Left            =   945
      TabIndex        =   31
      Top             =   4320
      Width           =   2310
      Begin VB.OptionButton optModalType 
         Caption         =   "Application"
         Height          =   240
         Index           =   0
         Left            =   90
         TabIndex        =   33
         Top             =   180
         Value           =   -1  'True
         Width           =   1140
      End
      Begin VB.OptionButton optModalType 
         Caption         =   "System"
         Height          =   240
         Index           =   1
         Left            =   1260
         TabIndex        =   32
         Top             =   180
         Width           =   870
      End
   End
   Begin VB.CheckBox chkOnTop 
      Caption         =   "Stay On Top"
      Height          =   240
      Left            =   3465
      TabIndex        =   30
      Top             =   4455
      Width           =   1320
   End
   Begin VB.CheckBox chkAlignRight 
      Caption         =   "Text Align Right"
      Height          =   240
      Left            =   3555
      TabIndex        =   29
      Top             =   2970
      Width           =   1500
   End
   Begin VB.CheckBox chkImageRight 
      Caption         =   "Image On Right"
      Height          =   240
      Left            =   3555
      TabIndex        =   28
      Top             =   3285
      Width           =   1500
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "Help/Info"
      Height          =   330
      Left            =   90
      TabIndex        =   27
      Top             =   6750
      Width           =   1140
   End
   Begin VB.Frame fraType 
      Height          =   465
      Left            =   945
      TabIndex        =   22
      Top             =   4725
      Width           =   4020
      Begin VB.OptionButton optOption 
         Caption         =   "Input Box"
         Height          =   195
         Index           =   2
         Left            =   2520
         TabIndex        =   25
         Top             =   180
         Width           =   1050
      End
      Begin VB.OptionButton optOption 
         Caption         =   "Function"
         Height          =   195
         Index           =   1
         Left            =   1305
         TabIndex        =   24
         Top             =   180
         Width           =   1050
      End
      Begin VB.OptionButton optOption 
         Caption         =   "Sub"
         Height          =   195
         Index           =   0
         Left            =   405
         TabIndex        =   23
         Top             =   180
         Value           =   -1  'True
         Width           =   645
      End
   End
   Begin VB.Frame fraHelp 
      Height          =   870
      Left            =   945
      TabIndex        =   14
      Top             =   1980
      Width           =   3975
      Begin VB.TextBox txtHelpFile 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Left            =   810
         TabIndex        =   20
         Top             =   495
         Width           =   3120
      End
      Begin VB.TextBox txtContext 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Left            =   2565
         TabIndex        =   16
         Top             =   180
         Width           =   1365
      End
      Begin VB.CheckBox chkHelp 
         Caption         =   "Help Button"
         Height          =   195
         Left            =   135
         TabIndex        =   15
         Top             =   225
         Width           =   1140
      End
      Begin VB.Label lblHelpFile 
         Caption         =   "HelpFile:"
         Enabled         =   0   'False
         Height          =   240
         Left            =   90
         TabIndex        =   19
         Top             =   540
         Width           =   825
      End
      Begin VB.Label lblContext 
         Caption         =   "Help Context ID:"
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   1350
         TabIndex        =   17
         Top             =   225
         Width           =   1230
      End
   End
   Begin VB.CommandButton cmdCopy 
      Caption         =   "Copy Code"
      Height          =   330
      Left            =   3690
      TabIndex        =   7
      Top             =   6750
      Width           =   1230
   End
   Begin VB.CommandButton cmdGenerate 
      Caption         =   "Generate Code"
      Height          =   330
      Left            =   2160
      TabIndex        =   6
      Top             =   6750
      Width           =   1500
   End
   Begin VB.TextBox txtCode 
      Height          =   1455
      Left            =   45
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Top             =   5220
      Width           =   4965
   End
   Begin VB.TextBox txtTitle 
      Height          =   330
      Left            =   945
      TabIndex        =   2
      Text            =   "Lewis's Cool MsgBox Generator"
      Top             =   45
      Width           =   3975
   End
   Begin VB.CommandButton cmdTest 
      Caption         =   "Test"
      Height          =   330
      Left            =   1305
      TabIndex        =   1
      Top             =   6750
      Width           =   825
   End
   Begin VB.TextBox txtMessage 
      Height          =   1050
      Left            =   945
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "Form1.frx":154A
      Top             =   405
      Width           =   4020
   End
   Begin VB.Frame fraDefault 
      Height          =   510
      Left            =   945
      TabIndex        =   8
      Top             =   1440
      Width           =   3975
      Begin VB.OptionButton optDefButton 
         Caption         =   "Button4"
         Height          =   240
         Index           =   3
         Left            =   2970
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   180
         Width           =   870
      End
      Begin VB.OptionButton optDefButton 
         Caption         =   "Button3"
         Height          =   240
         Index           =   2
         Left            =   2025
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   180
         Width           =   870
      End
      Begin VB.OptionButton optDefButton 
         Caption         =   "Button2"
         Height          =   240
         Index           =   1
         Left            =   1035
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   180
         Width           =   870
      End
      Begin VB.OptionButton optDefButton 
         Caption         =   "Button1"
         Height          =   240
         Index           =   0
         Left            =   90
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   180
         Width           =   870
      End
      Begin VB.TextBox txtDefault 
         Height          =   330
         Left            =   45
         TabIndex        =   21
         Top             =   135
         Visible         =   0   'False
         Width           =   3840
      End
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Image:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   2
      Left            =   270
      TabIndex        =   47
      Top             =   2925
      Width           =   555
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Buttons:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   3
      Left            =   180
      TabIndex        =   46
      Top             =   3645
      Width           =   735
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Modal:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   5
      Left            =   315
      TabIndex        =   45
      Top             =   4455
      Width           =   555
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Type:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   270
      TabIndex        =   26
      Top             =   4860
      Width           =   510
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Help:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   6
      Left            =   360
      TabIndex        =   18
      Top             =   2070
      Width           =   510
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Default:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   4
      Left            =   180
      TabIndex        =   9
      Top             =   1620
      Width           =   735
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Message:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   1
      Left            =   90
      TabIndex        =   4
      Top             =   405
      Width           =   870
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Title:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   0
      Left            =   450
      TabIndex        =   3
      Top             =   90
      Width           =   555
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Msgbox() Generator By Lewis Miller (aka Deth,Dethbomb)

'Some things "todo":
'1) Parse message a character at a time to replace items, the replace() function will
'   replace items inside strings that shouldnt be replaced.
'2) Add support to include function calls inside the string, currently those will need
'   to be manually added.


'Note: Some of the index's of the controls on the form are
       'manipulated in such a way to use the index as a operator
       'when calulating the message box flags, be careful
       'when changing anything to use correct index of control.
       'Also note: since speed is not a major requirement, the IIf()
       'function works nicely for in-line comparisons.

'use this instead of Chr$(34)
Private Const vbQuote As String = """"
Private Const vbComma As String = ", "

'storage for last selected values
Dim lngLastButtonType      As Long
Dim lngLastDefaultButton   As Long
Dim lngLastImageChosen     As Long
Dim lngLastModalType       As Long

'storage to hold the end result of all msgbox button numbers
Dim lngButtonResult As Long

Private Sub Form_Load()

'get saved settings
    txtTitle.Text = GetSetting("VB", "MsgBox Gen", "title", txtTitle)
    txtMessage.Text = GetSetting("VB", "MsgBox Gen", "msg", txtMessage)

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

'save settings
    SaveSetting "VB", "MsgBox Gen", "title", txtTitle
    SaveSetting "VB", "MsgBox Gen", "msg", txtMessage

End Sub


Private Sub chkAlignRight_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

'add or subtract new value depending on selection
    lngButtonResult = IIf(chkAlignRight, lngButtonResult + 524288, lngButtonResult - 524288)

End Sub

'disable or enable help frame contents depending on selection
Private Sub chkHelp_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    txtContext.Enabled = chkHelp
    txtHelpFile.Enabled = chkHelp
    lblHelpFile.Enabled = chkHelp
    lblContext.Enabled = chkHelp
    txtContext.BackColor = IIf(chkHelp, vbWhite, vbButtonFace)
    txtHelpFile.BackColor = txtContext.BackColor
    lngButtonResult = IIf(chkHelp, lngButtonResult + 16384, lngButtonResult - 16384)
    If chkHelp Then
        MsgBox "In Order To Use This You Have To Specify A 'Help Context ID' And A 'Help File' Also!", vbCritical
    End If

End Sub

Private Sub chkImageRight_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

'add or subtract newvalue depending on selection
    lngButtonResult = IIf(chkImageRight, lngButtonResult + 1048576, lngButtonResult - 1048576)

End Sub

Private Sub chkOnTop_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

'add or subtract value
    lngButtonResult = IIf(chkOnTop, lngButtonResult + 65536, lngButtonResult - 65536)

End Sub

'hide some things if its in Inputbox mode
Private Sub optoption_Click(Index As Integer)

    optImgOption(1).Visible = Not optOption(2)
    optImgOption(2).Visible = Not optOption(2)
    optImgOption(3).Visible = Not optOption(2)
    optImgOption(4).Visible = Not optOption(2)

    Label1(2).Visible = Not optOption(2)
    Label1(3).Visible = Not optOption(2)
    Label1(5).Visible = Not optOption(2)

    fraButtons.Visible = Not optOption(2)
    fraModal.Visible = Not optOption(2)

    chkAlignRight.Visible = Not optOption(2)
    chkImageRight.Visible = Not optOption(2)
    chkOnTop.Visible = Not optOption(2)

    txtDefault.Visible = optOption(2)
    txtDefault.ZOrder
    
End Sub

'msgbox test button , generate a msgbox based on selections
Private Sub cmdTest_Click()

    Dim RetMsg As VbMsgBoxResult
    Dim StrTemp As String, strMessageCode As String

    On Error GoTo ErrHandle                                'needed to enable checking for bad helpfile

    If chkHelp Then                                        'helpfile included?
        If (txtHelpFile = "") Or (txtContext = "") Then    'check to make sure there not empty
            MsgBox "Invalid Help File Or ContextID!", vbCritical
            Exit Sub
        End If
    End If
    
    strMessageCode = Replace$(Replace$(Replace$(Replace$(txtMessage, vbCrLf, " "), "|", vbCrLf), "& " & vbQuote & vbQuote & ",", ","), vbQuote & vbQuote & " & ", "")

    Select Case True

            'sub
        Case optOption(0).Value = True
            If chkHelp Then
                MsgBox strMessageCode, lngButtonResult, txtTitle, txtHelpFile, txtContext
            Else
                MsgBox strMessageCode, lngButtonResult, txtTitle
            End If

            'function
        Case optOption(1).Value = True
            If chkHelp Then
                RetMsg = MsgBox(strMessageCode, lngButtonResult, txtTitle, txtHelpFile, txtContext)
            Else
                RetMsg = MsgBox(strMessageCode, lngButtonResult, txtTitle)
            End If

            MsgBox "The result of your msgbox function is: " & Choose(RetMsg, " vbOk (1)", " vbCancel (2)", " vbAbort (3)", " vbRetry (4)", " vbIgnore (5)", " vbYes (6)", " vbNo (7)"), vbInformation

            'inputbox
        Case optOption(2).Value = True
            If chkHelp Then
                StrTemp = InputBox(strMessageCode, txtTitle, txtDefault, , , txtHelpFile, CLng(txtContext))
            Else
                StrTemp = InputBox(strMessageCode, txtTitle, txtDefault)
            End If

            If StrPtr(StrTemp) = 0 Then                    'cancelled, this is how you detect it...
                MsgBox "Inputbox Was Cancelled!", vbCritical
            Else
                MsgBox "-The result of your Inputbox() function is- " & vbCrLf & vbCrLf & IIf(StrTemp = "", "(Nothing or Empty)" & vbCrLf, StrTemp) & vbCrLf, vbInformation
            End If

    End Select

    Exit Sub
ErrHandle:
    MsgBox Err.Description, vbCritical, "MsgBox Gen"

End Sub

'same as above except it generates code instead of msgbox
Private Sub cmdGenerate_Click()

    Dim StrHelp As String, StrTemp As String

    If chkHelp Then
        If txtHelpFile <> "" And txtContext <> "" Then
            StrHelp = vbComma & vbQuote & txtHelpFile & vbQuote & vbComma & vbQuote & txtContext & vbQuote
        Else
            chkHelp.Value = 0
            Call chkHelp_MouseUp(0, 0, 0, 0)
        End If
    End If

    'Replace$ any quotes in the message text and dump it into a temp string
    'we use Chr$(34) instead of vbQuote, because some people may forget to define it.
    StrTemp = Replace$(txtMessage, vbQuote, vbQuote & " & Chr$(34) & " & vbQuote)
    If Left$(StrTemp, 3) = ", " & vbQuote Then StrTemp = Mid$(StrTemp, 3)

    'Replace$ any unintentional vbcrlf's left by a multiline textbox
    StrTemp = Replace$(StrTemp, vbCrLf, vbQuote & " & vbCrLf & " & vbQuote)
    StrTemp = Trim$(StrTemp)  'you didnt want those spaces at the end anyway, did you?  :)

    'Replace pipe symbols with vbcrlf
    StrTemp = Replace$(StrTemp, "|", vbQuote & " & vbCrLf & " & vbQuote)
    If Left$(StrTemp, 3) = ", v" Then StrTemp = Mid$(StrTemp, 3)
    If Left$(StrTemp, 2) = ", " Then StrTemp = Mid$(StrTemp, 3)

    Select Case True

            'regular msgbox, (sub - no return)
        Case optOption(0).Value = True
            StrTemp = "Msgbox " & vbQuote & StrTemp & vbQuote & vbComma & CStr(lngButtonResult) & vbComma & vbQuote & txtTitle & vbQuote & StrHelp

            'input msgbox, (function - expects return)
        Case optOption(1).Value = True
            StrTemp = "Dim RetMsg As vbMsgBoxResult" & vbCrLf & "RetMsg = Msgbox(" & vbQuote & StrTemp & vbQuote & vbComma & CStr(lngButtonResult) & vbComma & vbQuote & txtTitle & vbQuote & StrHelp & ")"

            'input box
        Case optOption(2).Value = True
            StrTemp = "Dim StrTemp As String" & vbCrLf & "StrTemp = Inputbox(" & vbQuote & StrTemp & vbQuote & vbComma & vbQuote & txtTitle & vbQuote & vbComma & vbQuote & txtDefault & vbQuote & IIf(StrHelp <> "", vbComma & vbComma & StrHelp, "") & ")" & vbCrLf & "If Not (StrPtr(strTemp) And Len(StrTemp)) Then Exit Sub"

    End Select

    '******************************************************************
    'remove junk quotes cause by replacing the pipe symbols with vbcrlf's
    'and quotes ,and then finally add it to the code text window
    '******************************************************************
    txtCode.Text = Replace$(Replace$(StrTemp, "& " & vbQuote & vbQuote & ",", ","), vbQuote & vbQuote & " & ", "")

End Sub

'copy to clip
Private Sub cmdCopy_Click()

    Clipboard.Clear
    Clipboard.SetText txtCode.Text

End Sub

Private Sub cmdHelp_Click()

    'yes, this msgbox was made with this generator :)
    MsgBox "This generator has the ability to add vbcrlf's into the generated code automatically by inserting the pipe '|' symbol into the message area," & vbCrLf & vbCrLf & " This causes a break in the text, at areas you use it." & vbCrLf & vbCrLf & "Quotes Can Also Be Used Any Where You Wish, They Are Auto Replaced With Code." & vbCrLf & vbCrLf & " Leave the help file section alone unless you know what your doing." & vbCrLf & vbCrLf, 64, "Lewis's Cool MsgBox Generator"

End Sub

Private Sub optButtonType_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    'add new value and subtract old value
    lngButtonResult = (lngButtonResult + Index) - lngLastButtonType&
    'store new value
    lngLastButtonType = CLng(Index)

End Sub

Private Sub optDefButton_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    'add new value and subtract old value
    lngButtonResult = lngButtonResult + (Index * 256) - lngLastDefaultButton
    'store new value
    lngLastDefaultButton = (Index * 256)

End Sub

Private Sub optImgOption_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)


    'add new value and subtract old value
    lngButtonResult = lngButtonResult + (Index * 16) - lngLastImageChosen
    'store new value
    lngLastImageChosen = (Index * 16)

End Sub

Private Sub optModalType_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    'add new value and subtract old value
    lngButtonResult = lngButtonResult + (Index * 4096) - lngLastModalType
    'store new value
    lngLastModalType = (Index * 4096)

End Sub
