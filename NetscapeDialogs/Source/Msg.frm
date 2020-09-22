VERSION 5.00
Begin VB.Form Msg 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Netscape Dialogs"
   ClientHeight    =   1470
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5880
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Msg.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1470
   ScaleWidth      =   5880
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows-Standard
   Begin NetscapeDialogs.NetscapeButton cmdButton 
      Height          =   330
      Index           =   0
      Left            =   3480
      TabIndex        =   1
      Top             =   1080
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   582
      Picture         =   "Msg.frx":08CA
   End
   Begin NetscapeDialogs.NetscapeButton cmdButton 
      Height          =   330
      Index           =   1
      Left            =   2160
      TabIndex        =   2
      Top             =   1080
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   582
      Picture         =   "Msg.frx":0A8C
   End
   Begin NetscapeDialogs.NetscapeButton cmdButton 
      Height          =   330
      Index           =   2
      Left            =   840
      TabIndex        =   3
      Top             =   1080
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   582
      Caption         =   "OK"
      Picture         =   "Msg.frx":0C4E
   End
   Begin VB.Label lbText 
      BackStyle       =   0  'Transparent
      Caption         =   "Willkommen zu NetscapeDialogs. Diese Dialoge sind dem neuen Unschlagbaren Netscape Navigator 6 nachempfunden"
      Height          =   735
      Left            =   720
      TabIndex        =   0
      Top             =   360
      Width           =   4815
   End
   Begin VB.Image IconImage 
      Height          =   480
      Index           =   3
      Left            =   120
      Picture         =   "Msg.frx":0E10
      Top             =   240
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image IconImage 
      Height          =   480
      Index           =   2
      Left            =   120
      Picture         =   "Msg.frx":16DA
      Top             =   240
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image IconImage 
      Height          =   480
      Index           =   1
      Left            =   120
      Picture         =   "Msg.frx":1FA4
      Top             =   240
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image IconImage 
      Height          =   480
      Index           =   0
      Left            =   120
      Picture         =   "Msg.frx":286E
      Top             =   240
      Visible         =   0   'False
      Width           =   480
   End
End
Attribute VB_Name = "Msg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public MsgResult As Integer
Private ButtonCount As Integer

Sub ResizeMsgbox()

Me.Width = lbText.Left + lbText.Width + 450

If ButtonCount = -1 Then Exit Sub

If ButtonCount = 1 Then
    With cmdButton(0)
        .Left = (Width - .Width) / 2
    End With
End If

End Sub

Sub SetMessage(MessageText As String)
lbText.Caption = MessageText

lbText.AutoSize = True
ResizeMsgbox

End Sub

Private Sub cmdButton_Click(Index As Integer)
MsgResult = Index
Unload Me

End Sub

Public Function ShowMessage(Message As String, Optional Icon As MessageIcon, Optional Style As MessageStyle, Optional Title As String) As MessageResults

    ButtonCount = -1
    SetMessage Message
    
    Select Case Style
        Case ncNormalOK
            SetButton 0, "OK", True, True, True
            ButtonCount = 1
        Case ncOkCancel
            SetButton 0, "OK", True, True, True
            SetButton 1, "Cancel", True, False, False
            ButtonCount = 2
        Case ncYesNo
            SetButton 0, "Yes", True, False, False
            SetButton 1, "No", True, False, False
            ButtonCount = 2
        Case ncOkCancelRetry
            SetButton 0, "Ok", True, True, True
            SetButton 1, "Cancel", True, False, False
            SetButton 2, "Retry", True, False, False
            ButtonCount = 3
        Case ncOkCancelIgnore
            SetButton 0, "Ok", True, True, True
            SetButton 1, "Cancel", True, False, False
            SetButton 2, "Ignore", True, False, False
            ButtonCount = 3
    End Select
    
    ResizeMsgbox
    
    With Msg
        .Caption = Title
        .IconImage(Icon).Visible = True
    End With
    
    Me.Show 1
    
    ShowMessage = MsgResult
    
End Function

Sub SetButton(ButtonIndex As Integer, bCaption As String, cVisible As Boolean, bFontBold As Boolean, bPicture As Boolean)
            With Msg.cmdButton(ButtonIndex)
                .Visible = cVisible
                .Caption = bCaption
                .FontBold = bFontBold
                .PictureShow = bPicture
            End With
End Sub

