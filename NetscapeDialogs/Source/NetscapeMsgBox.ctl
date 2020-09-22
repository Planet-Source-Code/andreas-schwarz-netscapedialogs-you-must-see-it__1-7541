VERSION 5.00
Begin VB.UserControl NetscapeMsgBox 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   ToolboxBitmap   =   "NetscapeMsgBox.ctx":0000
   Begin VB.Image Logo 
      Appearance      =   0  '2D
      BorderStyle     =   1  'Fest Einfach
      Height          =   510
      Left            =   0
      Picture         =   "NetscapeMsgBox.ctx":0312
      Top             =   0
      Width           =   510
   End
End
Attribute VB_Name = "NetscapeMsgBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Public Enum MessageIcon
    NoIcon = 0
    Error = 1
    Question = 2
    Notice = 3
    Warning = 4
End Enum

Public Enum MessageStyle
    ncNormalOK = 0
    ncOkCancel = 1
    ncYesNo = 2
    ncOkCancelRetry = 3
    ncOkCancelIgnore = 4
End Enum

Public Enum MessageResults
    ncOk = 0
    ncYes = 1
    ncNo = 2
    ncCancel = 3
    ncRetry = 4
    ncIgnore = 5
End Enum
    

    
Private Sub UserControl_Resize()
    Height = Logo.Height
    Width = Logo.Width
End Sub

Public Function MessageBox(Message As String, Optional Icon As MessageIcon, Optional Style As MessageStyle, Optional Title As String) As MessageResults

    MessageBox = Msg.ShowMessage(Message, Icon, Style, Title)
        
End Function
