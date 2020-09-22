VERSION 5.00
Begin VB.UserControl NetscapeButton 
   ClientHeight    =   3120
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4695
   ScaleHeight     =   208
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   313
   ToolboxBitmap   =   "NetscapeButton.ctx":0000
   Begin VB.Timer SelectFlagTimer 
      Interval        =   1
      Left            =   3480
      Top             =   360
   End
   Begin VB.Image Arrow 
      Height          =   135
      Left            =   2400
      Picture         =   "NetscapeButton.ctx":0312
      Top             =   960
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.Line CornerTop 
      BorderColor     =   &H00808080&
      X1              =   184
      X2              =   200
      Y1              =   112
      Y2              =   112
   End
   Begin VB.Line RightLine 
      BorderColor     =   &H00808080&
      X1              =   200
      X2              =   200
      Y1              =   160
      Y2              =   104
   End
   Begin VB.Line CornerLine 
      BorderColor     =   &H00808080&
      X1              =   192
      X2              =   200
      Y1              =   168
      Y2              =   160
   End
   Begin VB.Line BottomLine 
      BorderColor     =   &H00808080&
      X1              =   48
      X2              =   192
      Y1              =   168
      Y2              =   168
   End
   Begin VB.Line LeftLine 
      BorderColor     =   &H80000005&
      X1              =   48
      X2              =   48
      Y1              =   104
      Y2              =   168
   End
   Begin VB.Line TopLine 
      BorderColor     =   &H80000005&
      X1              =   48
      X2              =   200
      Y1              =   104
      Y2              =   104
   End
   Begin VB.Label kCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Button"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   1080
      TabIndex        =   0
      Top             =   720
      Width           =   465
   End
End
Attribute VB_Name = "NetscapeButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private Type ColorRGB
    R As Integer
    G As Integer
    B As Integer
End Type

Private ClickColor(1 To 2) As ColorRGB
'Standard-Eigenschaftswerte:
'Const m_def_Caption = ""
Const m_def_ToolTipText = ""
Const m_def_PictureShow = 0
'Eigenschaftsvariablen:
'Dim m_Caption As String
Dim m_ToolTipText As String
Dim m_PictureShow As Boolean
'Ereignisdeklarationen:
Event Click() 'MappingInfo=UserControl,UserControl,-1,Click
Attribute Click.VB_Description = "Tritt auf, wenn der Benutzer eine Maustaste über einem Objekt drückt und wieder losläßt."


Private Selected As Boolean




    
Private Sub InitializeColors()
    SetColor ClickColor(1), 255, 255, 255
    SetColor ClickColor(2), 104, 104, 104
    
End Sub

Private Sub InitializeLine()

    SetLine TopLine, 1, ScaleWidth - 1, 0, 0
    SetLine LeftLine, 0, 0, 1, ScaleHeight - 1
    SetLine BottomLine, 1, ScaleWidth - 1, ScaleHeight - 1, ScaleHeight - 1
    SetLine RightLine, ScaleWidth - 1, ScaleWidth - 1, 1, ScaleHeight - 1
    SetLine CornerLine, ScaleWidth - 2, ScaleWidth - 1, ScaleHeight - 2, ScaleHeight - 2
    SetLine CornerTop, ScaleWidth - 2, ScaleWidth - 1, 1, 1
    
End Sub

Private Sub SetColor(Color As ColorRGB, R As Integer, G As Integer, B As Integer)
    With Color
        .R = R
        .B = B
        .G = G
    End With
End Sub

Private Sub SetLineColor(LineControl As Line, Color As ColorRGB)
    With LineControl
        .BorderColor = RGB(Color.R, Color.G, Color.B)
    End With
End Sub

Private Sub SetState(Flag As Integer)
    '//////////////
    '// 1 = Unpressed
    '// 2 = PRESSED
    
    If Flag = 1 Then
        SetLineColor TopLine, ClickColor(1)
        SetLineColor LeftLine, ClickColor(1)
        SetLineColor BottomLine, ClickColor(2)
        SetLineColor RightLine, ClickColor(2)
        SetLineColor CornerLine, ClickColor(2)
        SetLineColor CornerTop, ClickColor(2)
    ElseIf Flag = 2 Then
        SetLineColor TopLine, ClickColor(2)
        SetLineColor LeftLine, ClickColor(2)
        SetLineColor BottomLine, ClickColor(1)
        SetLineColor RightLine, ClickColor(1)
        SetLineColor CornerLine, ClickColor(1)
        SetLineColor CornerTop, ClickColor(1)
    End If
    

    
End Sub


Private Sub kCaption_Click()
    UserControl_Click
End Sub

Private Sub kCaption_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    UserControl_MouseDown Button, Shift, x, y
End Sub


Private Sub kCaption_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Selected = True Then Exit Sub
    kCaption.FontUnderline = True
    kCaption.ForeColor = RGB(0, 0, 150)
    Selected = True
End Sub

Private Sub kCaption_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    UserControl_MouseUp Button, Shift, x, y
End Sub


Private Sub UserControl_Initialize()
    InitializeColors
    InitializeLine
    SetState 1
End Sub





Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    SetState 2
End Sub


Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Selected = False
    kCaption.FontUnderline = False
    kCaption.ForeColor = RGB(0, 0, 0)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    SetState 1
End Sub


Private Sub UserControl_Resize()
    InitializeLine
    With kCaption
        .Top = (ScaleHeight - .Height) / 2
        .Left = (ScaleWidth - .Width) / 2
    End With
    With Arrow
        .Top = (ScaleHeight - .Height) / 2
        .Left = ScaleWidth - .Width - 20
    End With
End Sub

Private Sub SetLine(LineControl As Line, X1 As Integer, X2 As Integer, Y1 As Integer, Y2 As Integer)
With LineControl
    .X1 = X1
    .X2 = X2
    .Y1 = Y1
    .Y2 = Y2
End With
End Sub

'ACHTUNG! DIE FOLGENDEN KOMMENTIERTEN ZEILEN NICHT ENTFERNEN ODER VERÄNDERN!
'MappingInfo=UserControl,UserControl,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Gibt die Hintergrundfarbe zurück, die verwendet wird, um Text und Grafik in einem Objekt anzuzeigen, oder legt diese fest."
    BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    UserControl.BackColor() = New_BackColor
    PropertyChanged "BackColor"
End Property
'
''ACHTUNG! DIE FOLGENDEN KOMMENTIERTEN ZEILEN NICHT ENTFERNEN ODER VERÄNDERN!
''MappingInfo=Caption,Caption,-1,Caption
'Public Property Get Caption() As String
'    Caption = kCaption.Caption
'End Property
'
'Public Property Let Caption(ByVal New_Caption As String)
'    kCaption.Caption() = New_Caption
'    PropertyChanged "Caption"
'    With kCaption
'        .Top = (ScaleHeight - .Height) / 2
'        .Left = (ScaleWidth - .Width) / 2
'    End With
'End Property

Private Sub UserControl_Click()
    RaiseEvent Click
End Sub

'ACHTUNG! DIE FOLGENDEN KOMMENTIERTEN ZEILEN NICHT ENTFERNEN ODER VERÄNDERN!
'MappingInfo=UserControl,UserControl,-1,hDC
Public Property Get hdc() As Long
Attribute hdc.VB_Description = "Gibt eine Zugriffsnummer (von Microsoft Windows) für den Gerätekontext des Objekts zurück."
    hdc = UserControl.hdc
End Property

'ACHTUNG! DIE FOLGENDEN KOMMENTIERTEN ZEILEN NICHT ENTFERNEN ODER VERÄNDERN!
'MappingInfo=Arrow,Arrow,-1,Picture
Public Property Get Picture() As Picture
Attribute Picture.VB_Description = "Gibt eine Grafik zurück, die in einem Steuerelement angezeigt werden soll, oder legt diese fest."
    Set Picture = Arrow.Picture
End Property

Public Property Set Picture(ByVal New_Picture As Picture)
    Set Arrow.Picture = New_Picture
    PropertyChanged "Picture"
End Property

'ACHTUNG! DIE FOLGENDEN KOMMENTIERTEN ZEILEN NICHT ENTFERNEN ODER VERÄNDERN!
'MemberInfo=13,0,0,
Public Property Get ToolTipText() As String
Attribute ToolTipText.VB_Description = "Gibt den Text zurück, der angezeigt wird, wenn die Maus über dem Steuerelement verweilt, oder legt den Text fest."
    ToolTipText = m_ToolTipText
End Property

Public Property Let ToolTipText(ByVal New_ToolTipText As String)
    m_ToolTipText = New_ToolTipText
    PropertyChanged "ToolTipText"
End Property

'ACHTUNG! DIE FOLGENDEN KOMMENTIERTEN ZEILEN NICHT ENTFERNEN ODER VERÄNDERN!
'MemberInfo=0,0,0,0
Public Property Get PictureShow() As Boolean
    PictureShow = m_PictureShow
End Property

Public Property Let PictureShow(ByVal New_PictureShow As Boolean)
    m_PictureShow = New_PictureShow
    PropertyChanged "PictureShow"
    If m_PictureShow = True Then
        Arrow.Visible = True
    With Arrow
        .Top = (ScaleHeight - .Height) / 2
        .Left = ScaleWidth - .Width - 20
    End With

    Else
        Arrow.Visible = False
    End If
End Property

'Eigenschaften für Benutzersteuerelement initialisieren
Private Sub UserControl_InitProperties()
    m_ToolTipText = m_def_ToolTipText
    m_PictureShow = m_def_PictureShow
'    m_Caption = m_def_Caption
End Sub

'Eigenschaftenwerte vom Speicher laden
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    UserControl.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
    kCaption.Caption = PropBag.ReadProperty("Caption", "Button")
    Set Picture = PropBag.ReadProperty("Picture", Nothing)
    m_ToolTipText = PropBag.ReadProperty("ToolTipText", m_def_ToolTipText)
    m_PictureShow = PropBag.ReadProperty("PictureShow", m_def_PictureShow)
'    m_Caption = PropBag.ReadProperty("Caption", m_def_Caption)
    kCaption.FontBold = PropBag.ReadProperty("FontBold", 0)
    kCaption.Caption = PropBag.ReadProperty("Caption", "Button")
End Sub

'Eigenschaftenwerte in den Speicher schreiben
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H8000000F)
    Call PropBag.WriteProperty("Caption", kCaption.Caption, "Button")
    Call PropBag.WriteProperty("Picture", Picture, Nothing)
    Call PropBag.WriteProperty("ToolTipText", m_ToolTipText, m_def_ToolTipText)
    Call PropBag.WriteProperty("PictureShow", m_PictureShow, m_def_PictureShow)
'    Call PropBag.WriteProperty("Caption", m_Caption, m_def_Caption)
    Call PropBag.WriteProperty("FontBold", kCaption.FontBold, 0)
End Sub
'
''ACHTUNG! DIE FOLGENDEN KOMMENTIERTEN ZEILEN NICHT ENTFERNEN ODER VERÄNDERN!
''MemberInfo=13,0,0,
'Public Property Get Caption() As String
'    Caption = m_Caption
'End Property
'
'Public Property Let Caption(ByVal New_Caption As String)
'    m_Caption = New_Caption
'    PropertyChanged "Caption"
'End Property

'ACHTUNG! DIE FOLGENDEN KOMMENTIERTEN ZEILEN NICHT ENTFERNEN ODER VERÄNDERN!
'MappingInfo=kCaption,kCaption,-1,FontBold
Public Property Get FontBold() As Boolean
Attribute FontBold.VB_Description = "Gibt Schriftstile für Fettschrift zurück oder legt diese fest."
    FontBold = kCaption.FontBold
End Property

Public Property Let FontBold(ByVal New_FontBold As Boolean)
    kCaption.FontBold() = New_FontBold
    PropertyChanged "FontBold"
End Property

'ACHTUNG! DIE FOLGENDEN KOMMENTIERTEN ZEILEN NICHT ENTFERNEN ODER VERÄNDERN!
'MappingInfo=kCaption,kCaption,-1,Caption
Public Property Get Caption() As String
Attribute Caption.VB_Description = "Gibt den Text zurück, der in der Titelleiste eines Objekts oder unter dem Symbol eines Objekts angezeigt wird, oder legt diesen fest."
    Caption = kCaption.Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    kCaption.Caption() = New_Caption
    PropertyChanged "Caption"
End Property

