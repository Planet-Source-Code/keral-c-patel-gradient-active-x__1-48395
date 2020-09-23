VERSION 5.00
Begin VB.UserControl ColorChooser 
   ClientHeight    =   510
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5310
   LockControls    =   -1  'True
   ScaleHeight     =   510
   ScaleWidth      =   5310
   ToolboxBitmap   =   "ColorChooser.ctx":0000
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   150
      Left            =   0
      ScaleHeight     =   120
      ScaleWidth      =   5280
      TabIndex        =   1
      Top             =   360
      Width           =   5310
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   150
      Left            =   0
      Picture         =   "ColorChooser.ctx":0312
      ScaleHeight     =   120
      ScaleWidth      =   5280
      TabIndex        =   0
      Top             =   0
      Width           =   5310
   End
End
Attribute VB_Name = "ColorChooser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Const m_def_PickedColor = &H0&
Private m_PickedColor As Variant 'Not OLE_COLOR other wise it creates some Problems
'This is For Events
Event Click()
Event DblClick()
Event KeyDown(KeyCode As Integer, Shift As Integer)
Event KeyPress(KeyAscii As Integer)
Event KeyUp(KeyCode As Integer, Shift As Integer)
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Private Red, Green, Blue


Public Sub About()

    Dim msgret As Byte
    msgret = MsgBox("This Color Chooser Control is Made By Keral.C.Patel." & vbCrLf & "                        ©2003 BlueSoft", vbOKOnly, "ColorChooser Control")

End Sub

Private Sub Picture1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    'This Procedure will display the gradient in the bottom line
    Gradient Picture2, vbWhite, Picture1.Point(X, Y)

End Sub

Private Sub Picture2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    'this procedure will asign the Color to the PickedColor Property
    m_PickedColor = Picture2.Point(X, Y)
    RaiseEvent Click
    PropertyChanged "PickedColor"

End Sub

Private Sub UserControl_Click()

    RaiseEvent Click

End Sub

Private Sub UserControl_DblClick()

    RaiseEvent DblClick

End Sub

Private Sub UserControl_Initialize()

    Gradient Picture2, vbWhite, m_PickedColor

End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)

    RaiseEvent KeyDown(KeyCode, Shift)

End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)

    RaiseEvent KeyPress(KeyAscii)

End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)

    RaiseEvent KeyUp(KeyCode, Shift)

End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    RaiseEvent MouseDown(Button, Shift, X, Y)

End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    RaiseEvent MouseMove(Button, Shift, X, Y)

End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    RaiseEvent MouseUp(Button, Shift, X, Y)

End Sub

Private Sub UserControl_InitProperties()

    'The Default when the control is first placed on form
    m_PickedColor = m_def_PickedColor
    Gradient Picture2, vbWhite, vbBlack

End Sub

Private Sub UserControl_Resize()

    'This will fix the height and width
    UserControl.Height = Picture2.Top + Picture2.Height
    UserControl.Width = Picture2.Width

End Sub

Public Property Get BackColor() As OLE_COLOR

    'This is the Backcolor of the Control
    BackColor = UserControl.BackColor

End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)

    UserControl.BackColor() = New_BackColor
    PropertyChanged "BackColor"

End Property

Public Property Get PickedColor() As Variant

    'This Property is used for ChoosingColor
    'The Get Procedure will Read the Picked Color
    PickedColor = m_PickedColor

End Property

Public Property Let PickedColor(ByVal New_PickedColor As Variant)

    'This Procedure will Write PickedColor
    m_PickedColor = New_PickedColor
    PropertyChanged "PickedColor"

End Property

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    'For Reading Properties From the PropertyBag
    UserControl.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
    m_PickedColor = PropBag.ReadProperty("PickedColor", m_def_PickedColor)

End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    'For saving Properties to the PropertyBag
    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H8000000F)
    Call PropBag.WriteProperty("PickedColor", m_PickedColor, m_def_PickedColor)

End Sub


Private Function Gradient(Pic As PictureBox, SClr, EClr)
Dim SRed, SGreen, SBlue, ERed, EGreen, EBlue, DifR, DifG, DifB, Cntr, Yi
    'This Function is for Drawing Gradients
    Pic.AutoRedraw = True: Pic.ScaleMode = 3
    RGBVal (SClr): SRed = Red: SGreen = Green: SBlue = Blue
    RGBVal (EClr): ERed = Red: EGreen = Green: EBlue = Blue
    DifR = ERed - SRed: DifG = EGreen - SGreen: DifB = EBlue - SBlue
    Cntr = Sqr(2) * Sqr((Pic.ScaleWidth * Pic.ScaleWidth) + (Pic.ScaleHeight * Pic.ScaleHeight))

    For Yi = 0 To Cntr

        SRed = SRed + (DifR / Cntr): If SRed < 0 Then SRed = 0

        SGreen = SGreen + (DifG / Cntr): If SGreen < 0 Then SGreen = 0

        SBlue = SBlue + (DifB / Cntr): If SBlue < 0 Then SBlue = 0

        Pic.Line (Yi, 0)-(0, Yi), RGB(SRed, SGreen, SBlue)

    Next

End Function

Private Function RGBVal(clrRGB As ColorConstants)

    'For Gradient Drawing
    Dim rr, gr, br As Long
    rr = 1: gr = 256: br = 65536
    Dim rest As Long
    rest = clrRGB \ br
    Blue = rest
    clrRGB = clrRGB Mod br

    If Blue < 0 Then Blue = 0

    rest = clrRGB \ gr
    Green = rest
    clrRGB = clrRGB Mod gr

    If Green < 0 Then Green = 0

    rest = clrRGB \ rr
    Red = rest
    clrRGB = clrRGB Mod rr

    If Red < 0 Then Red = 0

End Function


