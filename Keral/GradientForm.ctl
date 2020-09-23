VERSION 5.00
Begin VB.UserControl GradientForm 
   ClientHeight    =   2310
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   ControlContainer=   -1  'True
   ScaleHeight     =   154
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   312
   ToolboxBitmap   =   "GradientForm.ctx":0000
   Begin VB.PictureBox PicMain 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2265
      Left            =   300
      ScaleHeight     =   2265
      ScaleWidth      =   4650
      TabIndex        =   0
      Top             =   2310
      Width           =   4650
   End
End
Attribute VB_Name = "GradientForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'**************************************************************************************
'Date:-22/8/2003                    ***************************************************
'Author:- Keral.C.Patel.            ***************************************************
'©BlueSoft 2003                     ***************************************************
'Email:- keral82@keral.com          ***************************************************
'For Questions Email Me             ***************************************************
'**************************************************************************************
Option Explicit
'Public gradcolor As Variant
Const m_def_border = 10
Private m_border As Integer
Const m_def_StartColor = &HFFFFFF
Private m_StartColor  As OLE_COLOR
Const m_def_EndColor = &H0&
Private m_EndColor As OLE_COLOR
'This is For Events
Event Click()
Event DblClick()
Event KeyDown(KeyCode As Integer, Shift As Integer)
Event KeyPress(KeyAscii As Integer)
Event KeyUp(KeyCode As Integer, Shift As Integer)
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'This is for copying all the gradient fill to the control
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Const SRCCOPY = &HCC0020
Private Red, Green, Blue


Public Sub About()

    'This will Show About Message
    Dim msgret As Byte
    msgret = MsgBox("This Gradient-Form Control is Created By Keral.C.Patel." & vbCrLf & "                        ©2003 BlueSoft", vbOKOnly, "GradientForm Control")

End Sub

Private Sub UserControl_Click()

    'This is Very Simple Just Raise an Event and the Client application will handle all
    'the other things
    RaiseEvent Click

End Sub

Private Sub UserControl_DblClick()

    RaiseEvent DblClick

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

Private Sub UserControl_Initialize()

    'This is the Behaviour of the Control when it is Initialized
    'For example when opening a Project or When Opening the Control's Designer

    'setting scalemode
    PicMain.ScaleMode = 3
    'Setting Border
    Border = m_border
    'Calling procedure which handles other things
    UserControl_Resize

End Sub

Private Sub UserControl_InitProperties()

    'This is the Behaviour of the Control when it is First of all Placed on
    'The form from the Tool-Box. (Siting a Control)
    'Initializing Border
    m_border = m_def_border
    'Now setting Gradient
    m_StartColor = m_def_StartColor
    m_EndColor = m_def_EndColor
    Gradient PicMain, m_StartColor, m_EndColor
    Call procBitBlt

End Sub

Private Sub UserControl_Paint()

    'So that the Gradient is maintained when our control goes from Design-time to Run-time
    UserControl_Resize

End Sub

Private Sub UserControl_Resize()

    On Error Resume Next 'If Something Bad Happens
    
    'setting height and width of picMain
    PicMain.Height = UserControl.ScaleHeight
    PicMain.Width = UserControl.ScaleWidth
    PicMain.Left = UserControl.ScaleWidth
    PicMain.Top = UserControl.ScaleHeight
    'Now Drawing the Gradient
    Gradient PicMain, m_StartColor, m_EndColor
    'Here Bitblt Code which will Copy Gradient From PicMain to the Control
    'Accordiong the Border Width Specified
    Call procBitBlt

End Sub

Public Property Get Border() As Integer

    'For Reading the Border Property
    Border = m_border
    UserControl_Resize

End Property

Public Property Let Border(ByVal New_Border As Integer)

    'For Writing the Border Property to the PropertyBag
    m_border = New_Border
    PropertyChanged "Border"
    'Refreshing Control so that Changes are seen at the time of editng the Properties
    UserControl_Resize

End Property

Public Property Get StartColor() As OLE_COLOR
Attribute StartColor.VB_Description = "Changes or Sets the Starting Color in Gradient."
Attribute StartColor.VB_ProcData.VB_Invoke_Property = "PropertyPage1"

    'For Reading the StartColor
    'Here If you specify you property as OLE_COLOR then VB by default places
    'the color choosing Dialog in the Property Window
    StartColor = m_StartColor
    Gradient PicMain, m_StartColor, m_EndColor
    'Here also Bitblt code if neccesary
    Call procBitBlt

End Property

Public Property Let StartColor(ByVal New_StartColor As OLE_COLOR)

    'For writing the Start Color Property
    'The following Line will asign NewValue to the Private Variable
    m_StartColor = New_StartColor
    'For Drawing Gradient with new color
    Gradient PicMain, m_StartColor, m_EndColor
    'Refreshing control
    Call procBitBlt
    'You have to tell Visual Basic that this Property has Changed so that
    'It can Fire the Write Properties Event and Save this Properrty to make
    'It persistent
    PropertyChanged "StartColor"

End Property

Public Property Get EndColor() As OLE_COLOR
Attribute EndColor.VB_Description = "Changes or Sets the Ending Color in Gradient."
Attribute EndColor.VB_ProcData.VB_Invoke_Property = "PropertyPage2"

    'same as start color
    EndColor = m_EndColor
    Gradient PicMain, m_StartColor, m_EndColor
    'Here also Bitblt code if neccesary
    Call procBitBlt

End Property

Public Property Let EndColor(ByVal New_EndColor As OLE_COLOR)

    'same as startcolor
    m_EndColor = New_EndColor
    PropertyChanged "EndColor"
    Gradient PicMain, m_StartColor, m_EndColor
    'Here also Bitblt code if neccesary
    Call procBitBlt

End Property

Public Property Get BackColor() As OLE_COLOR

    'This is for the background color of the UserControl
    BackColor = UserControl.BackColor

End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)

    'This is for saving the BackGround Color of Usercontrol
    UserControl.BackColor() = New_BackColor
    PropertyChanged "BackColor"

End Property

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    'This event will fire whenever our control Needs to read Properties and work
    'accordingly to it. It Stores all the Values in an Object named PropertyBag.
    m_StartColor = PropBag.ReadProperty("StartColor", m_def_StartColor)
    m_EndColor = PropBag.ReadProperty("EndColor", m_def_EndColor)
    m_border = PropBag.ReadProperty("Border", m_def_border)
    UserControl.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)

End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    'This event will only fire If any property has been changed in
    'Property let Procedures and and if we have specified it with
    '  PropertyChanged "PropertyName"   Method
    Call PropBag.WriteProperty("StartColor", m_StartColor, m_def_StartColor)
    Call PropBag.WriteProperty("EndColor", m_EndColor, m_def_EndColor)
    Call PropBag.WriteProperty("Border", m_border, m_def_border)
    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H8000000F)

End Sub

Private Sub procBitBlt()

    'This if for refreshing control according to the proerties which have been set

    UserControl.Cls ' This will clear all the content of the UserControl
    'This Following Lines are I think Explanatory for those who have used BitBlt
    BitBlt UserControl.hDC, 0, 0, PicMain.Width, m_border, PicMain.hDC, 0, 0, SRCCOPY
    BitBlt UserControl.hDC, 0, 0, m_border, PicMain.Height, PicMain.hDC, 0, 0, SRCCOPY
    BitBlt UserControl.hDC, 0, PicMain.Height - m_border, PicMain.Width, m_border, PicMain.hDC, 0, PicMain.Height - m_border, SRCCOPY
    BitBlt UserControl.hDC, PicMain.Width - m_border, 0, m_border, PicMain.Height, PicMain.hDC, PicMain.Width - m_border, 0, SRCCOPY

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

