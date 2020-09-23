VERSION 5.00
Object = "*\AOCX.vbp"
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6090
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8250
   LinkTopic       =   "Form1"
   ScaleHeight     =   6090
   ScaleWidth      =   8250
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin GradientOCX.GradientForm GradientForm1 
      Height          =   6015
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   10610
      StartColor      =   16761024
      EndColor        =   8454143
      BackColor       =   16777215
      Begin VB.HScrollBar HScroll1 
         Height          =   255
         Left            =   1320
         Max             =   220
         TabIndex        =   6
         Top             =   1440
         Value           =   10
         Width           =   5415
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "About"
         Height          =   375
         Left            =   3480
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   5040
         Width           =   1455
      End
      Begin GradientOCX.ColorChooser ColorChooser2 
         Height          =   510
         Left            =   1440
         TabIndex        =   2
         Top             =   4080
         Width           =   5310
         _ExtentX        =   9366
         _ExtentY        =   900
         BackColor       =   16777215
      End
      Begin GradientOCX.ColorChooser ColorChooser1 
         Height          =   510
         Left            =   1440
         TabIndex        =   1
         Top             =   3240
         Width           =   5310
         _ExtentX        =   9366
         _ExtentY        =   900
         BackColor       =   16777215
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Border"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   3600
         TabIndex        =   7
         Top             =   1080
         Width           =   720
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "EndColor"
         Height          =   195
         Left            =   1440
         TabIndex        =   4
         Top             =   3840
         Width           =   645
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "StartColor"
         Height          =   195
         Left            =   1440
         TabIndex        =   3
         Top             =   3000
         Width           =   690
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub ColorChooser1_Click()

    GradientForm1.StartColor = ColorChooser1.PickedColor

End Sub

Private Sub ColorChooser2_Click()

    GradientForm1.EndColor = ColorChooser2.PickedColor

End Sub

Private Sub Command2_Click()

    GradientForm1.About
    ColorChooser1.About

End Sub

Private Sub Form_Load()

    GradientForm1.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight

End Sub

Private Sub HScroll1_Change()
GradientForm1.Border = HScroll1.Value
End Sub
