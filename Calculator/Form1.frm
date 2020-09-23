VERSION 5.00
Begin VB.Form form1 
   BackColor       =   &H00FF8080&
   BorderStyle     =   0  'None
   Caption         =   "Calculator"
   ClientHeight    =   4830
   ClientLeft      =   150
   ClientTop       =   495
   ClientWidth     =   4320
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   18
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "Form1.frx":030A
   ScaleHeight     =   4830
   ScaleWidth      =   4320
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   4200
      Top             =   360
   End
   Begin VB.CommandButton Number 
      BackColor       =   &H00C0C0FF&
      Caption         =   "."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   10
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   4200
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000001&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   495
      Left            =   840
      MaxLength       =   11
      TabIndex        =   32
      Top             =   120
      Width           =   3375
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFF80&
      Caption         =   "="
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   4
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   4200
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFF80&
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   3
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   3480
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFF80&
      Caption         =   "/"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   2
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   3480
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFF80&
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   2760
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFF80&
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   2760
      Width           =   615
   End
   Begin VB.CommandButton MemClear 
      BackColor       =   &H00FFC0FF&
      Caption         =   "MC"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   26
      ToolTipText     =   "Memory Clear"
      Top             =   1440
      Width           =   615
   End
   Begin VB.CommandButton AddMem 
      BackColor       =   &H00FFC0FF&
      Caption         =   "M+"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   25
      ToolTipText     =   "Memory Add"
      Top             =   1440
      Width           =   615
   End
   Begin VB.CommandButton Cube 
      BackColor       =   &H0000FFFF&
      Caption         =   "x3"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   24
      ToolTipText     =   "Cube"
      Top             =   840
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton invTan 
      BackColor       =   &H0000FFFF&
      Caption         =   "Tan-1"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   23
      ToolTipText     =   "Inverse Tangent"
      Top             =   840
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton invCos 
      BackColor       =   &H0000FFFF&
      Caption         =   "Cos-1"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   22
      ToolTipText     =   "Inverse Cosine"
      Top             =   840
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton invSin 
      BackColor       =   &H0000FFFF&
      Caption         =   "Sin-1"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   21
      ToolTipText     =   "Inverse Sine"
      Top             =   840
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton Shift 
      BackColor       =   &H00FF8080&
      Caption         =   "Shift"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   20
      ToolTipText     =   "Inverse Trigonometry"
      Top             =   840
      Width           =   615
   End
   Begin VB.CommandButton Sine 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Sin"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   19
      ToolTipText     =   "Sine"
      Top             =   840
      Width           =   615
   End
   Begin VB.CommandButton Tangent 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Tan"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   18
      ToolTipText     =   "Tangent"
      Top             =   840
      Width           =   615
   End
   Begin VB.CommandButton Cosine 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Cos"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   17
      ToolTipText     =   "Cosine"
      Top             =   840
      Width           =   615
   End
   Begin VB.CommandButton SquareRoot 
      BackColor       =   &H00FFC0FF&
      Caption         =   "SqRt"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "Square Root"
      Top             =   1440
      Width           =   615
   End
   Begin VB.CommandButton Square 
      BackColor       =   &H00C0FFFF&
      Caption         =   "x2"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "Square"
      Top             =   840
      Width           =   615
   End
   Begin VB.CommandButton Off 
      BackColor       =   &H008080FF&
      Caption         =   "OFF"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   2040
      Width           =   615
   End
   Begin VB.CommandButton Recall 
      BackColor       =   &H00FFC0FF&
      Caption         =   "RCL"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Recall the value stored"
      Top             =   1440
      Width           =   615
   End
   Begin VB.CommandButton Store 
      BackColor       =   &H00FFC0FF&
      Caption         =   "STO"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Store the value"
      Top             =   1440
      Width           =   615
   End
   Begin VB.CommandButton Sign 
      BackColor       =   &H00C0C0FF&
      Caption         =   "+/-"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   4200
      Width           =   615
   End
   Begin VB.CommandButton Clear 
      BackColor       =   &H008080FF&
      Caption         =   "AC"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "All Clear"
      Top             =   2040
      Width           =   615
   End
   Begin VB.CommandButton Number 
      BackColor       =   &H00C0C0FF&
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   9
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   2040
      Width           =   615
   End
   Begin VB.CommandButton Number 
      BackColor       =   &H00C0C0FF&
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   8
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2040
      Width           =   615
   End
   Begin VB.CommandButton Number 
      BackColor       =   &H00C0C0FF&
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   7
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2040
      Width           =   615
   End
   Begin VB.CommandButton Number 
      BackColor       =   &H00C0C0FF&
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   6
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2760
      Width           =   615
   End
   Begin VB.CommandButton Number 
      BackColor       =   &H00C0C0FF&
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   5
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2760
      Width           =   615
   End
   Begin VB.CommandButton Number 
      BackColor       =   &H00C0C0FF&
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   4
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2760
      Width           =   615
   End
   Begin VB.CommandButton Number 
      BackColor       =   &H00C0C0FF&
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   3
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3480
      Width           =   615
   End
   Begin VB.CommandButton Number 
      BackColor       =   &H00C0C0FF&
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   2
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3480
      Width           =   615
   End
   Begin VB.CommandButton Number 
      BackColor       =   &H00C0C0FF&
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3480
      Width           =   615
   End
   Begin VB.CommandButton Number 
      BackColor       =   &H00C0C0FF&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4200
      Width           =   615
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   1320
      TabIndex        =   35
      Top             =   840
      Width           =   1815
   End
   Begin VB.Label CmdMEM 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   120
      TabIndex        =   34
      ToolTipText     =   "Memory Indicator"
      Top             =   240
      Width           =   495
   End
   Begin VB.Menu muEdit 
      Caption         =   "&Edit"
      Begin VB.Menu muCopy 
         Caption         =   "&Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu muPaste 
         Caption         =   "&Paste"
         Shortcut        =   ^V
      End
   End
   Begin VB.Menu muView 
      Caption         =   "&View"
      Begin VB.Menu muStandard 
         Caption         =   "&Standard"
      End
      Begin VB.Menu muScientific 
         Caption         =   "Sc&ientific"
      End
      Begin VB.Menu mnuFileDash1 
         Caption         =   "-"
      End
      Begin VB.Menu muReadme 
         Caption         =   "View Readme"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu muSkin 
      Caption         =   "S&kin"
      Begin VB.Menu muSkin1 
         Caption         =   "Skin &1"
         Shortcut        =   ^Q
      End
      Begin VB.Menu muSkin2 
         Caption         =   "Skin &2"
         Shortcut        =   ^W
      End
      Begin VB.Menu muSkin3 
         Caption         =   "Skin &3"
         Shortcut        =   ^E
      End
      Begin VB.Menu muSkin4 
         Caption         =   "Skin &4"
         Shortcut        =   ^R
      End
   End
End
Attribute VB_Name = "form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim DecPt As Integer
Dim opnre As Integer
Dim Cleartxt As Integer

Dim Value As Double
Dim Choice As Integer
Dim txt As Double

Private Sub AddMem_Click()
txt = txt + Text1
If txt <> 0 Then
    CmdMEM.Caption = "M"
  Else
    CmdMEM.Caption = ""
End If
End Sub

Private Sub Clear_Click()
Text1.Text = ""
Value = 0
DecPt = 0
opnre = 0
Cleartxt = 0
Text1 = 0

Sign.Enabled = False
invSin.Enabled = True
invCos.Enabled = True

Cube.Enabled = False
Square.Enabled = False
SquareRoot.Enabled = False
End Sub

Private Sub Command1_Click(Index As Integer)

If opnre = 0 Or Index = 4 Then
            If Choice = 0 Then
                 Value = Value + Val(Text1)
            ElseIf Choice = 1 Then
                 Value = Value - Val(Text1)
            ElseIf Choice = 2 Then
                If Val(Text1) = 0 Then
                    MsgBox ("SORRY CANNOT DIVIDE BY ZERO")
                    Exit Sub
                Else
                    Value = Value / Val(Text1)
                End If
            ElseIf Choice = 3 Then
                 Value = Value * Val(Text1)
            End If
            Text1 = Str(Value)
            Cleartxt = 0
End If
        opnre = 1
        Choice = Index
        DecPt = 0
End Sub

Private Sub Cosine_Click()
Text1.Text = Cos(Text1 * 3.14159265358979 / 180)
End Sub

Private Sub Cube_Click()
If Text1.Text <> 0 Then
      Text1.Text = (Text1.Text) ^ 3
      Cube.Enabled = True
   Else
      Cube.Enabled = False
End If
End Sub

Private Sub Form_Load()
Text1 = 0
Value = 0
Choice = 0
Cleartxt = 0
opnre = 0
Shift.Visible = False
Sine.Visible = False
invSin.Visible = False
Cosine.Visible = False
invCos.Visible = False
Tangent.Visible = False
invTan.Visible = False
Square.Visible = False
End Sub

Private Sub invCos_Click()
If Text1.Text = "" Then
Text1.Text = "Error"
invCos.Enabled = False
invSin.Enabled = False
Else
Text1.Text = (Atn(-Text1 / Sqr(-Text1 * Text1 + 1)) * 180 / 3.14159265358979) + (2 * Atn(1) * 180 / 3.14159265358979)
invCos.Enabled = False
invSin.Enabled = False
End If
End Sub

Private Sub invSin_Click()
If Text1.Text = "" Then
Text1.Text = "Error"
invSin.Enabled = False
invCos.Enabled = False
Else
Text1.Text = Atn(Text1 / Sqr(-Text1 * Text1 + 1)) * 180 / 3.14159265358979
invSin.Enabled = False
invCos.Enabled = False
End If
End Sub

Private Sub invTan_Click()
Text1.Text = Atn(Text1) * 180 / 3.14159265358979
End Sub

Private Sub MemClear_Click()
txt = 0
CmdMEM.Caption = ""
End Sub

Private Sub muCopy_Click()
If TypeOf Screen.ActiveControl Is TextBox Then _
   Clipboard.SetText Screen.ActiveControl.SelText
End Sub

Private Sub muPaste_Click()
If TypeOf Screen.ActiveControl Is TextBox Then Screen.ActiveControl.SelText = Clipboard.GetText()
End Sub

Private Sub muReadme_Click()
Form2.Show
'Me.Hide
End Sub

Private Sub muScientific_Click()
Label1.Visible = False
Shift.Visible = True
Sine.Visible = True
Cosine.Visible = True
Tangent.Visible = True
Square.Visible = True
End Sub

Private Sub muSkin1_Click()
form1.Picture = Form4.Picture
End Sub

Private Sub muSkin2_Click()
form1.Picture = Form2.Picture
End Sub

Private Sub muSkin3_Click()
form1.Picture = Form3.Picture
End Sub

Private Sub muSkin4_Click()
form1.Picture = Form5.Picture
End Sub

Private Sub muStandard_Click()
Shift.Visible = False
Sine.Visible = False
invSin.Visible = False
Cosine.Visible = False
invCos.Visible = False
Tangent.Visible = False
invTan.Visible = False
Square.Visible = False
Cube.Visible = False
Label1.Visible = True
End Sub

Private Sub Number_Click(Index As Integer)
If Choice = 4 Then
    Value = 0
    Text1 = " "
    Choice = 0
End If
opnre = 0
  If Cleartxt = 0 Then
    Text1 = " "
  End If
    Cleartxt = 1
    If Number(Index).Caption <> "." Then
            If Text1 <> " 0" Then
                Text1 = Text1 & Number(Index).Caption
            Else
                Text1 = " " & Number(Index).Caption
            End If
            If Text1 < 1 And Text1 > -1 Then
                invSin.Enabled = True
                invCos.Enabled = True
              Else
                invSin.Enabled = False
                invCos.Enabled = False
            End If
     Else
            If DecPt = 0 Then
                Text1 = Text1 & "."
                DecPt = 1
            Else
                MsgBox ("ILLEGAL MOVE")
            End If
    End If
  
Cube.Enabled = True
Square.Enabled = True
SquareRoot.Enabled = True
Sign.Enabled = True
End Sub

Private Sub Off_Click()
Dim Msg, Style, Title, Response, MyString
Msg = "Thank You for Using my calculator. " & Chr(13) & Chr(13) & "Do you want to Exit?"
Style = vbYesNo + vbQuestion + vbDefaultButton2
Title = "Exit"

Response = MsgBox(Msg, Style, Title)
If Response = vbYes Then
     End
   MyString = "Yes"
Else
   MyString = "No"
End If

End Sub

Private Sub Recall_Click()
If txt = 0 Then
       Text1.Text = "0"
       SquareRoot.Enabled = False
       Cube.Enabled = False
       Square.Enabled = False
    Else
       Text1.Text = txt
       SquareRoot.Enabled = True
       Cube.Enabled = True
       Square.Enabled = True
End If
End Sub

Private Sub Shift_Click()
Tangent.Visible = Not Tangent.Visible
Cosine.Visible = Not Cosine.Visible
Sine.Visible = Not Sine.Visible
invSin.Visible = Not invSin.Visible
invCos.Visible = Not invCos.Visible
invTan.Visible = Not invTan.Visible
Cube.Visible = Not Cube.Visible
End Sub

Private Sub Sign_Click()
If Text1 < 0 Then
        Text1.Text = -Text1
    Else
        Text1.Text = -Text1
End If
End Sub

Private Sub Sine_Click()
Text1.Text = Sin(Text1 * 3.14159265358979 / 180)
End Sub

Private Sub Square_Click()
If Text1.Text <> 0 Then
     Text1.Text = (Text1) ^ 2
     Square.Enabled = True
  Else
     Square.Enabled = False
End If
End Sub

Private Sub SquareRoot_Click()
If Text1.Text <> 0 Then
Text1.Text = Sqr(Text1.Text)
SquareRoot.Enabled = True
Else
SquareRoot.Enabled = False
End If
End Sub

Private Sub Store_Click()
txt = Text1
If txt <> 0 Then
    CmdMEM.Caption = "M"
  Else
    CmdMEM.Caption = ""
End If
End Sub

Private Sub Tangent_Click()
Text1.Text = Tan(Text1 * 3.14159265358979 / 180)
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii >= Asc("0") And KeyAscii <= ("9") Then
KeyAscii = 0
ElseIf KeyAscii < Asc("0") Or KeyAscii > ("9") Then
KeyAscii = 0
End If
End Sub

Private Sub Text1_Mousemove(Button As Integer, Shift As Integer, x As Single, Y As Single)
muCopy.Enabled = IIf(Text1.SelLength > 0, True, False)
End Sub
Private Sub Text1_GotFocus()
muPaste.Enabled = Not (Clipboard.GetText = "")
End Sub

Private Sub Timer1_Timer()
Label1.Caption = Time
End Sub
