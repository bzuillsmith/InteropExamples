VERSION 5.00
Object = "{CC7518BB-34A4-4FA4-B80E-D195059DD36E}#1.0#0"; "mscoree.dll"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5985
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9090
   LinkTopic       =   "Form1"
   ScaleHeight     =   5985
   ScaleWidth      =   9090
   StartUpPosition =   3  'Windows Default
   Begin InteropExampleCtl.UserControl1 TheControl 
      Height          =   3135
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   8775
      Object.Visible         =   "True"
      Enabled         =   "True"
      ForegroundColor =   "-2147483630"
      BackgroundColor =   "-2147483633"
      BackColor       =   "Control"
      ForeColor       =   "ControlText"
      Location        =   "8, 48"
      Name            =   "UserControl1"
      Size            =   "585, 209"
      Object.TabIndex        =   "0"
   End
   Begin VB.Label AppTitle 
      Caption         =   "Interop Example"
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
    Dim oCalc As Calculator
    Set oCalc = New Calculator
    Me.AppTitle.Caption = CStr(oCalc.Add(1, 2))
End Sub

Private Sub Form_Load()
    Randomize
End Sub

Private Sub TheControl_Click()
    Dim oCalc As Calculator
    Set oCalc = New Calculator
    Me.AppTitle.Caption = CStr(oCalc.Add(Rnd * 100, 4))
End Sub
