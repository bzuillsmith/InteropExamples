VERSION 5.00
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
    Me.AppTitle.Caption = CStr(oCalc.Add(1, 2))
End Sub

Private Sub Form_Load()
 
End Sub

