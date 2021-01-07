VERSION 5.00
Object = "{CC7518BB-34A4-4FA4-B80E-D195059DD36E}#1.0#0"; "InteropExample.tlb"
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
   Begin VB.CommandButton Command 
      Caption         =   "Open a Form"
      Height          =   255
      Left            =   3720
      TabIndex        =   3
      Top             =   240
      Width           =   2175
   End
   Begin VB.TextBox Text 
      Height          =   285
      Left            =   1920
      TabIndex        =   2
      Text            =   "Text"
      Top             =   240
      Width           =   1335
   End
   Begin InteropExampleCtl.UserControl1 TheControl 
      Height          =   3135
      Left            =   120
      TabIndex        =   1
      Top             =   1440
      Width           =   8775
      Object.Visible         =   "True"
      Enabled         =   "True"
      ForegroundColor =   "-2147483630"
      BackgroundColor =   "-2147483633"
      BackColor       =   "Control"
      ForeColor       =   "ControlText"
      Location        =   "8, 96"
      Name            =   "UserControl1"
      Size            =   "585, 209"
      Object.TabIndex        =   "0"
   End
   Begin VB.Label AppTitle 
      Caption         =   ".NET Control is below"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   1080
      Width           =   2295
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command_Click()
    Form2.Show vbModal
End Sub
