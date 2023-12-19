VERSION 5.00
Begin VB.Form report 
   Caption         =   "report"
   ClientHeight    =   7125
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14040
   LinkTopic       =   "Form1"
   ScaleHeight     =   7125
   ScaleWidth      =   14040
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   6480
      TabIndex        =   0
      Top             =   3360
      Width           =   1215
   End
End
Attribute VB_Name = "report"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
rent.DataGrid1.DataSource = "rent.adodc1"
rent.DataGrid1.Visible



End Sub
