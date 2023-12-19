VERSION 5.00
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "Menu"
   ClientHeight    =   9525
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   17265
   LinkTopic       =   "MDIForm1"
   Picture         =   "MDIForm1.frx":0000
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuEdit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuNew 
      Caption         =   "New"
      Begin VB.Menu mnuAuto 
         Caption         =   "Auto"
      End
      Begin VB.Menu mnuRent 
         Caption         =   "Rent"
      End
      Begin VB.Menu mnuDriver 
         Caption         =   "Driver"
      End
      Begin VB.Menu mnuStudent 
         Caption         =   "Student"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "View"
      Begin VB.Menu mnuStRepo 
         Caption         =   "Student Report"
      End
      Begin VB.Menu mnuRepo 
         Caption         =   "Reports"
      End
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub mnuARepo_Click()
ReportA.Show

End Sub

Private Sub mnuAuto_Click()
auto.Adodc1.Recordset.MoveLast
auto.Show

End Sub

Private Sub mnuDriver_Click()
driver.Adodc1.Recordset.MoveLast
driver.Show


End Sub

Private Sub mnuDrRepo_Click()
ReportD.Show

End Sub

Private Sub mnuEdit_Click()
Unload Me

End Sub

Private Sub mnuRent_Click()

rent.Adodc1.Recordset.MoveLast
rent.Show

End Sub

Private Sub mnuReRepo_Click()

End Sub

Private Sub mnuRepo_Click()
ReportA.Show
End Sub

Private Sub mnuStRepo_Click()
SReport.Show


End Sub

Private Sub mnuStudent_Click()
student.Adodc1.Recordset.MoveLast

student.Show

End Sub
