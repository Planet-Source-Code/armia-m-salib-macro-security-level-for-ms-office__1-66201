VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Set"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
'kk = 1 for low security, 2 for meduim and 3 for high secuirity.
kk = 1
C = Chr$(kk)


For I = 0 To 50
'I change for diffrent version of MS Office
s = Str$(I)
s = Mid(s, 2) + ".0"
s = "Software\Microsoft\Office\" + s
B = bGetRegValue(HKEY_CURRENT_USER, s, "(Default)")
If B = 1 Then
'We found install version of Office

'To change security level to wanted for PowerPoint
g = s + "\PowerPoint\Security"
B = bSetRegValue(HKEY_CURRENT_USER, g, "Level", C)

'To change security level to wanted for Word
g = s + "\Word\Security"
B = bSetRegValue(HKEY_CURRENT_USER, g, "Level", C)
End If
Next I
End
End Sub
