VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form frmExplain 
   Caption         =   "Short Explanation Links to Fibonacci and Why Golden Mean is so Amazing"
   ClientHeight    =   9705
   ClientLeft      =   2130
   ClientTop       =   675
   ClientWidth     =   10860
   Icon            =   "frmExplain.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   9705
   ScaleWidth      =   10860
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   9255
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   10455
      ExtentX         =   18441
      ExtentY         =   16325
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
End
Attribute VB_Name = "frmExplain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
frmExplain.WebBrowser1.Navigate (App.Path & "\fib.html")

End Sub
