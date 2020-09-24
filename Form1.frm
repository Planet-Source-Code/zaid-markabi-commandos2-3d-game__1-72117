VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Commandos II - Setup"
   ClientHeight    =   2610
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   4710
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2610
   ScaleWidth      =   4710
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Click Here to Get [ COMMANDOS.rar ]"
      Height          =   495
      Left            =   360
      TabIndex        =   0
      Top             =   1680
      Width           =   3855
   End
   Begin VB.Label Label2 
      Caption         =   "so.. i put all files ( with DLL ) in single rar file, and you have to unzip it using [ WinRar ] ."
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   4335
   End
   Begin VB.Label Label1 
      Caption         =   "There are many DLL files, but PlanetSourceCode will remove all of them ."
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   4335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error Resume Next

FileCopy App.Path + "\COMMANDOS II.files here", App.Path + "\COMMANDOS.rar"

MsgBox " See [ COMMANDOS.rar ] and unzip it using Winrar" + vbCrLf + vbCrLf + "then you will see all files , ActiveX files included !"

Shell App.Path + "\COMMANDOS.rar"

End
End Sub
