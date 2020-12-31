Attribute VB_Name = "AbrirBase"

Rem API usado para ler arquivo INI
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

Rem API usada para escrever em um arquivo INI.
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

Public Sub LerIni(sServer As String, sDb As String)
sServer = ReadINI("CONEXAO", "SERVER", App.Path & "\CONFIG.INI")
sDb = ReadINI("CONEXAO", "DB", App.Path & "\CONFIG.INI")
End Sub
Public Function ReadINI(Secao As String, Entrada As String, Arquivo As String)
Dim Retlen As String
Dim ret As String
ret = String$(255, 0)
Retlen = GetPrivateProfileString(Secao, Entrada, "", ret, Len(ret), Arquivo)
ret = Left$(ret, Retlen)
ReadINI = ret
End Function
Public Sub WriteINI(Secao As String, Entrada As String, texto As String, Arquivo As String)
WritePrivateProfileString Secao, Entrada, texto, Arquivo
End Sub
