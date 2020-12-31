VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GeraLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private texto As String
Dim vDate, vTime As Date
Public Sub Registrar(vData As String)
vDate = Date
vTime = Time
       
      Open "logsis.txt" For Append As #1
      
      Print #1, vData & " - Data:" & vDate & " - Hora:" & vTime
    
    Close #1
End Sub
