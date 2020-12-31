VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ConectaDatabase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"Conexao"
Const Provedor = "driver={SQL SERVER}"

Public Conexao As New ADODB.Connection
Public Command As New ADODB.Command
Public ComandoLang As New ADODB.Command
Public Consulta As New ADODB.Recordset
Private Query As String
Private QueryLang As String
Private ServerApl As String
Private DBApl As String
Private RegAffected As Long
Dim log As New GeraLog


Public Sub Conectar()
On Error GoTo Serve_Error

    LerIni ServerApl, DBApl
    log.Registrar "Arquivo .INI Lido"
       

    Conexao.ConnectionString = Provedor & ";server=" & ServerApl & ";database=" & DBApl
    
    Conexao.Open
    log.Registrar "Banco de dados acessado"
    
   On Error GoTo 0
   Exit Sub
   
Serve_Error:
    log.Registrar "Erro ao acessar o banco de dados " & Err.Description
    MsgBox "Error " & Err.Number & " (" & Err.Description & ")"
   
    
End Sub

'Grava as informações quando o botão GravarDados é clicado
Public Sub GravarDados(vData As Collection)
  On Error GoTo Data_Error
 
    Conectar
    
    'Inicia a transacao
    Conexao.BeginTrans
    
    Set Command.ActiveConnection = Conexao
    Set ComandoLang.ActiveConnection = Conexao
      
    For Each ite In vData
    Query = "Insert Into CountryInfo (Cod, sISOCode,  sName, sCapitalCity, sPhoneCode, sContinentCode, sCurrencyISOCode, sCountryFlag) values ('" & ite.Codigo & "','" & ite.IsoCode & "','" & Replace(ite.Name, "'", "''") & "','" & Replace(ite.CapitalCity, "'", "''") & "','" & ite.PhoneCode & "','" & ite.ContinentCode & "','" & ite.CurrencyIsoCode & "','" & ite.CountryFla & "')"
    Conexao.Execute Query, RegAffected, adExecuteNoRecords
        For Each lang In ite.LangCollection
            QueryLang = "Insert Into Languages (IsoCode,  sNameLang, Cod_Country) values ('" & lang.LangIsoCode & "','" & lang.LangName & "','" & lang.CodigoContryInfo & "')"
           Conexao.Execute QueryLang, RegAffected, adExecuteNoRecords
        Next
    Next
 
   'comita os dados da transacao
   Conexao.CommitTrans
   
   MsgBox "Dados Gravados com sucesso!", vbOK = vbInformation, "Mensagem"
   log.Registrar "Dados gravados no banco de dados"
   
On Error GoTo 0
   Exit Sub
Data_Error:
 'Algo deu errado
 log.Registrar "Erro " & Err.Description
 
 'Descarta as alterações
 Conexao.RollbackTrans
  MsgBox "Error " & Err.Number & " (" & Err.Description & ")"
End Sub

