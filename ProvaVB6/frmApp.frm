VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form frmApp 
   Caption         =   "Teste"
   ClientHeight    =   8250
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   12840
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   8250
   ScaleWidth      =   12840
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnFechar 
      Caption         =   "Fechar"
      Height          =   810
      Left            =   9675
      TabIndex        =   3
      Top             =   240
      Width           =   2940
   End
   Begin VB.CommandButton btnSalvarDados 
      Caption         =   "Salvar Dados"
      Height          =   810
      Left            =   3285
      TabIndex        =   2
      Top             =   240
      Width           =   2940
   End
   Begin VB.TextBox txtDadosRetornados 
      Height          =   7005
      Left            =   135
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   1
      Top             =   1140
      Width           =   12435
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   9945
      Top             =   7470
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.CommandButton btnBaixarDados 
      Caption         =   "Baixar Dados"
      Height          =   810
      Left            =   135
      TabIndex        =   0
      Top             =   240
      Width           =   2940
   End
End
Attribute VB_Name = "frmApp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim clang As New LanguageCountry
Dim contryes As New Collection
Dim lang As New Collection
Dim pais As New Country
Dim fields As New Collection

Dim indiceLang As Integer
Dim vcodigo As Integer

Dim enderecoUrl As String
Dim textParse As String
Dim arquivoXml As DOMDocument60
Dim log As New GeraLog


Private Sub btnBaixarDados_Click()

Dim posicao As Integer
Dim str1 As String

On Error GoTo Click_Error

Screen.MousePointer = vbHourglass

Set arquivoXml = New DOMDocument60

enderecoUrl = "http://webservices.oorsprong.org/websamples.countryinfo/CountryInfoService.wso/FullCountryInfoAllCountries"

With Inet1
  .AccessType = icDirect
  .Proxy = ""
  .Protocol = icHTTP
  textParse = .OpenURL(enderecoUrl)
End With



arquivoXml.loadXML textParse

For Each obj In arquivoXml.documentElement.childNodes
   
    str1 = obj.childNodes(0).Text
    posicao = InStr(1, str1, "A", 1)
    If posicao = 0 Or posicao > 1 Then
        arquivoXml.documentElement.removeChild obj
    End If

Next

txtDadosRetornados.Text = arquivoXml.xml

Screen.MousePointer = vbDefault
log.Registrar "Xml Lido com sucesso!"
On Error GoTo 0
   Exit Sub

Click_Error:
 'Algo deu errado
 log.Registrar "Erro " & Err.Description
 Screen.MousePointer = vbDefault
 MsgBox "Error " & Err.Number & " (" & Err.Description & ")"


End Sub

Private Sub btnFechar_Click()
log.Registrar "Encerrou o sistmea"
Unload Me
End Sub

Private Sub btnSalvarDados_Click()

On Error GoTo Save_Error

Screen.MousePointer = vbHourglass

vcodigo = 1

If arquivoXml Is Nothing Then
   MsgBox "Não existe informações para serem gravadas", vbOKOnly = vbExclamation, "Atenção"
   Screen.MousePointer = vbDefault
   Exit Sub
End If


For Each obj In arquivoXml.documentElement.childNodes
    indiceLang = 0
    pais.Codigo = vcodigo
    pais.IsoCode = obj.childNodes(0).Text
    pais.Name = obj.childNodes(1).Text
    pais.CapitalCity = obj.childNodes(2).Text
    pais.PhoneCode = obj.childNodes(3).Text
    pais.ContinentCode = obj.childNodes(4).Text
    pais.CurrencyIsoCode = obj.childNodes(5).Text
    pais.CountryFla = obj.childNodes(6).Text
    
    Dim languageCount As Integer
    languageCount = obj.childNodes(7).childNodes.length
    
    While indiceLang <= languageCount - 1
        clang.CodigoContryInfo = vcodigo
        clang.LangIsoCode = obj.childNodes(7).childNodes(indiceLang).childNodes(0).Text
        clang.LangName = obj.childNodes(7).childNodes(indiceLang).childNodes(1).Text
        pais.LangCollection.Add clang.LangName, clang.LangIsoCode, clang.CodigoContryInfo
        indiceLang = indiceLang + 1
    Wend

   contryes.Add pais, pais.IsoCode
   vcodigo = vcodigo + 1
Set pais = Nothing
Set clang = Nothing

Next


 Dim conecta As New ConectaDatabase
 conecta.GravarDados contryes
 
 Screen.MousePointer = vbDefault
 log.Registrar "Dados Gravados "
 On Error GoTo 0
   Exit Sub

Save_Error:
 'Algo deu errado
  log.Registrar "Erro " & Err.Description
  Screen.MousePointer = vbDefault
  MsgBox "Error " & Err.Number & " (" & Err.Description & ")"
End Sub

