VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LanguageCountry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "LangNome"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'Classe usada para pegar as informações das linguagens de cada pais

Private msIsoCode As String
Private msName As String
Private mvarCodigoLang As Integer 'local copy
Private mvarCodigoContryInfo As Integer 'local copy
Public Property Let CodigoContryInfo(ByVal vData As Integer)
    mvarCodigoContryInfo = vData
End Property

Public Property Get CodigoContryInfo() As Integer
    CodigoContryInfo = mvarCodigoContryInfo
End Property

Public Property Let CodigoLang(ByVal vData As Integer)
    mvarCodigoLang = vData
End Property

Public Property Get CodigoLang() As Integer
    CodigoLang = mvarCodigoLang
End Property

Public Property Let LangIsoCode(novoValor As String)
 msIsoCode = novoValor
End Property
Public Property Get LangIsoCode() As String
  LangIsoCode = msIsoCode
End Property
Public Property Let LangName(novoValor As String)
 msName = novoValor
End Property
Public Property Get LangName() As String
  LangName = msName
End Property
