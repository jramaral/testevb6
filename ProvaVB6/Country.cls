VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Country"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"LangCollection"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'Classe usada para pegar as informações de cada pais inclusive as linguages

Private msIsoCode As String
Private msName As String
Private msCapitalCity As String
Private mlanguages As Collection
Private mvarLangCollection As LangCollection

Private mvarCodigo As Integer

Private mvarPhoneCode As Integer
Private mvarContinentCode As String
Private mvarCurrencyIsoCode As String
Private mvarCountryFla As Variant
Public Property Let CountryFla(ByVal vData As Variant)

    mvarCountryFla = vData
End Property
Public Property Set CountryFla(ByVal vData As Variant)
    Set mvarCountryFla = vData
End Property
Public Property Get CountryFla() As Variant
    If IsObject(mvarCountryFla) Then
        Set CountryFla = mvarCountryFla
    Else
        CountryFla = mvarCountryFla
    End If
End Property
Public Property Let CurrencyIsoCode(ByVal vData As String)
    mvarCurrencyIsoCode = vData
End Property
Public Property Get CurrencyIsoCode() As String
    CurrencyIsoCode = mvarCurrencyIsoCode
End Property

Public Property Let ContinentCode(ByVal vData As String)
    mvarContinentCode = vData
End Property
Public Property Get ContinentCode() As String
    ContinentCode = mvarContinentCode
End Property
Public Property Let PhoneCode(ByVal vData As Integer)
    mvarPhoneCode = vData
End Property
Public Property Get PhoneCode() As Integer
    PhoneCode = mvarPhoneCode
End Property

Public Property Let Codigo(ByVal vData As Integer)
    mvarCodigo = vData
End Property
Public Property Get Codigo() As Integer
    Codigo = mvarCodigo
End Property
Public Property Let IsoCode(novoValor As String)
 msIsoCode = novoValor
End Property
Public Property Get IsoCode() As String
  IsoCode = msIsoCode
End Property
Public Property Let Name(novoValor As String)
 msName = novoValor
End Property
Public Property Get Name() As String
  Name = msName
End Property
Public Property Let CapitalCity(novoValor As String)
 msCapitalCity = novoValor
End Property
Public Property Get CapitalCity() As String
 CapitalCity = msCapitalCity
End Property
Public Property Let Languages(novoValor As Collection)
 mlanguages = novoValor
End Property
Public Property Get Languages() As Collection
 Languages = mlanguages
End Property

Public Property Get LangCollection() As LangCollection
    If mvarLangCollection Is Nothing Then
        Set mvarLangCollection = New LangCollection
    End If
    Set LangCollection = mvarLangCollection
End Property

Public Property Set LangCollection(vData As LangCollection)
    Set mvarLangCollection = vData
End Property
Private Sub Class_Terminate()
    Set mvarLangCollection = Nothing
End Sub





