VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Conexao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'variável localpara manter o  valor da propriedade
Private mvarHostName As String
Private mvarPort As Integer
Private mvarUserName As String
Private mvarPwd As String
Public Property Let Pwd(ByVal vData As String)
    mvarPwd = vData
End Property

Public Property Get Pwd() As String
    Pwd = mvarPwd
End Property

Public Property Let UserName(ByVal vData As String)
    mvarUserName = vData
End Property


Public Property Get UserName() As String
    UserName = mvarUserName
End Property

Public Property Let Port(ByVal vData As Integer)
    mvarPort = vData
End Property

Public Property Get Port() As Integer
    Port = mvarPort
End Property

Public Property Let HostName(ByVal vData As String)
    mvarHostName = vData
End Property

Public Property Get HostName() As String
    HostName = mvarHostName
End Property



