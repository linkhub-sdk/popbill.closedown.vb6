VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PBCorpState"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public CorpNum As String
Public ctype As String
Public state As String
Public stateDate As String
Public checkDate As String
Public typeDate As String

Public Sub fromDictionary(dic As Variant)
    If IsNull(dic.Item("corpNum")) = False Then
        CorpNum = dic.Item("corpNum")
    Else
        CorpNum = ""
    End If
    
    If IsNull(dic.Item("type")) = False Then
        ctype = dic.Item("type")
    Else
        ctype = ""
    End If
    
    If IsNull(dic.Item("state")) = False Then
        state = dic.Item("state")
    Else
        state = ""
    End If
    
    If IsNull(dic.Item("stateDate")) = False Then
        stateDate = dic.Item("stateDate")
    Else
        stateDate = ""
    End If
    
    If IsNull(dic.Item("checkDate")) = False Then
        checkDate = dic.Item("checkDate")
    Else
        checkDate = ""
    End If
    
    If IsNull(dic.Item("typeDate")) = False Then
        typeDate = dic.Item("typeDate")
    Else
        typeDate = ""
    End If
End Sub
