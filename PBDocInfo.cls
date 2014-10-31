VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PBDocInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public itemCode                As Integer
Public itemKey                 As String
Public invoiceNum              As String
Public mgtKey                  As String

Public stateCode               As Integer
Public taxType                 As String
Public purposeType             As String

Public writeDate               As String

Public senderCorpName        As String
Public senderCorpNum         As String
Public receiverCorpName        As String
Public receiverCorpNum         As String

Public supplyCostTotal         As String
Public taxTotal                As String

Public issueDT                 As String
Public stateDT                 As String
Public openYN                  As Boolean
Public openDT                  As String

Public stateMemo               As String

Public regDT                   As String

Public Sub fromDictionary(dic As Variant)

    itemCode = dic.Item("itemCode")
    itemKey = dic.Item("itemKey")
    invoiceNum = dic.Item("invoiceNum")
    mgtKey = dic.Item("mgtKey")
    
    stateCode = dic.Item("stateCode")
    taxType = dic.Item("taxType")
    purposeType = dic.Item("purposeType")
     
    writeDate = dic.Item("writeDate")
    
    senderCorpName = dic.Item("senderCorpName")
    senderCorpNum = dic.Item("senderCorpNum")
    receiverCorpName = dic.Item("receiverCorpName")
    receiverCorpNum = dic.Item("receiverCorpNum")
    
    supplyCostTotal = dic.Item("supplyCostTotal")
    taxTotal = dic.Item("taxTotal")
    
    issueDT = dic.Item("issueDT")
    stateDT = dic.Item("stateDT")
    openYN = dic.Item("openYN")
    openDT = dic.Item("openDT")
    
    stateMemo = dic.Item("stateMemo")
    
    regDT = dic.Item("regDT")
    

End Sub

