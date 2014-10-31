VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PBDocDetail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public serialNum       As Integer
Public purchaseDT      As String
Public itemName        As String
Public spec            As String
Public unit            As String
Public qty             As String
Public unitCost        As String
Public supplyCost      As String
Public tax             As String
Public remark          As String
Public spare1          As String
Public spare2          As String
Public spare3          As String
Public spare4          As String
Public spare5          As String



Public Function toDictionary() As Variant
    Set toDictionary = CreateObject("Scripting.Dictionary")
    
    toDictionary.Add "serialNum", serialNum
    toDictionary.Add "purchaseDT", purchaseDT
    toDictionary.Add "itemName", itemName
    toDictionary.Add "spec", spec
    toDictionary.Add "unit", unit
    toDictionary.Add "qty", qty
    toDictionary.Add "unitCost", unitCost
    toDictionary.Add "supplyCost", supplyCost
    toDictionary.Add "tax", tax
    toDictionary.Add "remark", remark
    toDictionary.Add "spare1", spare1
    toDictionary.Add "spare2", spare2
    toDictionary.Add "spare3", spare3
    toDictionary.Add "spare4", spare4
    toDictionary.Add "spare5", spare5
    
    
End Function

Public Sub fromDictionary(dic As Variant)

    serialNum = dic.Item("serialNum")
    purchaseDT = dic.Item("purchaseDT")
    itemName = dic.Item("itemName")
    spec = dic.Item("spec")
    unit = dic.Item("unit")
    qty = dic.Item("qty")
    unitCost = dic.Item("unitCost")
    supplyCost = dic.Item("supplyCost")
    tax = dic.Item("tax")
    remark = dic.Item("remark")
    spare1 = dic.Item("spare1")
    spare2 = dic.Item("spare2")
    spare3 = dic.Item("spare3")
    spare4 = dic.Item("spare4")
    spare5 = dic.Item("spare5")
    
    
End Sub
    
