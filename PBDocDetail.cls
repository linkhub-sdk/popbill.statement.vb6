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
Public spare6          As String
Public spare7          As String
Public spare8          As String
Public spare9          As String
Public spare10          As String
Public spare11          As String
Public spare12          As String
Public spare13          As String
Public spare14          As String
Public spare15          As String
Public spare16          As String
Public spare17          As String
Public spare18          As String
Public spare19          As String
Public spare20          As String




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
    toDictionary.Add "spare6", spare6
    toDictionary.Add "spare7", spare7
    toDictionary.Add "spare8", spare8
    toDictionary.Add "spare9", spare9
    toDictionary.Add "spare10", spare10
    toDictionary.Add "spare11", spare11
    toDictionary.Add "spare12", spare12
    toDictionary.Add "spare13", spare13
    toDictionary.Add "spare14", spare14
    toDictionary.Add "spare15", spare15
    toDictionary.Add "spare16", spare16
    toDictionary.Add "spare17", spare17
    toDictionary.Add "spare18", spare18
    toDictionary.Add "spare19", spare19
    toDictionary.Add "spare20", spare20
    
    
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
    spare6 = dic.Item("spare6")
    spare7 = dic.Item("spare7")
    spare8 = dic.Item("spare8")
    spare9 = dic.Item("spare9")
    spare10 = dic.Item("spare10")
    spare11 = dic.Item("spare11")
    spare12 = dic.Item("spare12")
    spare13 = dic.Item("spare13")
    spare14 = dic.Item("spare14")
    spare15 = dic.Item("spare15")
    spare16 = dic.Item("spare16")
    spare17 = dic.Item("spare17")
    spare18 = dic.Item("spare18")
    spare19 = dic.Item("spare19")
    spare20 = dic.Item("spare20")
    
End Sub
    
