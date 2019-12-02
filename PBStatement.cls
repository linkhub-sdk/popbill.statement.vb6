VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PBStatement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public sendNum              As String
Public receiveNum           As String
Public memo                 As String
Public emailSubject         As String

Public itemCode             As Integer
Public mgtKey               As String
Public invoiceNum           As String
Public formCode             As String
Public writeDate            As String
Public taxType              As String

Public senderCorpNum      As String
Public senderTaxRegID     As String
Public senderCorpName     As String
Public senderCEOName      As String
Public senderAddr         As String
Public senderBizClass     As String
Public senderBizType      As String
Public senderContactName  As String
Public senderDeptName     As String
Public senderTEL          As String
Public senderHP           As String
Public senderEmail        As String
Public senderFAX           As String

Public receiverCorpNum      As String
Public receiverTaxRegID     As String
Public receiverCorpName     As String
Public receiverCEOName      As String
Public receiverAddr         As String
Public receiverBizClass     As String
Public receiverBizType      As String
Public receiverContactName  As String
Public receiverDeptName     As String
Public receiverTEL          As String
Public receiverHP           As String
Public receiverEmail        As String
Public receiverFAX          As String

Public taxTotal             As String
Public supplyCostTotal      As String
Public totalAmount          As String
Public purposeType          As String
Public serialNum            As String
Public remark1              As String
Public remark2              As String
Public remark3              As String
Public businessLicenseYN    As Boolean
Public bankBookYN           As Boolean
Public faxsendYN            As Boolean
Public smssendYN            As Boolean
Public autoacceptYN         As Boolean


Public detailList As Collection
Public propertyBag As Dictionary


Public Function toDictionary() As Variant
    Set toDictionary = CreateObject("Scripting.Dictionary")
    
    toDictionary.Add "sendNum", sendNum
    toDictionary.Add "receiveNum", receiveNum
    toDictionary.Add "memo", memo
    toDictionary.Add "emailSubject", emailSubject
    
    toDictionary.Add "itemCode", itemCode
    toDictionary.Add "mgtKey", mgtKey
    toDictionary.Add "invoiceNum", invoiceNum
    toDictionary.Add "formCode", formCode
    
    toDictionary.Add "writeDate", writeDate
    
    toDictionary.Add "taxType", taxType
    
    toDictionary.Add "senderCorpNum", senderCorpNum
    toDictionary.Add "senderTaxRegID", senderTaxRegID
    toDictionary.Add "senderCorpName", senderCorpName
    toDictionary.Add "senderCEOName", senderCEOName
    toDictionary.Add "senderAddr", senderAddr
    toDictionary.Add "senderBizClass", senderBizClass
    toDictionary.Add "senderBizType", senderBizType
    toDictionary.Add "senderContactName", senderContactName
    toDictionary.Add "senderDeptName", senderDeptName
    toDictionary.Add "senderTEL", senderTEL
    toDictionary.Add "senderHP", senderHP
    toDictionary.Add "senderEmail", senderEmail
    toDictionary.Add "senderFAX", senderFAX
    
    toDictionary.Add "receiverCorpNum", receiverCorpNum
    toDictionary.Add "receiverTaxRegID", receiverTaxRegID
    toDictionary.Add "receiverCorpName", receiverCorpName
    toDictionary.Add "receiverCEOName", receiverCEOName
    toDictionary.Add "receiverAddr", receiverAddr
    toDictionary.Add "receiverBizType", receiverBizType
    toDictionary.Add "receiverBizClass", receiverBizClass
    toDictionary.Add "receiverContactName", receiverContactName
    toDictionary.Add "receiverDeptName", receiverDeptName
    toDictionary.Add "receiverTEL", receiverTEL
    toDictionary.Add "receiverHP", receiverHP
    toDictionary.Add "receiverEmail", receiverEmail
    toDictionary.Add "receiverFax", receiverFAX
    
    toDictionary.Add "taxTotal", taxTotal
    toDictionary.Add "supplyCostTotal", supplyCostTotal
    toDictionary.Add "totalAmount", totalAmount
    toDictionary.Add "purposeType", purposeType
    toDictionary.Add "serialNum", serialNum
    
    toDictionary.Add "businessLicenseYN", businessLicenseYN
    toDictionary.Add "bankBookYN", bankBookYN
    toDictionary.Add "faxsendYN", faxsendYN
    toDictionary.Add "smssendYN", smssendYN
    toDictionary.Add "autoAcceptYN", autoacceptYN
    
    toDictionary.Add "remark1", remark1
    toDictionary.Add "remark2", remark2
    toDictionary.Add "remark3", remark3
    
    If (detailList Is Nothing) = False Then
        Dim t_detailCollection As New Collection
        Dim row As PBDocDetail
        For Each row In detailList
            t_detailCollection.Add row.toDictionary()
        Next
        If t_detailCollection.Count > 0 Then
            toDictionary.Add "detailList", t_detailCollection
        End If
    End If
   
   If Not (propertyBag Is Nothing) Then
        toDictionary.Add "propertyBag", propertyBag
    End If
    
End Function

Public Sub fromDictionary(dic As Variant)

    itemCode = dic.Item("itemCode")
    mgtKey = dic.Item("mgtKey")
    invoiceNum = dic.Item("invoiceNum")
    formCode = dic.Item("formCode")
      
    writeDate = dic.Item("writeDate")
    taxType = dic.Item("taxType")
    
    senderCorpNum = dic.Item("senderCorpNum")
    senderTaxRegID = dic.Item("senderTaxRegID")
    senderCorpName = dic.Item("senderCorpName")
    senderCEOName = dic.Item("senderCEOName")
    senderAddr = dic.Item("senderAddr")
    senderBizClass = dic.Item("senderBizClass")
    senderBizType = dic.Item("senderBizType")
    senderContactName = dic.Item("senderContactName")
    senderDeptName = dic.Item("senderDeptName")
    senderTEL = dic.Item("senderTEL")
    senderHP = dic.Item("senderHP")
    senderEmail = dic.Item("senderEmail")
    senderFAX = dic.Item("senderFAX")
    
    receiverCorpNum = dic.Item("receiverCorpNum")
    receiverTaxRegID = dic.Item("receiverTaxRegID")
    receiverCorpName = dic.Item("receiverCorpName")
    receiverCEOName = dic.Item("receiverCEOName")
    receiverAddr = dic.Item("receiverAddr")
    receiverBizClass = dic.Item("receiverBizClass")
    receiverBizType = dic.Item("receiverBizType")
    receiverContactName = dic.Item("receiverContactName")
    receiverDeptName = dic.Item("receiverDeptName")
    receiverTEL = dic.Item("receiverTEL")
    receiverHP = dic.Item("receiverHP")
    receiverEmail = dic.Item("receiverEmail")
    
    taxTotal = dic.Item("taxTotal")
    supplyCostTotal = dic.Item("supplyCostTotal")
    totalAmount = dic.Item("totalAmount")
    purposeType = dic.Item("purposeType")
    serialNum = dic.Item("serialNum")
    remark1 = dic.Item("remark1")
    remark2 = dic.Item("remark2")
    remark3 = dic.Item("remark3")
    
    businessLicenseYN = dic.Item("businessLicenseYN")
    bankBookYN = dic.Item("bankBookYN")
    faxsendYN = dic.Item("faxsendYN")
    smssendYN = dic.Item("smssendYN")
    autoacceptYN = dic.Item("autoAcceptYN")
 
    remark1 = dic.Item("remark1")
    remark2 = dic.Item("remark2")
    remark3 = dic.Item("remark3")
    
    Dim tCollection As Collection
    
    Set tCollection = dic.Item("detailList")
    
    If (tCollection Is Nothing) = False Then
        Set detailList = New Collection
        
        Dim T As Variant
        Dim detail As PBDocDetail
            
            
        For Each T In tCollection
            Set detail = New PBDocDetail
            
            detail.fromDictionary T
            
            detailList.Add detail
            
        Next
    
    End If
    
    Set propertyBag = dic.Item("propertyBag")
  
End Sub
