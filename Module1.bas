Attribute VB_Name = "Module1"
Sub QuickBooks()

 'create excel application
  Dim oXLApp  As Excel.Application
 'create excel workbook
  Dim oXLWBook As Excel.Workbook
 'create sheet within that workbook
  Dim oXLWSheet As Excel.Worksheet
 'get value from excel cell
  Dim target As String
  
  
  'create object
   Set App = New Excel.Application
   'set workbook path, this is vb  application plus folder with excel spreadsheet
   workBookPath = "C:\Documents and Settings\daj.LOT\Desktop\Proyecto de practica QuickBooks.xls"
   'set workbook and worksheet
   Set WBook = App.Workbooks.Add(workBookPath)
   Set WSheet = WBook.ActiveSheet
   'get value from particular cell
   target = ActiveSheet.Range("A1")
   ActiveSheet.Range("A1").Select
  
 ' Creating the conection with QuickBooks
   
   Dim xmlCom  As QBXMLRP2Lib.RequestProcessor2
   Set xmlCom = New QBXMLRP2Lib.RequestProcessor2
   xmlCom.OpenConnection "", "IntQB"
   
   Dim pass As String
   Dim sessionBegun As Boolean
   pass = xmlCom.BeginSession("", qbFileOpenDoNotCare)
   sessionBegan = True
   
   'Check which version of qbXML to use
    Dim supportedVersion As String
    'supportedVersion = qbXMLLatestVersion(xmlCom, pass)

   'Check which version of qbXML to use
    supportedVersion = qbXMLLatestVersion(xmlCom, pass)
  
    Dim addr4supported As Boolean
    addr4supported = False
    
    Dim requestMsgSet As IMsgSetRequest
    If (Val(supportedVersion) >= 2) Then
        qbXMLVersionSpec = "<?qbxml version=""" & supportedVersion & """?>"
        addr4supported = True
    ElseIf (supportedVersion = "1.1") Then
        qbXMLVersionSpec = "<!DOCTYPE QBXML PUBLIC '-//INTUIT//DTD QBXML QBD " _
                           & supportedVersion & "//EN' 'http://developer.intuit.com'>"
    Else
        MsgBox "You are apparently running QuickBooks 2002 Release 1, we strongly recommend that you use QuickBooks' online update feature to obtain the latest fixes and enhancements", vbExclamation
        qbXMLVersionSpec = "<!DOCTYPE QBXML PUBLIC '-//INTUIT//DTD QBXML QBD " _
                           & supportedVersion & "//EN' 'http://developer.intuit.com'>"
    End If
    
    'Stating creating XML request
   Dim invoiceRq As MSXML2.DOMDocument
   Set invoiceRq = CreateObject("MSXML2.DOMDocument")
   
   Dim root As IXMLDOMNode
   Set root = invoiceRq.createElement("QBXML")
   invoiceRq.appendChild root
   
   Dim msgNode As IXMLDOMNode
   Set msgNode = invoiceRq.createElement("msgNode")
   root.appendChild msgNode
   
   Dim invoiceInfoNode As IXMLDOMNode
   Set invoiceInfoNode = invoiceRq.createElement("invoiceInfoNode")
   msgNode.appendChild invoiceInfoNode
   
   'Setting Atributes of DOM Document
   Dim idAttr As IXMLDOMAttribute
   Set idAttr = invoiceRq.createAttribute("RequestID")
   idAttr.Text = "0"
   msgNode.Attributes.setNamedItem idAttr
   
   
   'Creating tags for each data
   
   Dim customer As IXMLDOMElement
   Set customer = invoiceRq.createElement("CustomerCode")
   customer.Text = target
   invoiceInfoNode.appendChild customer
   
   
   
   Dim invoiceDate As IXMLDOMElement
   Set invoiceDate = invoiceRq.createElement("Date")
   invoiceDate.Text = ActiveCell.Offset(0, 1)
   ActiveCell.Offset(0, 1).Select
   invoiceInfoNode.appendChild invoiceDate
   
   Dim invoiceNumber As IXMLDOMElement
   Set invoiceNumber = invoiceRq.createElement("Number")
   invoiceNumber.Text = ActiveCell.Offset(0, 1)
   ActiveCell.Offset(0, 1).Select
   invoiceInfoNode.appendChild invoiceNumber
   
   Dim PO As IXMLDOMElement
   Set PO = invoiceRq.createElement("PO")
   PO.Text = ActiveCell.Offset(0, 1)
   ActiveCell.Offset(0, 1).Select
   invoiceInfoNode.appendChild PO
   
   Dim rep As IXMLDOMElement
   Set rep = invoiceRq.createElement("Rep")
   rep.Text = ActiveCell.Offset(0, 1)
   ActiveCell.Offset(0, 1).Select
   invoiceInfoNode.appendChild rep
   
   Dim value As IXMLDOMElement
   Set value = invoiceRq.createElement("Value")
   value.Text = ActiveCell.Offset(0, 1)
   ActiveCell.Offset(0, 1).Select
   invoiceInfoNode.appendChild value


Dim strXMLRequest As String
Dim strXMLResponse As String

strXMLRequest = "<?xml version=""1.0""?>" & _
                qbXMLVersionSpec & root.xml
                
strXMLResponse = xmlCom.ProcessRequest(pass, strXMLRequest)

MsgBox strXMLRequest, vbOKOnly, "RequestXML"

End Sub


Function qbXMLLatestVersion(rp As RequestProcessor2, ticket As String) As String
    Dim strXMLVersions() As String
    Dim xml As New DOMDocument

    'Create the QBXML aggregate
    Dim rootElement As IXMLDOMNode
    Set rootElement = xml.createElement("QBXML")
    xml.appendChild rootElement
  
    'Add the QBXMLMsgsRq aggregate to the QBXML aggregate
    Dim QBXMLMsgsRqNode As IXMLDOMNode
    Set QBXMLMsgsRqNode = xml.createElement("QBXMLMsgsRq")
    rootElement.appendChild QBXMLMsgsRqNode


    'Set the QBXMLMsgsRq onError attribute to continueOnError
    Dim onErrorAttr As IXMLDOMAttribute
    Set onErrorAttr = xml.createAttribute("onError")
    onErrorAttr.Text = "stopOnError"
    QBXMLMsgsRqNode.Attributes.setNamedItem onErrorAttr
  
    'Add the InvoiceAddRq aggregate to QBXMLMsgsRq aggregate
    Dim HostQuery As IXMLDOMNode
    Set HostQuery = xml.createElement("HostQueryRq")
    QBXMLMsgsRqNode.appendChild HostQuery
    
    strXMLRequest = _
        "<?xml version=""1.0"" ?>" & _
        "<!DOCTYPE QBXML PUBLIC '-//INTUIT//DTD QBXML QBD 1.0//EN' 'http://developer.intuit.com'>" _
        & rootElement.xml

    strXMLResponse = rp.ProcessRequest(ticket, strXMLRequest)
    Dim QueryResponse As New DOMDocument

    'Parse the response XML
    QueryResponse.async = False
    QueryResponse.loadXML (strXMLResponse)

    Dim supportedVersions As IXMLDOMNodeList
    Set supportedVersions = QueryResponse.getElementsByTagName("SupportedQBXMLVersion")
    
    Dim VersNode As IXMLDOMNode
    
    Dim i As Long
    Dim vers As Double
    Dim LastVers As Double
    LastVers = 0
    For i = 0 To supportedVersions.Length - 1
        Set VersNode = supportedVersions.Item(i)
        vers = VersNode.FirstChild.Text
        If (vers > LastVers) Then
            LastVers = vers
            qbXMLLatestVersion = VersNode.FirstChild.Text
        End If
    Next i
End Function


