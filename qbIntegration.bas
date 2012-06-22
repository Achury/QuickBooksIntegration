'Code by Daniel Achury (DAJ)

Public Sub QBFC_AddInvoice()
    
    'Declare various utility variables
    Dim booSessionBegun As Boolean
    booSessionBegun = False
  
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
     workBookPath = "Y:\Registers CF\Resumen Facturas.xls"
     'set workbook and worksheet
     Set WBook = App.Workbooks.Add(workBookPath)
     Set WSheet = WBook.ActiveSheet
     'get value from particular cell
     target = ActiveSheet.Range("A3")
     ActiveSheet.Range("A3").Select
     
     
     On Error GoTo ErrHandler

    'We want to know if we've begun a session so we can end it if an
    'error sends us to the exception handler
    booSessionBegun = False

    'Create the session manager object using QBFC, and use this
    'object to open a connection and begin a session with QuickBooks.
    Dim SessionManager As New QBSessionManager
    SessionManager.OpenConnection "", "IDN Add Invoice Sample"
    SessionManager.BeginSession "", omDontCare
    booSessionBegun = True
    
    Dim supportedVersion As String
    'supportedVersion = QBFCLatestVersion(SessionManager)
  
    Dim addr4supported As Boolean
    addr4supported = False
    ' Create the message set request object
    Dim requestMsgSet As IMsgSetRequest
    If (supportedVersion >= "6.0") Then
        Set requestMsgSet = SessionManager.CreateMsgSetRequest("US", 6, 0)
        addr4supported = True
    ElseIf (supportedVersion >= "5.0") Then
        Set requestMsgSet = SessionManager.CreateMsgSetRequest("US", 5, 0)
        addr4supported = True
    ElseIf (supportedVersion >= "4.0") Then
        Set requestMsgSet = SessionManager.CreateMsgSetRequest("US", 4, 0)
        addr4supported = True
    ElseIf (supportedVersion >= "3.0") Then
        Set requestMsgSet = SessionManager.CreateMsgSetRequest("US", 3, 0)
        addr4supported = True
    ElseIf (supportedVersion >= "2.0") Then
        Set requestMsgSet = SessionManager.CreateMsgSetRequest("US", 2, 0)
        addr4supported = True
    ElseIf (supportedVersion = "1.1") Then
        Set requestMsgSet = SessionManager.CreateMsgSetRequest("US", 1, 1)
    Else
        'MsgBox "You are apparently running QuickBooks 2002 Release 1, we strongly recommend that you use QuickBooks' online update feature to obtain the latest fixes and enhancements", vbExclamation
        Set requestMsgSet = SessionManager.CreateMsgSetRequest("US", 1, 0)
    End If
    
    'Verify if the invoice we are gonna add to quickbooks is not created yet
    Dim invoiceExist As Integer
    Dim requestMsg As QBFC10Lib.IMsgSetRequest
    Set requestMsg = SessionManager.CreateMsgSetRequest("US", 1, 0)
    requestMsg.Attributes.OnError = roeContinue
    
    
    Dim icom As QBFC10Lib.IInvoiceQuery
    Set icom = requestMsg.AppendInvoiceQueryRq
    
    icom.ORInvoiceQuery.InvoiceFilter.ORRefNumberFilter.RefNumberFilter.MatchCriterion.SetValue mcEndsWith
    icom.ORInvoiceQuery.InvoiceFilter.ORRefNumberFilter.RefNumberFilter.RefNumber.SetValue target
    
    Dim responseSet As QBFC10Lib.IMsgSetResponse
    Set responseSet = SessionManager.DoRequests(requestMsg)
    
    Dim rp As QBFC10Lib.IResponse
    Set rp = responseSet.ResponseList.GetAt(0)
    
    If (rp.StatusCode = 0) Then
        invoiceExist = MsgBox("This invoice number is already registered in QuickBooks do you want to continue?", vbOKCancel)
    End If
    
    If (invoiceExist = vbCancel) Then
        'MsgBox "exiting"
        Exit Sub
    End If
        'MsgBox "Continue...."
    
    ' Initialize the message set request's attributes
    requestMsgSet.Attributes.OnError = roeStop

    ' Add the request to the message set request object
    Dim invoiceAdd As IInvoiceAdd
    Set invoiceAdd = requestMsgSet.AppendInvoiceAddRq
    
    ' Set the IinvoiceAdd field values
    invoiceAdd.RefNumber.SetValue target
    invoiceAdd.TxnDate.SetValue ActiveCell.Offset(0, 1)
    ActiveCell.Offset(0, 1).Select
    invoiceAdd.CustomerRef.FullName.SetValue ActiveCell.Offset(0, 2)
    ActiveCell.Offset(0, 2).Select
    invoiceAdd.PONumber.SetValue ActiveCell.Offset(0, 1)
    ActiveCell.Offset(0, 6).Select
    invoiceAdd.SalesRepRef.FullName.SetValue ActiveCell.Offset(0, 1)
    
    
    ' Create the first line item for the invoice
    Set invoiceLineAdd = invoiceAdd.ORInvoiceLineAddList.Append.invoiceLineAdd
    
    ' Set the values for the invoice line
    invoiceLineAdd.ItemRef.FullName.SetValue "Materias Primas"
    invoiceLineAdd.ORRatePriceLevel.Rate.SetValue ActiveCell.Offset(0, -7)
    ActiveCell.Offset(0, -9).Select

    ' Perform the request and obtain a response from QuickBooks
    Dim responseMsgSet As IMsgSetResponse
    Set responseMsgSet = SessionManager.DoRequests(requestMsgSet)
    
    ' Close the session and connection with QuickBooks.
    SessionManager.EndSession
    booSessionBegun = False
    SessionManager.CloseConnection
        
    ' Uncomment the following to see the request and response XML for debugging
    'MsgBox requestMsgSet.ToXMLString, vbOKOnly, "RequestXML"
    'MsgBox responseMsgSet.ToXMLString, vbOKOnly, "ResponseXML"
    
    ' Interpret the response
    Dim response As IResponse
    
    ' The response list contains only one response,
    ' which corresponds to our single request
    Set response = responseMsgSet.ResponseList.GetAt(0)
     
    'msg = "Status: Code = " & CStr(response.StatusCode) &
           ' ", Message = " & response.StatusMessage & _
            '", Severity = " & response.StatusSeverity & vbCrLf
        
    ' The Detail property of the IResponse object
    ' returns a Ret object for Add and Mod requests.
    ' In this case, the Ret object is IInvoiceRet.
    
    'For help finding out the Detail's type, uncomment the following
    'line:
    'MsgBox response.Detail.Type.GetAsString
    
    Dim invoiceRet As IInvoiceRet
    Set invoiceRet = response.Detail

    If (invoiceRet Is Nothing) Then
        MsgBox msg
        Exit Sub
    End If
    MsgBox "Finished. The invoice #: " & target & " is already registered in QuickBooks"
    ir
    Range("A3").Select
    Selection.EntireRow.Delete
    Exit Sub
    
ErrHandler:
    If Err.Number = &H80040416 Then
        MsgBox "You must have QuickBooks running with the company" & vbCrLf & _
               "file open to use this program."
        SessionManager.CloseConnection
        End
    ElseIf Err.Number = &H80040422 Then
        MsgBox "This QuickBooks company file is open in single user mode and" & vbCrLf & _
               "another application is already accessing it.  Please exit the" & vbCrLf & _
               "other application and run this application again."
        SessionManager.CloseConnection
        End
    Else
        MsgBox "HRESULT = " & Err.Number & " (" & Hex(Err.Number) & _
               ") " & vbCrLf & vbCrLf & Err.Description
    
        If booSessionBegun Then
            SessionManager.EndSession
            SessionManager.CloseConnection
        End If
    
        End
    End If
End Sub
