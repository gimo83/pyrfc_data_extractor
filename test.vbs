' Register the SAP RFC library first (run as admin):
' regsvr32 sapnwrfc.dll


' Basic subroutine definition
Sub getPOInvoice(po)


    '--------------------------------------------------
    ' Call RFC function Read PO history
    '--------------------------------------------------
    Set rfcFunc = rfcLib.Add("RFC_READ_TABLE")
    ' Set parameters
    rfcFunc.Exports("QUERY_TABLE").Value = "EKBE"
    rfcFunc.Exports("DELIMITER").Value = "|"
    rfcFunc.Exports("NO_DATA").Value = " "
    ' rfcFunc.Exports("FIELDS").Value = 
    rfcFunc.Exports("ROWSKIPS").Value = "0"
    rfcFunc.Exports("ROWCOUNT").Value = "0"
    ' rfcFunc.Exports("OPTIONS").Value = 

    ' --------------------------------------------------
    ' Set FIELDS parameter (table of fields to return)
    ' --------------------------------------------------
    Set fieldsTable = rfcFunc.Tables("FIELDS")
    fieldsTable.AppendRow
    fieldsTable(1, "FIELDNAME") = "EBELN"  ' PO number
    fieldsTable.AppendRow
    fieldsTable(2, "FIELDNAME") = "EBELP"  ' PO item
    fieldsTable.AppendRow
    fieldsTable(3, "FIELDNAME") = "ZEKKN"  ' movement type
    fieldsTable.AppendRow
    fieldsTable(4, "FIELDNAME") = "VGABE"  ' movement type
    fieldsTable.AppendRow
    fieldsTable(5, "FIELDNAME") = "GJAHR"  ' fiscal year
    fieldsTable.AppendRow
    fieldsTable(6, "FIELDNAME") = "BELNR"  ' document number
    fieldsTable.AppendRow
    fieldsTable(7, "FIELDNAME") = "BUZEI"  ' document number


    ' --------------------------------------------------
    ' Set OPTIONS parameter (WHERE clause conditions)
    ' --------------------------------------------------
    Set optionsTable = rfcFunc.Tables("OPTIONS")
    optionsTable.AppendRow
    optionsTable(1, "TEXT") = "EBELN = '"& po_num &"'"  ' Specific PO
    optionsTable.AppendRow
    optionsTable(1, "TEXT") = "VGABE = '2'"  ' 



    ' Execute
    If rfcFunc.Call = False Then
        WScript.Echo "RFC call failed"
    Else
        ' Get results
        result = rfcFunc.Imports("RESULT_PARAM").Value
        WScript.Echo "Result: " & result
    End If


End Sub



Set rfcLib = CreateObject("SAP.Functions")
rfcLib.Connection.ApplicationServer = "sapr3p.sbm.com.sa"
rfcLib.Connection.SystemNumber = "00"
rfcLib.Connection.Client = "810"
rfcLib.Connection.User = "e003154"
' rfcLib.Connection.Password = "yourpass"

If rfcLib.Connection.Logon(0, False) <> True Then
    WScript.Echo "Logon failed"
    WScript.Quit
End If


getPOInvoice "4500001234"



rfcLib.Connection.Logoff