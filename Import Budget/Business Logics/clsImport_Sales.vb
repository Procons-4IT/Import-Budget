Imports System.IO
Imports System.Text.RegularExpressions
Imports System.Text
Imports System
Imports Microsoft.Office.Interop.Excel

'Work Pending
'Series
'Testing
'Delivery

Public Class clsImport_Sales
    Inherits clsBase
    Private strQuery As String
    Private oRecordSet As SAPbobsCOM.Recordset
    Private oComboBox As SAPbouiCOM.ComboBox
    Private oGrid As SAPbouiCOM.Grid
    Private oComboColumn As SAPbouiCOM.ComboBoxColumn
    Private oEditColumn As SAPbouiCOM.EditTextColumn
    Private oDt_Import As SAPbouiCOM.DataTable
    Private oDt_ErrorLog As SAPbouiCOM.DataTable

    Public Sub New()
        MyBase.New()
    End Sub

    Public Sub LoadForm()
        Try
            oForm = oApplication.Utilities.LoadForm(xml_SalImp, frm_SalImp)
            oForm = oApplication.SBO_Application.Forms.ActiveForm()
            oForm.Freeze(True)
            Initialize(oForm)
            oForm.Freeze(False)
        Catch ex As Exception
            oForm.Freeze(False)
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Public Sub LoadPriceHistory(ByVal aCode As String, ByVal aItemCode As String)
        Try
            oForm = oApplication.Utilities.LoadForm("frm_PriceHistory.xml", "frm_PriceHistory")
            oForm = oApplication.SBO_Application.Forms.ActiveForm()
            oForm.Freeze(True)
            Dim oEdit As SAPbouiCOM.EditText
            oEdit = oForm.Items.Item("3").Specific
            oEdit.String = aCode
            Dim strQry As String
            strQry = " SELECT Top 20 ItemCode,T1.[Dscription] 'Item Name',Quantity,Price,DocNum 'Document Number',T0.DocEntry ,T0.DocDate 'Document Date',isnull(T2.[FirstName],'') + ' ' + isnull(T2.MiddleName,'') + ' ' + isnull(T2.lastName,'')  'Document Owner' FROM OINV T0  INNER JOIN INV1 T1 ON T0.[DocEntry] = T1.[DocEntry] left outer JOIN OHEM T2 ON T0.[OwnerCode] = T2.[empID] where T0.CardCode='" & aCode & "' and T1.ItemCode='" & aItemCode & "' order by T0.DocEntry Desc"
            oGrid = oForm.Items.Item("4").Specific
            oGrid.DataTable.ExecuteQuery(strQry)
            oEditColumn = oGrid.Columns.Item("ItemCode")
            oEditColumn.LinkedObjectType = "4"
            oEditColumn = oGrid.Columns.Item("DocEntry")
            oEditColumn.LinkedObjectType = "13"
            For intRow As Integer = 0 To oGrid.DataTable.Rows.Count - 1
                oGrid.RowHeaders.SetText(intRow, intRow + 1)
            Next
            oGrid.Columns.Item("Item Name").Visible = False
            oGrid.AutoResizeColumns()


            strQry = " SELECT Top 20 ItemCode,T1.[Dscription] 'Item Name',Quantity,Price,DocNum 'Document Number',CardCode,CardName 'Customer Name', T0.DocEntry,T0.DocDate 'Document Date',isnull(T2.[FirstName],'') + ' ' + isnull(T2.MiddleName,'') + ' ' + isnull(T2.lastName,'')  'Document Owner'  FROM OINV T0  INNER JOIN INV1 T1 ON T0.[DocEntry] = T1.[DocEntry]  left outer JOIN OHEM T2 ON T0.[OwnerCode] = T2.[empID] where  T1.ItemCode='" & aItemCode & "' order by T0.DocEntry Desc"
            oGrid = oForm.Items.Item("8").Specific
            oGrid.DataTable.ExecuteQuery(strQry)

            oEditColumn = oGrid.Columns.Item("ItemCode")
            oEditColumn.LinkedObjectType = "4"
            oEditColumn = oGrid.Columns.Item("CardCode")
            oEditColumn.LinkedObjectType = "2"
            oEditColumn = oGrid.Columns.Item("DocEntry")
            oEditColumn.LinkedObjectType = "13"
            oForm.Items.Item("8").Enabled = False
            For intRow As Integer = 0 To oGrid.DataTable.Rows.Count - 1
                oGrid.RowHeaders.SetText(intRow, intRow + 1)
            Next
            oGrid.Columns.Item("Item Name").Visible = False
            oGrid.AutoResizeColumns()
            oForm.PaneLevel = 1
            oForm.Freeze(False)
        Catch ex As Exception
            oForm.Freeze(False)
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub


    Public Sub LoadStockDetails(ByVal aCode As String, ByVal aItemCode As String)
        Try
            oForm = oApplication.Utilities.LoadForm("frm_StockDetails.xml", "frm_StockDetails")
            oForm = oApplication.SBO_Application.Forms.ActiveForm()
            oForm.Freeze(True)
            Dim oEdit As SAPbouiCOM.EditText
            oEdit = oForm.Items.Item("3").Specific
            oEdit.String = aCode
            Dim strQry As String
            strQry = "  SELECT T0.[ItemCode], T0.[WhsCode], T0.[OnHand], T0.[IsCommited] ,T0.[OnOrder],isnull(T0.OnHand,0) - isnull(T0.IsCommited,0) +Isnull(T0.[OnOrder],0)  'Available' FROM OITW T0 where T0.ItemCode='" & aItemCode & "' order by T0.WhsCode"
            oGrid = oForm.Items.Item("5").Specific
            oGrid.DataTable.ExecuteQuery(strQry)
            oEdit = oForm.Items.Item("3").Specific
            oEdit.String = aItemCode
            oEditColumn = oGrid.Columns.Item("WhsCode")
            oEditColumn.LinkedObjectType = "64"
            oGrid.Columns.Item("IsCommited").Visible = False
            oGrid.Columns.Item("OnOrder").Visible = False
            oGrid.Columns.Item("Available").Visible = False
            oEditColumn = oGrid.Columns.Item("OnHand")
            oEditColumn.ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto

            oEditColumn = oGrid.Columns.Item("Available")
            oEditColumn.ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto

            For intRow As Integer = 0 To oGrid.DataTable.Rows.Count - 1
                oGrid.RowHeaders.SetText(intRow, intRow + 1)
            Next
            oGrid.AutoResizeColumns()
            oForm.PaneLevel = 1
            oForm.Freeze(False)
        Catch ex As Exception
            oForm.Freeze(False)
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub


    Private Function ReadExcelData_New(ByVal aForm As SAPbouiCOM.Form, ByVal strID As String) As Boolean
        Dim strWhs As String = "" 'CType(aForm.Items.Item("30").Specific, SAPbouiCOM.EditText).Value
        Dim strPath As String = CType(aForm.Items.Item(strID).Specific, SAPbouiCOM.StaticText).Caption
        Dim strItemCode As String
        Dim strCardCode As String = "" ''CType(aForm.Items.Item("3").Specific, SAPbouiCOM.EditText).Value
        Dim excel As Application = New Application
        ' Open Excel spreadsheet.
        Dim w As Workbook = excel.Workbooks.Open(strPath)
        Dim dblAmt As Single
        Try
            Dim ostatic As SAPbouiCOM.StaticText
            ostatic = aForm.Items.Item("78").Specific '
            Dim oTest As SAPbobsCOM.Recordset
            oTest = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Try
                '  oTest.DoQuery("Delete from [Z_SIIM1]")

            Catch ex As Exception

            End Try
            Dim strCode As String = "1" ' oApplication.Utilities.getMaxCode("@S_QImport", "Code")
            'Dim s1 As String = "Insert Into [@S_QImport] values('" & strCode & "','" & strCode & "')"
            'oTest.DoQuery(s1)
           
            Dim oEdit As SAPbouiCOM.EditText
            oEdit = aForm.Items.Item("81").Specific
            oEdit.String = strCode
            ' CType(oForm.Items.Item("30").Specific
            Dim strWhsCode As String = "" 'oApplication.Utilities.getEditTextvalue(aForm, "30")
            oTest.DoQuery("Delete From [Z_OBDG] where RefCode='" & strCode & "'")
            Dim strRefCode As String
            strRefCode = oApplication.Utilities.getEditTextvalue(aForm, "81")
            ' Loop over all sheets.
            For i As Integer = 1 To w.Sheets.Count
                ' Get sheet.
                Dim sheet As Worksheet = w.Sheets(i)
                ' Get range.
                Dim r As Range = sheet.UsedRange
                ' Load all cells into 2d array.
                Dim array(,) As Object = r.Value(XlRangeValueDataType.xlRangeValueDefault)
                oApplication.Utilities.createDatatable()
                ' Scan the cells.
                If array IsNot Nothing Then
                    Dim bound0 As Integer = array.GetUpperBound(0)
                    Dim bound1 As Integer = array.GetUpperBound(1)
                    ' Loop over all elements.
                    _oDt.Rows.Clear()
                   
                    For j As Integer = 1 To bound0
                        ostatic.Caption = "Reading Excel Data " & j & " of " & bound0
                        If j <> 1 Then
                            'oApplication.Utilities.ReadExcelDat_Budget(j, array(j, 1), array(j, 2), array(j, 3), array(j, 4), array(j, 5), array(j, 6), array(j, 7), array(j, 8), array(j, 9), array(j, 10), aForm, j - 1, strRefCode, strWhsCode)
                            oApplication.Utilities.ReadExcelDat_Budget(oForm, j, array(j, 1), array(j, 2), array(j, 2), array(j, 3), array(j, 4), array(j, 5), array(j, 6), array(j, 7), array(j, 8), array(j, 9), array(j, 10), array(j, 11), array(j, 12), array(j, 13), array(j, 14), array(j, 15), array(j, 16), strRefCode)

                        End If
                    Next
                End If
            Next
            ostatic.Caption = "Readin Excel Data Completed"
            w.Close()
            excel.Quit()
            If Not IsNothing(_oDt) Then
                If _oDt.Rows.Count > 0 Then
                    ' Dim _unPivot As DataTable = unPivotTable(oForm, _oDt)
                    Dim strDtXML As String = oApplication.Utilities.getXMLstring(_oDt) ' _unPivot)
                    Dim s As String = "Exec [Insert_Budget] '" + strDtXML + "'"
                    oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    oRecordSet.DoQuery("Exec [Insert_Budget] '" + strDtXML + "'")
                    Return True
                End If
            End If
        Catch ex As Exception
            w.Close()
            excel.Quit()
            Throw ex
        Finally
            ReleaseComObject(w)
            ReleaseComObject(excel)
        End Try
        Return True
    End Function

    Private Sub ReleaseComObject(ByVal obj As Object)
        Try
            System.Runtime.InteropServices.Marshal.ReleaseComObject(obj)
            obj = Nothing
        Catch ex As Exception
            obj = Nothing
        End Try
    End Sub

    Private Sub UpdateRequiredQty(aForm As SAPbouiCOM.Form)
        Dim strWhs As String
        Dim oEdit As SAPbouiCOM.EditText
        oEdit = aForm.Items.Item("83").Specific
        strWhs = oEdit.String
        aForm.Freeze(True)
        Dim strCode As String
        '  Dim oEdit As SAPbouiCOM.EditText
        oEdit = oForm.Items.Item("81").Specific
        strCode = oEdit.String
        Try
            If 1 = 1 Then ' strWhs <> "" Then
                Dim oTest As SAPbobsCOM.Recordset
                oTest = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oGrid = aForm.Items.Item("68").Specific
                oApplication.Utilities.Message("Processing....", SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                For intRow As Integer = 0 To oGrid.DataTable.Rows.Count - 1

                    If strWhs = "" Then
                        strWhs = oGrid.DataTable.GetValue("Warehouse", intRow)
                    End If
                    ' oTemp.DoQuery("Select  OnHand,(OnHand-IsCommited+OnOrder) 'Available'  from OITW where whscode='" & strWhs & "' and  ItemCode='" & aItemCode.Replace("-", "") & "'")

                    oTest.DoQuery(" SELECT Sum(T0.[OnHand]) FROM OITW T0 where  T0.ItemCode='" & oGrid.DataTable.GetValue("CODE", intRow) & "'")
                    Dim Aqty As Double = oTest.Fields.Item(0).Value


                    oTest.DoQuery(" SELECT T0.[ItemCode], T0.[WhsCode], T0.[OnHand], T0.[IsCommited] ,T0.[OnOrder],isnull(T0.OnHand,0) - isnull(T0.IsCommited,0) +Isnull(T0.[OnOrder],0)  'Available' FROM OITW T0 where T0.WhsCode='" & strWhs & "' and  T0.ItemCode='" & oGrid.DataTable.GetValue("CODE", intRow) & "' order by T0.WhsCode")
                    If oTest.RecordCount > 0 Then
                        '      MsgBox(oTest.Fields.Item("Available").Value)
                        If (oGrid.DataTable.GetValue("OnHand", intRow) < oGrid.DataTable.GetValue("QTY", intRow)) And oTest.Fields.Item("OnHand").Value > 0 Then
                            oGrid.DataTable.SetValue("Warehouse", intRow, strWhs)
                            oGrid.DataTable.SetValue("OnHand", intRow, oTest.Fields.Item("OnHand").Value)
                            oGrid.DataTable.SetValue("AQTY", intRow, Aqty) 'oTest.Fields.Item("Available").Value)
                            If oGrid.DataTable.GetValue("QTY", intRow) >= oTest.Fields.Item("OnHand").Value Then
                                oGrid.DataTable.SetValue("Warehouse", intRow, strWhs)
                                ''  oGrid.DataTable.SetValue("OnHand", intRow, oTest.Fields.Item("OnHand").Value)
                                '  oGrid.DataTable.SetValue("AQTY", intRow, oTest.Fields.Item("Available").Value)
                                oGrid.DataTable.SetValue("QTY", intRow, oTest.Fields.Item("OnHand").Value)
                                strQuery = "Update Z_SIIM1 set WhsCode='" & strWhs & "',OnHand='" & oTest.Fields.Item("OnHand").Value & "',AQTY='" & Aqty & "', Price='" & oGrid.DataTable.GetValue("Price", intRow) & "', Qty='" & oGrid.DataTable.GetValue("QTY", intRow) & "'  where  RefCode='" & strCode & "' and  RowID='" & oGrid.DataTable.GetValue("RowID", intRow) & "'"
                                oTest.DoQuery(strQuery)
                            Else
                                '   oGrid.DataTable.SetValue("QTY", intRow, oTest.Fields.Item("OnHand").Value)

                                strQuery = "Update Z_SIIM1 set WhsCode='" & strWhs & "',OnHand='" & oTest.Fields.Item("OnHand").Value & "',AQTY='" & Aqty & "', Price='" & oGrid.DataTable.GetValue("Price", intRow) & "', Qty='" & oGrid.DataTable.GetValue("QTY", intRow) & "'  where  RefCode='" & strCode & "' and  RowID='" & oGrid.DataTable.GetValue("RowID", intRow) & "'"
                                oTest.DoQuery(strQuery)
                            End If
                        End If
                    End If
                Next
            End If
            formatGrid(aForm)
            oApplication.Utilities.Message("Operation completed successfully....", SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            aForm.Freeze(False)
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            aForm.Freeze(False)
        End Try
    End Sub
#Region "Menu Event"
    Public Overrides Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        Try
            oForm = oApplication.SBO_Application.Forms.ActiveForm
            Select Case pVal.MenuUID
                Case mnu_SalImp
                    LoadForm()
                Case mnu_FIRST, mnu_LAST, mnu_NEXT, mnu_PREVIOUS
                Case mnu_ADD
            End Select
        Catch ex As Exception

        End Try
    End Sub
#End Region

#Region "Item Event"
    Public Overrides Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.FormTypeEx = frm_SalImp Then
                Select Case pVal.BeforeAction
                    Case True
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "68" And pVal.ColUID = "AQTY" Then
                                    oGrid = oForm.Items.Item("68").Specific
                                    LoadStockDetails(CType(oForm.Items.Item("3").Specific, SAPbouiCOM.EditText).Value, oGrid.DataTable.GetValue("CODE", pVal.Row))
                                    BubbleEvent = False
                                    Exit Sub
                                End If
                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "57" Then 'Browse
                                    oApplication.Utilities.OpenFileDialogBox(oForm, "56")
                                    'ElseIf pVal.ItemUID = "62" And oForm.PaneLevel > 2 Then
                                    '    If Validation(oForm) = False Then
                                    '        BubbleEvent = False
                                    '        Exit Sub
                                    '    End If
                                ElseIf (pVal.ItemUID = "58") Then 'Next
                                    oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                    If 1 = 1 Then 'CType(oForm.Items.Item("17").Specific, SAPbouiCOM.StaticText).Caption <> "" Then
                                        If oApplication.Utilities.ValidateFile(oForm, "56") Then
                                            If ReadExcelData_New(oForm, "56") Then ' oApplication.Utilities.GetExcelData(oForm, "56") Then
                                                oForm.Items.Item("71").Enabled = True
                                                oApplication.Utilities.Message(" Data Imported Successfully....", SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                                                oApplication.Utilities.Message("Press Next to Proceed....", SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                                            Else
                                                BubbleEvent = False
                                            End If
                                        End If
                                    Else
                                        oApplication.Utilities.Message("Select File to Import....", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                        BubbleEvent = False
                                    End If
                                ElseIf (pVal.ItemUID = "63") Then
                                    If Not validate(oForm) Then
                                        BubbleEvent = False
                                    Else
                                        changeFormPane(oForm, pVal.ItemUID)
                                    End If
                                ElseIf (pVal.ItemUID = "65") Then
                                    changeFormPane(oForm, pVal.ItemUID)
                                ElseIf (pVal.ItemUID = "66") Then
                                    changeFormPane(oForm, pVal.ItemUID)
                                ElseIf (pVal.ItemUID = "67") Then
                                    changeFormPane(oForm, pVal.ItemUID)
                                ElseIf (pVal.ItemUID = "70") Then
                                    changeFormPane(oForm, pVal.ItemUID)
                                ElseIf (pVal.ItemUID = "71") Then
                                    loadData(oForm)
                                    changeFormPane(oForm, pVal.ItemUID)
                                ElseIf (pVal.ItemUID = "62") Then
                                    oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                    If oForm.PaneLevel > 2 Then
                                        If Validation(oForm) = False Then
                                            BubbleEvent = False
                                            Exit Sub
                                        End If
                                    End If
                                    Dim _retVal As Integer = oApplication.SBO_Application.MessageBox("Proceed with Import Budget Details?", 2, "Yes", "No", "")
                                    If _retVal = 1 Then
                                        ' UpdateRecords(oForm)
                                        If oApplication.Utilities.saveAsDraft(oForm) Then
                                            oApplication.Utilities.Message("Budget Imported Sucessfully....", SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                                            oForm.Close()
                                        Else
                                            oApplication.Utilities.Message(oApplication.Company.GetLastErrorDescription(), SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                        End If
                                    Else
                                        BubbleEvent = False
                                    End If
                                ElseIf (pVal.ItemUID = "64" Or pVal.ItemUID = "72" Or pVal.ItemUID = "59") Then
                                    oForm.Close()
                                End If
                        End Select
                    Case False
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "82" Then
                                    If oApplication.SBO_Application.MessageBox("Do you want to Update the Requested Quantities?", , "Continue", "Cancel") = 2 Then
                                        Exit Sub
                                    End If
                                    UpdateRequiredQty(oForm)
                                End If

                            Case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "68" And pVal.ColUID = "Price" Then
                                    oGrid = oForm.Items.Item("68").Specific
                                    LoadPriceHistory(CType(oForm.Items.Item("3").Specific, SAPbouiCOM.EditText).Value, oGrid.DataTable.GetValue("CODE", pVal.Row))
                                End If
                            Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                Dim oCFLEvento As SAPbouiCOM.IChooseFromListEvent
                                Dim oDataTable As SAPbouiCOM.DataTable
                                Dim strCustomer, strName, strCurrency, strWareHouse, strTaxCode, strAcctCode As String
                                Try
                                    oCFLEvento = pVal
                                    oDataTable = oCFLEvento.SelectedObjects
                                    If pVal.ItemUID = "3" And Not IsNothing(oDataTable) Then
                                        strCustomer = oDataTable.GetValue("CardCode", 0)
                                        strName = oDataTable.GetValue("CardName", 0)
                                        strCurrency = oDataTable.GetValue("Currency", 0)
                                        Try
                                            CType(oForm.Items.Item("3").Specific, SAPbouiCOM.EditText).Value = strCustomer
                                            CType(oForm.Items.Item("5").Specific, SAPbouiCOM.EditText).Value = strName
                                        Catch ex As Exception
                                            CType(oForm.Items.Item("3").Specific, SAPbouiCOM.EditText).Value = strCustomer
                                            CType(oForm.Items.Item("5").Specific, SAPbouiCOM.EditText).Value = strName
                                        End Try
                                        loadComboBasedOnCustomer(oForm, strCustomer)
                                    ElseIf (pVal.ItemUID = "30") And Not IsNothing(oDataTable) Then
                                        strWareHouse = oDataTable.GetValue("WhsCode", 0)
                                        CType(oForm.Items.Item("30").Specific, SAPbouiCOM.EditText).Value = strWareHouse
                                    ElseIf (pVal.ItemUID = "46") And Not IsNothing(oDataTable) Then
                                        strTaxCode = oDataTable.GetValue("Code", 0)
                                        CType(oForm.Items.Item("46").Specific, SAPbouiCOM.EditText).Value = strTaxCode
                                    ElseIf (pVal.ItemUID = "49") And Not IsNothing(oDataTable) Then
                                        strAcctCode = oDataTable.GetValue("AcctCode", 0)
                                        CType(oForm.Items.Item("49").Specific, SAPbouiCOM.EditText).Value = strAcctCode
                                    ElseIf pVal.ItemUID = "83" Then
                                        strAcctCode = oDataTable.GetValue("WhsCode", 0)
                                        CType(oForm.Items.Item("83").Specific, SAPbouiCOM.EditText).Value = strAcctCode
                                    ElseIf pVal.ItemUID = "68" And pVal.ColUID = "Warehouse" Then
                                        Dim oTemp As SAPbobsCOM.Recordset
                                        oGrid = oForm.Items.Item(pVal.ItemUID).Specific
                                        oGrid.DataTable.SetValue("Warehouse", pVal.Row, oDataTable.GetValue("WhsCode", 0))
                                        oTemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                        oTemp.DoQuery("Select  OnHand,(OnHand-IsCommited+OnOrder) 'Available'  from OITW where whscode='" & oDataTable.GetValue("WhsCode", 0) & "' and  ItemCode='" & oGrid.DataTable.GetValue("CODE", pVal.Row).ToString.Replace("-", "") & "'")
                                        oGrid.DataTable.SetValue("OnHand", pVal.Row, oTemp.Fields.Item(0).Value)
                                        oGrid.DataTable.SetValue("AQTY", pVal.Row, oTemp.Fields.Item(1).Value)
                                    End If
                                Catch ex As Exception

                                End Try
                            Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "33" Then
                                    fillAddress(oForm, CType(oForm.Items.Item("3").Specific, SAPbouiCOM.EditText).Value, "S")
                                ElseIf pVal.ItemUID = "38" Then
                                    fillAddress(oForm, CType(oForm.Items.Item("3").Specific, SAPbouiCOM.EditText).Value, "B")
                                ElseIf pVal.ItemUID = "44" Then
                                    fillDueDate(oForm, CType(oForm.Items.Item("44").Specific, SAPbouiCOM.ComboBox).Selected.Value)
                                ElseIf pVal.ItemUID = "77" Then
                                    loadData(oForm)
                                End If
                        End Select
                End Select
            End If
        Catch ex As Exception
            oForm.Freeze(False)
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
#End Region

#Region "Data Events"
    Public Sub FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
        Try

        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
#End Region

#Region "Right Click"
    Public Sub RightClickEvent(ByRef eventInfo As SAPbouiCOM.ContextMenuInfo, ByRef BubbleEvent As Boolean)
        Try
            oForm = oApplication.SBO_Application.Forms.Item(eventInfo.FormUID)
            If oForm.TypeEx = frm_SalImp Then

            End If
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
#End Region

#Region "Function"

    Private Sub Initialize(ByVal oForm As SAPbouiCOM.Form)
        Try
            oForm.PaneLevel = 2
            'CType(oForm.Items.Item("24").Specific, SAPbouiCOM.EditText).Value = System.DateTime.Now.ToString("yyyyMMdd")
            'CType(oForm.Items.Item("26").Specific, SAPbouiCOM.EditText).Value = System.DateTime.Now.ToString("yyyyMMdd")
            'CType(oForm.Items.Item("28").Specific, SAPbouiCOM.EditText).Value = System.DateTime.Now.ToString("yyyyMMdd")
            oForm.DataSources.UserDataSources.Add("filter", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
            oForm.DataSources.UserDataSources.Add("RefNo", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
            'oComboBox = oForm.Items.Item("77").Specific
            'oComboBox.ValidValues.Add("0", "All Items")
            'oComboBox.ValidValues.Add("O", "On Hand")
            'oComboBox.ValidValues.Add("A", "Available Qty")
            'oComboBox.ValidValues.Add("P", "Partial")
            'oComboBox.Select(0, SAPbouiCOM.BoSearchKey.psk_Index)
            'oComboBox.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly
            Dim oEdit As SAPbouiCOM.EditText
            'oEdit = oForm.Items.Item("30").Specific
            'oEdit.String = "01"
            oEdit = oForm.Items.Item("81").Specific
            oEdit.DataBind.SetBound(True, "", "RefNo")
            'initializeDataTable(oForm)
            'addStyle(oForm)
            'loadCombo(oForm)
            'cflFilter(oForm)

            ''Select User Series
            'oComboBox = oForm.Items.Item("20").Specific
            'Dim strUserSeries As String = getUserSeries(oForm).ToString()
            'If strUserSeries.Length > 0 And strUserSeries <> "0" Then
            '    oComboBox.Select(strUserSeries, SAPbouiCOM.BoSearchKey.psk_ByValue)
            'End If

            'oForm.Items.Item("3").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub initializeDataTable(ByVal oForm As SAPbouiCOM.Form)
        Try
            oForm.DataSources.DataTables.Add("Dt_Import")
            oForm.DataSources.DataTables.Add("Dt_ErrorLog")

            oDt_Import = oForm.DataSources.DataTables.Item("Dt_Import")
            oDt_Import.ExecuteQuery("Select CODE,[DESCRIPTION],QTY,OnHand,AQTY,Price,RowID From Z_SIIM1 Where 1 = 2 ")
            oGrid = oForm.Items.Item("68").Specific
            oGrid.DataTable = oDt_Import

            oDt_ErrorLog = oForm.DataSources.DataTables.Item("Dt_ErrorLog")
            oDt_ErrorLog.ExecuteQuery("Select CODE,[DESCRIPTION],QTY,AQTY,ERROR As 'Error' From Z_SIIM1 Where 1 = 2 ")
            oGrid = oForm.Items.Item("69").Specific
            oGrid.DataTable = oDt_ErrorLog

            formatGrid(oForm)
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub addStyle(ByVal oForm As SAPbouiCOM.Form)
        Try
            oForm.Items.Item("1").TextStyle = 7
            oForm.Items.Item("31").TextStyle = 7
            oForm.Items.Item("40").TextStyle = 7
            oForm.Items.Item("54").TextStyle = 7
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub loadCombo(ByVal oForm As SAPbouiCOM.Form)
        Try
            oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

            'Series Fill
            oComboBox = oForm.Items.Item("20").Specific
            strQuery = "Select Series,SeriesName From NNM1 Where ObjectCode = '23'"
            oRecordSet.DoQuery(strQuery)
            If Not oRecordSet.EoF Then
                oComboBox.ValidValues.Add("", "")
                While Not oRecordSet.EoF
                    oComboBox.ValidValues.Add(oRecordSet.Fields.Item("Series").Value, oRecordSet.Fields.Item("SeriesName").Value)
                    oRecordSet.MoveNext()
                End While
            End If
            oComboBox.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly

            'Currency List
            oComboBox = oForm.Items.Item("11").Specific
            strQuery = "Select CurrCode,CurrName From OCRN"
            oRecordSet.DoQuery(strQuery)
            If Not oRecordSet.EoF Then
                oComboBox.ValidValues.Add("", "")
                While Not oRecordSet.EoF
                    oComboBox.ValidValues.Add(oRecordSet.Fields.Item("CurrCode").Value, oRecordSet.Fields.Item("CurrName").Value)
                    oRecordSet.MoveNext()
                End While
            End If
            oComboBox.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly

            'Price List
            oComboBox = oForm.Items.Item("14").Specific
            strQuery = "Select ListNum,ListName From OPLN"
            oRecordSet.DoQuery(strQuery)
            If Not oRecordSet.EoF Then
                oComboBox.ValidValues.Add("", "")
                While Not oRecordSet.EoF
                    oComboBox.ValidValues.Add(oRecordSet.Fields.Item("ListNum").Value, oRecordSet.Fields.Item("ListName").Value)
                    oRecordSet.MoveNext()
                End While
            End If
            oComboBox.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly

            'Sales Employee
            oComboBox = oForm.Items.Item("16").Specific
            strQuery = "Select SlpCode,SlpName From OSLP"
            oRecordSet.DoQuery(strQuery)
            If Not oRecordSet.EoF Then
                oComboBox.ValidValues.Add("", "")
                While Not oRecordSet.EoF
                    oComboBox.ValidValues.Add(oRecordSet.Fields.Item("SlpCode").Value, oRecordSet.Fields.Item("SlpName").Value)
                    oRecordSet.MoveNext()
                End While
            End If
            oComboBox.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly

            'Ship Type
            oComboBox = oForm.Items.Item("36").Specific
            strQuery = "Select TrnspCode,TrnspName From OSHP"
            oRecordSet.DoQuery(strQuery)
            If Not oRecordSet.EoF Then
                oComboBox.ValidValues.Add("", "")
                While Not oRecordSet.EoF
                    oComboBox.ValidValues.Add(oRecordSet.Fields.Item("TrnspCode").Value, oRecordSet.Fields.Item("TrnspName").Value)
                    oRecordSet.MoveNext()
                End While
            End If
            oComboBox.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly

            'Payment Type
            oComboBox = oForm.Items.Item("44").Specific
            strQuery = "Select GroupNum,PymntGroup From OCTG"
            oRecordSet.DoQuery(strQuery)
            If Not oRecordSet.EoF Then
                oComboBox.ValidValues.Add("", "")
                While Not oRecordSet.EoF
                    oComboBox.ValidValues.Add(oRecordSet.Fields.Item("GroupNum").Value, oRecordSet.Fields.Item("PymntGroup").Value)
                    oRecordSet.MoveNext()
                End While
            End If
            oComboBox.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly

            'Tax Code
            oComboBox = oForm.Items.Item("46").Specific
            strQuery = " Select Code,Name From OVTG Where Category = 'O' And Inactive = 'N' "
            oRecordSet.DoQuery(strQuery)
            If Not oRecordSet.EoF Then
                oComboBox.ValidValues.Add("", "")
                While Not oRecordSet.EoF
                    oComboBox.ValidValues.Add(oRecordSet.Fields.Item("Code").Value, oRecordSet.Fields.Item("Name").Value)
                    oRecordSet.MoveNext()
                End While
            End If
            Try
                oComboBox.Select("X0", SAPbouiCOM.BoSearchKey.psk_ByValue)
            Catch ex As Exception

            End Try
            oComboBox.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly

            'Payment Methond
            oComboBox = oForm.Items.Item("51").Specific
            strQuery = "Select PaymethCod,Descript From OPYM"
            oRecordSet.DoQuery(strQuery)
            If Not oRecordSet.EoF Then
                oComboBox.ValidValues.Add("", "")
                While Not oRecordSet.EoF
                    oComboBox.ValidValues.Add(oRecordSet.Fields.Item("PaymethCod").Value, oRecordSet.Fields.Item("Descript").Value)
                    oRecordSet.MoveNext()
                End While
            End If
            oComboBox.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly


        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub cflFilter(ByVal oForm As SAPbouiCOM.Form)
        Try

            Dim oCFLs As SAPbouiCOM.ChooseFromListCollection
            Dim oCons As SAPbouiCOM.Conditions
            Dim oCon As SAPbouiCOM.Condition
            oCFLs = oForm.ChooseFromLists
            Dim oCFL As SAPbouiCOM.ChooseFromList
            Dim oCFLCreationParams As SAPbouiCOM.ChooseFromListCreationParams
            oCFLCreationParams = oApplication.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)

            'Customer
            oCFL = oCFLs.Item("CFL_4")
            oCons = oCFL.GetConditions()
            oCon = oCons.Add()
            oCon.Alias = "CardType"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCon.CondVal = "C"
            oCFL.SetConditions(oCons)

            'Account
            oCFL = oCFLs.Item("CFL_10")
            oCons = oCFL.GetConditions()
            oCon = oCons.Add()
            oCon.Alias = "LocManTran"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCon.CondVal = "Y"
            oCFL.SetConditions(oCons)

            ''WareHouse
            'oCFL = oCFLs.Item("CFL_8")
            'oCons = oCFL.GetConditions()
            'oCon = oCons.Add()
            'oCon.Alias = "BinActivat"
            'oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            'oCon.CondVal = "Y"
            'oCFL.SetConditions(oCons)

            'oCFL = oCFLs.Item("CFL_10")
            'oCons = oCFL.GetConditions()
            'oCon = oCons.Add()
            'oCon.Alias = "Postable"
            'oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            'oCon.CondVal = "Y"
            'oCFL.SetConditions(oCons)

            'MessageBox.Show(oCFL.GetConditions().GetAsXML())

        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub loadComboBasedOnCustomer(ByVal oForm As SAPbouiCOM.Form, ByVal strCardCode As String)
        Dim oCustomer As SAPbobsCOM.BusinessPartners
        oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oCustomer = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oBusinessPartners)
        Try
            oForm.Freeze(True)

            'Contact Person           
            oApplication.Utilities.RemoveComboValues(oForm, "7")
            oComboBox = oForm.Items.Item("7").Specific
            strQuery = "Select CntctCode,Name From OCPR Where CardCode = '" + strCardCode + "' And Active = 'Y'"
            oRecordSet.DoQuery(strQuery)
            If Not oRecordSet.EoF Then
                oComboBox.ValidValues.Add("", "")
                While Not oRecordSet.EoF
                    oComboBox.ValidValues.Add(oRecordSet.Fields.Item("CntctCode").Value, oRecordSet.Fields.Item("Name").Value)
                    oRecordSet.MoveNext()
                End While
            Else
                oComboBox.ValidValues.Add("", "")
                oComboBox.SelectExclusive(0, SAPbouiCOM.BoSearchKey.psk_Index)
            End If
            oComboBox.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly

            'Ship To
            oApplication.Utilities.RemoveComboValues(oForm, "33")
            oComboBox = oForm.Items.Item("33").Specific
            strQuery = "Select LineNum,Address From CRD1 Where CardCode = '" + strCardCode + "' And AdresType = 'S'"
            oRecordSet.DoQuery(strQuery)
            If Not oRecordSet.EoF Then
                oComboBox.ValidValues.Add("", "")
                While Not oRecordSet.EoF
                    oComboBox.ValidValues.Add(oRecordSet.Fields.Item("LineNum").Value, oRecordSet.Fields.Item("Address").Value)
                    oRecordSet.MoveNext()
                End While
            Else
                oComboBox.ValidValues.Add("", "")
                oComboBox.SelectExclusive(0, SAPbouiCOM.BoSearchKey.psk_Index)
            End If
            oComboBox.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly

            'Bill To
            oApplication.Utilities.RemoveComboValues(oForm, "38")
            oComboBox = oForm.Items.Item("38").Specific
            strQuery = "Select LineNum,Address From CRD1 Where CardCode = '" + strCardCode + "' And AdresType = 'B'"
            oRecordSet.DoQuery(strQuery)
            If Not oRecordSet.EoF Then
                oComboBox.ValidValues.Add("", "")
                While Not oRecordSet.EoF
                    oComboBox.ValidValues.Add(oRecordSet.Fields.Item("LineNum").Value, oRecordSet.Fields.Item("Address").Value)
                    oRecordSet.MoveNext()
                End While
            Else
                oComboBox.ValidValues.Add("", "")
                oComboBox.SelectExclusive(0, SAPbouiCOM.BoSearchKey.psk_Index)
            End If
            oComboBox.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly

            If oCustomer.GetByKey(strCardCode) Then
                If oCustomer.Currency <> "##" Then
                    oComboBox = oForm.Items.Item("11").Specific 'Contact Person
                    If oCustomer.Currency <> "" Then
                        oComboBox.SelectExclusive(oCustomer.Currency, SAPbouiCOM.BoSearchKey.psk_ByValue)
                    End If
                End If

                oComboBox = oForm.Items.Item("7").Specific 'Contact Person
                If oCustomer.ContactPerson <> "" Then
                    oComboBox.SelectExclusive(oCustomer.ContactPerson, SAPbouiCOM.BoSearchKey.psk_ByDescription)
                End If


                oComboBox = oForm.Items.Item("14").Specific 'Price List
                If oCustomer.PriceListNum.ToString() <> "" Then
                    oComboBox.SelectExclusive(oCustomer.PriceListNum.ToString(), SAPbouiCOM.BoSearchKey.psk_ByValue)
                End If

                oComboBox = oForm.Items.Item("16").Specific 'Sales Employee
                If oCustomer.SalesPersonCode.ToString() <> "" Then
                    oComboBox.SelectExclusive(oCustomer.SalesPersonCode.ToString(), SAPbouiCOM.BoSearchKey.psk_ByValue)
                End If

                oComboBox = oForm.Items.Item("33").Specific 'Ship To
                If oCustomer.ShipToDefault <> "" Then
                    oComboBox.SelectExclusive(oCustomer.ShipToDefault, SAPbouiCOM.BoSearchKey.psk_ByDescription)
                Else
                    oComboBox.SelectExclusive(0, SAPbouiCOM.BoSearchKey.psk_Index)
                End If

                oComboBox = oForm.Items.Item("38").Specific 'Bill To
                If oCustomer.BilltoDefault <> "" Then
                    oComboBox.SelectExclusive(oCustomer.BilltoDefault, SAPbouiCOM.BoSearchKey.psk_ByDescription)
                Else
                    oComboBox.SelectExclusive(0, SAPbouiCOM.BoSearchKey.psk_Index)
                End If


                oComboBox = oForm.Items.Item("36").Specific 'Shiping Type
                If oCustomer.ShippingType.ToString() Then
                    oComboBox.SelectExclusive(oCustomer.ShippingType.ToString(), SAPbouiCOM.BoSearchKey.psk_ByValue)
                End If

                CType(oForm.Items.Item("53").Specific, SAPbouiCOM.EditText).Value = oCustomer.FederalTaxID ' Vat
                CType(oForm.Items.Item("42").Specific, SAPbouiCOM.EditText).Value = "Sales Quotation - " + strCardCode ' Journal Remrarks

                oComboBox = oForm.Items.Item("44").Specific 'Payment Type
                If oCustomer.PayTermsGrpCode.ToString() <> "" Then
                    oComboBox.SelectExclusive(oCustomer.PayTermsGrpCode.ToString(), SAPbouiCOM.BoSearchKey.psk_ByValue)
                End If

                oComboBox = oForm.Items.Item("51").Specific 'Payment Methond
                If oCustomer.PeymentMethodCode <> "" Then
                    oComboBox.SelectExclusive(oCustomer.PeymentMethodCode, SAPbouiCOM.BoSearchKey.psk_ByValue)
                End If

                'CType(oForm.Items.Item("49").Specific, SAPbouiCOM.EditText).Value = oCustomer.DebitorAccount ' Control Accon
            End If
            oForm.Freeze(False)
        Catch ex As Exception
            oForm.Freeze(False)
            Throw ex
        Finally
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oCustomer)
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet)
        End Try
    End Sub

    Private Function Validation(aform As SAPbouiCOM.Form) As Boolean
        oGrid = aform.Items.Item("68").Specific
        'For intRow As Integer = 0 To oGrid.DataTable.Rows.Count - 1
        '    If oGrid.DataTable.GetValue("QTY", intRow) > oGrid.DataTable.GetValue("OnHand", intRow) Then
        '        oApplication.Utilities.Message("Requested quantity should be less than or equal to On Quantity Quantity. Line NO : " & intRow + 1 & " : ItemCode : " & oGrid.DataTable.GetValue("CODE", intRow), SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        '        oGrid.Columns.Item("QTY").Click(intRow, False, 1)
        '        Return False
        '    End If
        'Next
        Return True
    End Function

    Private Function validate(ByVal oForm As SAPbouiCOM.Form) As Boolean
        Dim _retVal As Boolean = True
        Try
            If CType(oForm.Items.Item("27").Specific, SAPbouiCOM.EditText).Value = "" Then 'Card Code
                ' oApplication.Utilities.Message("Select Importing File....", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                ' _retVal = False
            End If
        Catch ex As Exception
            Throw ex
        End Try
        Return _retVal
    End Function

    Private Sub fillAddress(ByVal oForm As SAPbouiCOM.Form, ByVal strCardCode As String, ByVal strType As String)
        Try
            oForm.Freeze(True)
            oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            strQuery = "Select ISNULL(Block,'') + ' ' + ISNULL(Street,'') + ' ' + ISNULL(City,'') + ' ' + ISNULL(ZipCode,'') + ' ' + ISNULL(City,'')  From CRD1 Where CardCode = '" + strCardCode + "' And AdresType = '" + strType + "'"
            oRecordSet.DoQuery(strQuery)
            If Not oRecordSet.EoF Then
                If strType = "S" Then
                    CType(oForm.Items.Item("34").Specific, SAPbouiCOM.EditText).Value = oRecordSet.Fields.Item(0).Value
                Else
                    CType(oForm.Items.Item("39").Specific, SAPbouiCOM.EditText).Value = oRecordSet.Fields.Item(0).Value
                End If
            Else
                If strType = "S" Then
                    CType(oForm.Items.Item("34").Specific, SAPbouiCOM.EditText).Value = String.Empty
                Else
                    CType(oForm.Items.Item("39").Specific, SAPbouiCOM.EditText).Value = String.Empty
                End If
            End If
            oForm.Freeze(False)
        Catch ex As Exception
            oForm.Freeze(False)
            Throw ex
        End Try
    End Sub

    Private Sub fillDueDate(ByVal oForm As SAPbouiCOM.Form, ByVal strType As String)
        Try
            oForm.Freeze(True)
            oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            strQuery = "Select ExtraMonth,ExtraDays From OCTG Where GroupNum = '" + strType + "'"
            oRecordSet.DoQuery(strQuery)
            If Not oRecordSet.EoF Then
                Dim oPostingDate As DateTime = oApplication.Utilities.GetDateTimeValue(CType(oForm.Items.Item("24").Specific, SAPbouiCOM.EditText).Value)
                CType(oForm.Items.Item("26").Specific, SAPbouiCOM.EditText).Value = oPostingDate.AddMonths(oRecordSet.Fields.Item(0).Value).AddDays(oRecordSet.Fields.Item(1).Value).ToString("yyyyMMdd")
            End If
            oForm.Freeze(False)
        Catch ex As Exception
            oForm.Freeze(False)
            Throw ex
        End Try
    End Sub

    Private Function getUserSeries(ByVal oForm As SAPbouiCOM.Form) As Integer
        Dim _retVal As Integer = 0
        Try
            Dim oCmpSrv As SAPbobsCOM.CompanyService
            Dim oSeriesService As SAPbobsCOM.SeriesService
            Dim oSeries As SAPbobsCOM.Series
            Dim oDocumentTypeParams As SAPbobsCOM.DocumentTypeParams
            oCmpSrv = oApplication.Company.GetCompanyService()
            oSeriesService = oCmpSrv.GetBusinessService(SAPbobsCOM.ServiceTypes.SeriesService)
            oSeries = oSeriesService.GetDataInterface(SAPbobsCOM.SeriesServiceDataInterfaces.ssdiSeries)
            oDocumentTypeParams = oSeriesService.GetDataInterface(SAPbobsCOM.SeriesServiceDataInterfaces.ssdiDocumentTypeParams)
            oDocumentTypeParams.Document = 23
            oSeries = oSeriesService.GetDefaultSeries(oDocumentTypeParams)
            _retVal = oSeries.Series
            Return _retVal
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Private Sub clearControls(ByVal oForm As SAPbouiCOM.Form)
        Try

        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub loadData(ByVal oForm As SAPbouiCOM.Form)
        Try
            oForm.Freeze(True)
            Dim strCondition As String

            Dim strCode As String
            Dim oEdit As SAPbouiCOM.EditText
            oEdit = oForm.Items.Item("81").Specific
            strCode = oEdit.String
            ' oComboBox = oForm.Items.Item("77").Specific
            ' strCondition = oComboBox.Selected.Value
            ' oDt_Import = oForm.DataSources.DataTables.Item("Dt_Import")
            strQuery = " Select * from [Z_OBDG]"
            strQuery += " where (SUCCESS = '1' and RefCode='" & strCode & "') order by convert(Numeric,RowID)"
            'If strCondition = "O" Then
            '    strQuery += " and ((Convert(Decimal(18,2),T0.OnHand)) > = (Convert(Decimal(18,2),QTY)))"
            'ElseIf strCondition = "A" Then
            '    strQuery += "  and ( ((Convert(Decimal(18,2),T0.OnHand)) > = (Convert(Decimal(18,2),QTY))) or (Convert(Decimal(18,2),AQTY)) >  (Convert(Decimal(18,2),QTY)))"
            'ElseIf strCondition = "P" Then

            '    strQuery += "and( (Convert(Decimal(18,2),T0.OnHand)) > = (Convert(Decimal(18,2),QTY)) or ( (Convert(Decimal(18,2),QTY)) > (Convert(Decimal(18,2),T0.OnHand)) and (Convert(Decimal(18,2),T0.OnHand))>0)) "
            'Else

            '    strQuery += " and 1=1"
            'End If
            ' oDt_Import.ExecuteQuery(strQuery)
            oGrid = oForm.Items.Item("68").Specific
            oGrid.DataTable.ExecuteQuery(strQuery) ' = oDt_Import

            '   oDt_ErrorLog = oForm.DataSources.DataTables.Item("Dt_ErrorLog")
            strQuery = "Select RowID,Year,AcctCode,AcctName,OcrCode,OcrName,ERROR As 'Error' From  Z_OBDG T0 Where SUCCESS = '0' and RefCode='" & strCode & "'"
            '  oDt_ErrorLog.ExecuteQuery(strQuery)
            oGrid = oForm.Items.Item("69").Specific
            oGrid.DataTable.ExecuteQuery(strQuery) ' = oDt_ErrorLog

            '  formatGrid(oForm)
            oGrid = oForm.Items.Item("68").Specific
            'oEditColumn = oGrid.Columns.Item("Warehouse")
            'oEditColumn.ChooseFromListUID = "CFL_18"
            'oEditColumn.ChooseFromListAlias = "WhsCode"
            'oEditColumn.LinkedObjectType = "64"
            oForm.Freeze(False)
        Catch ex As Exception

            oForm.Freeze(False)
            Throw ex
        End Try
    End Sub
    Private Sub UpdateRecords(ByVal oForm As SAPbouiCOM.Form)
        oGrid = oForm.Items.Item("68").Specific
        Dim strQuery As String
        Dim Otemp As SAPbobsCOM.Recordset
        Otemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

        Dim strCode As String
        Dim oEdit As SAPbouiCOM.EditText
        oEdit = oForm.Items.Item("81").Specific
        strCode = oEdit.String
        For intRow As Integer = 0 To oGrid.DataTable.Rows.Count - 1
            strQuery = "Update Z_SIIM1 set Price='" & oGrid.DataTable.GetValue("Price", intRow) & "', Qty='" & oGrid.DataTable.GetValue("QTY", intRow) & "' ,Flag='Y' where RefCode='" & strCode & "' and  RowID='" & oGrid.DataTable.GetValue("RowID", intRow) & "'"
            Otemp.DoQuery(strQuery)
        Next
    End Sub
    Private Sub formatGrid(ByVal oForm As SAPbouiCOM.Form)
        Try
            oForm.Freeze(True)
            oGrid = oForm.Items.Item("68").Specific
            formatAll(oForm, oGrid)
            oGrid.RowHeaders.SetText(0, "#")
            For intIndex As Int16 = 0 To oGrid.Rows.Count - 1
                Try
                    Dim dblQty As Double = oGrid.DataTable.GetValue("QTY", intIndex)
                    Dim onhand As Double = oGrid.DataTable.GetValue("OnHand", intIndex)

                    If onhand < dblQty Then
                        ' oGrid.CommonSetting.SetRowBackColor(intIndex + 1, RGB(204, 255, 255))
                        oGrid.CommonSetting.SetRowFontColor(intIndex + 1, RGB(255, 0, 0))
                    Else
                        oGrid.CommonSetting.SetRowFontColor(intIndex + 1, RGB(0, 0, 0))

                    End If
                Catch ex As Exception

                End Try
                oGrid.RowHeaders.SetText(intIndex, intIndex + 1)
            Next
            '  oGrid.CollapseLevel = 1

            oGrid = oForm.Items.Item("69").Specific
            oGrid.RowHeaders.SetText(0, "#")
            For intIndex As Int16 = 0 To oGrid.Rows.Count - 1


                oGrid.RowHeaders.SetText(intIndex, intIndex + 1)
            Next
            oGrid.Columns.Item("CODE").Type = SAPbouiCOM.BoGridColumnType.gct_EditText
            oEditColumn = oGrid.Columns.Item("CODE")
            oEditColumn.LinkedObjectType = "4"
            '  oGrid.CollapseLevel = 1
            oForm.Freeze(False)
        Catch ex As Exception
            oForm.Freeze(False)
            Throw ex
        End Try
    End Sub

    Private Sub formatAll(ByVal oForm As SAPbouiCOM.Form, ByVal oGrid As SAPbouiCOM.Grid)
        Try
            oForm.Items.Item("68").Enabled = True
            oGrid.Columns.Item("CODE").TitleObject.Caption = "Item Code"
            oGrid.Columns.Item("DESCRIPTION").TitleObject.Caption = "Description"
            oGrid.Columns.Item("QTY").TitleObject.Caption = "Quantity Required"
            oGrid.Columns.Item("OnHand").TitleObject.Caption = "On Hand"
            oGrid.Columns.Item("AQTY").TitleObject.Caption = "Available Quantity"
            oGrid.Columns.Item("CODE").Type = SAPbouiCOM.BoGridColumnType.gct_EditText
            oEditColumn = oGrid.Columns.Item("CODE")
            oEditColumn.LinkedObjectType = "4"

            oGrid.Columns.Item("QTY").RightJustified = True
            oGrid.Columns.Item("QTY").Editable = True
            oGrid.Columns.Item("Price").Editable = True
            oGrid.Columns.Item("AQTY").RightJustified = True
            oGrid.Columns.Item("OnHand").RightJustified = True

            oGrid.Columns.Item("CODE").Editable = False
            oGrid.Columns.Item("DESCRIPTION").Editable = False
            oGrid.Columns.Item("OnHand").Editable = False
            oGrid.Columns.Item("AQTY").Editable = False
            oEditColumn = oGrid.Columns.Item("AQTY")
            oEditColumn.LinkedObjectType = "64"
            oGrid.Columns.Item("RowID").Editable = False



        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub changeFormPane(ByVal oForm As SAPbouiCOM.Form, ByVal btnID As String)
        Try
            Select Case btnID
                Case "63"
                    oForm.PaneLevel = oForm.PaneLevel + 1
                Case "65"
                    oForm.PaneLevel = oForm.PaneLevel - 1
                Case "66"
                    oForm.PaneLevel = 3
                Case "67"
                    oForm.PaneLevel = 4
                Case "71"
                    oForm.PaneLevel = oForm.PaneLevel + 1
                Case "70"
                    oForm.PaneLevel = oForm.PaneLevel - 1
                Case Else
            End Select
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

#End Region

End Class
