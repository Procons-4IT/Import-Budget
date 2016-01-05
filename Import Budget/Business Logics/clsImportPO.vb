Imports System.IO
Imports System.Text.RegularExpressions
Imports System.Text
Imports System
Imports Microsoft.Office.Interop.Excel
Public Class clsImportPO
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
            oForm = oApplication.Utilities.LoadForm(xml_POImp, frm_POImp)
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


    Private Function ReadExcelData_New(ByVal aForm As SAPbouiCOM.Form, ByVal strID As String) As Boolean
        Dim strWhs As String = CType(aForm.Items.Item("30").Specific, SAPbouiCOM.EditText).Value
        Dim strPath As String = CType(aForm.Items.Item(strID).Specific, SAPbouiCOM.StaticText).Caption
        Dim strItemCode As String
        Dim strCardCode As String = CType(aForm.Items.Item("3").Specific, SAPbouiCOM.EditText).Value
        Dim excel As Application = New Application
        ' Open Excel spreadsheet.
        Dim w As Workbook = excel.Workbooks.Open(strPath)
        Dim dblAmt As Single
        Try
            Dim ostatic As SAPbouiCOM.StaticText
            ostatic = aForm.Items.Item("78").Specific

            Dim oTest As SAPbobsCOM.Recordset
            oTest = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Try
                ' oTest.DoQuery("Delete from [Z_OPOR]")
            Catch ex As Exception

            End Try
            Dim strCode As String = oApplication.Utilities.getMaxCode("@S_PImport", "Code")
            Dim s1 As String = "Insert Into [@S_PImport] values('" & strCode & "','" & strCode & "')"
            oTest.DoQuery(s1)

            Dim oEdit As SAPbouiCOM.EditText
            oEdit = aForm.Items.Item("81").Specific
            oEdit.String = strCode
            oTest.DoQuery("Delete From [Z_OPOR] where RefCode='" & strCode & "'")
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
                '   oApplication.Utilities.createDatatable()
                ' Scan the cells.
                If array IsNot Nothing Then
                    Dim bound0 As Integer = array.GetUpperBound(0)
                    Dim bound1 As Integer = array.GetUpperBound(1)
                    ' Loop over all elements.
                    _oDtPO.Rows.Clear()

                    For j As Integer = 1 To bound0
                        ostatic.Caption = "Reading Excel Data " & j & " of " & bound0
                        Try
                            If j > 2 And array(j, 1).ToString <> "" Then
                                oApplication.Utilities.ReadExcelDat_PO(array(j, 1), array(j, 2), array(j, 3), array(j, 4), aForm, j - 1, strrefcode)
                            End If
                        Catch ex As Exception

                        End Try
                    Next
                End If
            Next
            ostatic.Caption = "Readin Excel Data Completed"
            w.Close()
            excel.Quit()
            If Not IsNothing(_oDtPO) Then
                If _oDtPO.Rows.Count > 0 Then
                    ' Dim _unPivot As DataTable = unPivotTable(oForm, _oDt)
                    Dim strDtXML As String = oApplication.Utilities.getXMLstring(_oDtPO) ' _unPivot)
                    Dim s As String = "Exec [Insert_POImport] '" + strDtXML + "'"
                    oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    oRecordSet.DoQuery("Exec [Insert_POImport] '" + strDtXML + "'")
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
#Region "Menu Event"
    Public Overrides Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        Try
            oForm = oApplication.SBO_Application.Forms.ActiveForm
            Select Case pVal.MenuUID
                Case mnu_POImp
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
            If pVal.FormTypeEx = frm_POImp Then
                Select Case pVal.BeforeAction
                    Case True
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "57" Then 'Browse
                                    oApplication.Utilities.OpenFileDialogBox(oForm, "56")
                                ElseIf (pVal.ItemUID = "58") Then 'Next
                                    oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                    If CType(oForm.Items.Item("17").Specific, SAPbouiCOM.StaticText).Caption <> "" Then
                                        If oApplication.Utilities.ValidateFile(oForm, "56") Then
                                            If ReadExcelData_New(oForm, "56") Then ' oApplication.Utilities.GetExcelData(oForm, "56") Then
                                                oForm.Items.Item("71").Enabled = True
                                                oApplication.Utilities.Message("PO Data Imported Successfully....", SAPbouiCOM.BoStatusBarMessageType.smt_Success)
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
                                    Dim _retVal As Integer = oApplication.SBO_Application.MessageBox("Proceed with creating a PO ?", 2, "Yes", "No", "")
                                    If _retVal = 1 Then
                                        Dim oRec As SAPbobsCOM.Recordset
                                        oRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                                        oRec.DoQuery("Select * from Z_OPOR where ERROR='Item Code Not Found'")
                                        If oRec.RecordCount > 0 Then
                                            _retVal = oApplication.SBO_Application.MessageBox("Some of the Item Codes are not available. Do you want to Create Missing Items ?", 2, "Yes", "No", "")
                                            If _retVal = 2 Then
                                                Exit Sub
                                            Else
                                                If oApplication.Utilities.CreateItemCode(oForm) = False Then
                                                    Exit Sub
                                                Else
                                                    loadData(oForm)
                                                End If
                                            End If
                                        End If

                                        UpdateRecords(oForm)
                                        If oApplication.Utilities.saveAsDraft_PO(oForm) Then
                                            oApplication.Utilities.Message("PO  Document Created Sucessfully....", SAPbouiCOM.BoStatusBarMessageType.smt_Success)
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
            oForm.PaneLevel = 1
            CType(oForm.Items.Item("24").Specific, SAPbouiCOM.EditText).Value = System.DateTime.Now.ToString("yyyyMMdd")
            CType(oForm.Items.Item("26").Specific, SAPbouiCOM.EditText).Value = System.DateTime.Now.ToString("yyyyMMdd")
            CType(oForm.Items.Item("28").Specific, SAPbouiCOM.EditText).Value = System.DateTime.Now.ToString("yyyyMMdd")
            oForm.DataSources.UserDataSources.Add("filter", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
            oForm.DataSources.UserDataSources.Add("RefNo", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
            oComboBox = oForm.Items.Item("77").Specific
            oComboBox.ValidValues.Add("0", "All Items")
            oComboBox.ValidValues.Add("O", "On Hand")
            oComboBox.ValidValues.Add("A", "Available Qty")
            oComboBox.ValidValues.Add("P", "Partial")
            oComboBox.Select(0, SAPbouiCOM.BoSearchKey.psk_Index)
            oComboBox.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly
            Dim oEdit As SAPbouiCOM.EditText
            oEdit = oForm.Items.Item("30").Specific
            oEdit.String = "01"
            oEdit = oForm.Items.Item("81").Specific
            oEdit.DataBind.SetBound(True, "", "RefNo")
            initializeDataTable(oForm)
            addStyle(oForm)
            loadCombo(oForm)
            cflFilter(oForm)

            'Select User Series
            oComboBox = oForm.Items.Item("20").Specific
            Dim strUserSeries As String = getUserSeries_GRPO(oForm).ToString()
            If strUserSeries.Length > 0 And strUserSeries <> "0" Then
                oComboBox.Select(strUserSeries, SAPbouiCOM.BoSearchKey.psk_ByValue)
            End If

            oForm.Items.Item("3").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
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

            ' formatGrid(oForm)
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
            strQuery = "Select Series,SeriesName From NNM1 Where ObjectCode = '22'"
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
            oCon.CondVal = "S"
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

    Private Function validate(ByVal oForm As SAPbouiCOM.Form) As Boolean
        Dim _retVal As Boolean = True
        Try
            If CType(oForm.Items.Item("3").Specific, SAPbouiCOM.EditText).Value = "" Then 'Card Code
                oApplication.Utilities.Message("Select Customer....", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                _retVal = False
            ElseIf CType(oForm.Items.Item("5").Specific, SAPbouiCOM.EditText).Value = "" Then ' CardName
                ' oApplication.Utilities.Message("Customer Name Cannot be Blank....", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                '  _retVal = False
            ElseIf CType(oForm.Items.Item("24").Specific, SAPbouiCOM.EditText).Value = "" Then ' DocDate
                oApplication.Utilities.Message("Select Document Date....", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                _retVal = False
            ElseIf CType(oForm.Items.Item("26").Specific, SAPbouiCOM.EditText).Value = "" Then ' Due Date
                oApplication.Utilities.Message("Select Document Due Date....", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                _retVal = False
            ElseIf CType(oForm.Items.Item("28").Specific, SAPbouiCOM.EditText).Value = "" Then ' Tax Date
                oApplication.Utilities.Message("Select Posting Date....", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                _retVal = False
            ElseIf CType(oForm.Items.Item("30").Specific, SAPbouiCOM.EditText).Value = "" Then ' WareHouse
                oApplication.Utilities.Message("Select WareHouse....", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                _retVal = False
            ElseIf (CType(oForm.Items.Item("30").Specific, SAPbouiCOM.EditText).Value = "") Then
                'oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                'strQuery = " Select WhsCode From OWHS Where BinActivat = 'Y' And WhsCode = '" + CType(oForm.Items.Item("30").Specific, SAPbouiCOM.EditText).Value + "'"
                'oRecordSet.DoQuery(strQuery)
                If 1 = 1 Then
                    '   oApplication.Utilities.Message("Selected WareHouse Should be Bin Managed....", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    ' _retVal = False
                End If
                'ElseIf CType(oForm.Items.Item("9").Specific, SAPbouiCOM.EditText).Value = "" Then ' Num At Card
                '    oApplication.Utilities.Message("Enter Reference....", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                '    _retVal = False
                'ElseIf CType(oForm.Items.Item("49").Specific, SAPbouiCOM.EditText).Value = "" Then ' Control Acc
                '    oApplication.Utilities.Message("Select Control Account....", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                '    _retVal = False
            ElseIf CType(oForm.Items.Item("46").Specific, SAPbouiCOM.ComboBox).Selected.Value = "" Then ' Tax Code
                '  oApplication.Utilities.Message("Select Tax Code....", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                '  _retVal = False
            ElseIf CType(oForm.Items.Item("20").Specific, SAPbouiCOM.ComboBox).Selected.Value = "" Then ' Sale Employee
                oApplication.Utilities.Message("Select Document Series....", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                _retVal = False
                'ElseIf CType(oForm.Items.Item("11").Specific, SAPbouiCOM.ComboBox).Selected.Value = "" Then ' Sale Employee
                '    oApplication.Utilities.Message("Select Document Currency....", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                '    _retVal = False
                'ElseIf CType(oForm.Items.Item("16").Specific, SAPbouiCOM.ComboBox).Selected.Value = "" Then ' Sale Employee
                '    oApplication.Utilities.Message("Select Sales Employee....", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                '    _retVal = False
                'ElseIf CType(oForm.Items.Item("33").Specific, SAPbouiCOM.ComboBox).Selected.Value = "" Then ' Ship To
                '    oApplication.Utilities.Message("Select Shipping Type....", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                '    _retVal = False
                'ElseIf CType(oForm.Items.Item("38").Specific, SAPbouiCOM.ComboBox).Selected.Value = "" Then ' Bill To
                '    oApplication.Utilities.Message("Select Bill To Address....", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                '    _retVal = False
                'ElseIf CType(oForm.Items.Item("36").Specific, SAPbouiCOM.ComboBox).Selected.Value = "" Then ' Shipping Type
                '    oApplication.Utilities.Message("Select Ship To Address....", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                '    _retVal = False
                'ElseIf CType(oForm.Items.Item("7").Specific, SAPbouiCOM.ComboBox).Selected.Value = "" Then ' Contact Persion
                '    oApplication.Utilities.Message("Select Contact Person....", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                '    _retVal = False
                'ElseIf CType(oForm.Items.Item("44").Specific, SAPbouiCOM.ComboBox).Selected.Value = "" Then ' Payment Terms
                '    oApplication.Utilities.Message("Select Payment Terms....", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                '    _retVal = False
                'ElseIf CType(oForm.Items.Item("51").Specific, SAPbouiCOM.ComboBox).Selected.Value = "" Then ' Payment Methond
                '    oApplication.Utilities.Message("Select Payment Method....", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                '    _retVal = False
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

    Private Function getUserSeries_GRPO(ByVal oForm As SAPbouiCOM.Form) As Integer
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
            oDocumentTypeParams.Document = 22
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
            oComboBox = oForm.Items.Item("77").Specific
            strCondition = oComboBox.Selected.Value
            Dim strCode As String
            Dim oEdit As SAPbouiCOM.EditText
            oEdit = oForm.Items.Item("81").Specific
            strCode = oEdit.String
            oDt_Import = oForm.DataSources.DataTables.Item("Dt_Import")
            strQuery = " Select RowID,CODE,T1.[ItemName] As [DESCRIPTION], (Convert(Decimal(18,2),QTY)) as 'QTY',(Convert(Decimal(18,2),T0.Price)) 'Price' "
            strQuery += "  From Z_OPOR T0 "
            strQuery += " JOIN OITM T1 On T1.ItemCode = T0.CODE "
            'strQuery += " JOIN OBIN T2 On T2.SL1Code = T0.BIN "
            ' strQuery += " Group By Code,T0.[CODE],SUCCESS "
            ' strQuery += " Having Sum(Convert(Decimal(18,2),Qty))>0' <= Sum(Convert(Decimal(18,2),AQty)) "
            'strQuery += " AND Sum(Convert(Decimal(18,2),Qty)) > 0 "
            strQuery += " where SUCCESS = '1' and RefCode='" & strCode & "'"
            'If strCondition = "O" Then
            '    strQuery += " and (Convert(Decimal(18,2),T0.OnHand)) > = (Convert(Decimal(18,2),QTY))"
            'ElseIf strCondition = "A" Then
            '    strQuery += "  and ((Convert(Decimal(18,2),T0.OnHand)) > = (Convert(Decimal(18,2),QTY))) or (Convert(Decimal(18,2),AQTY)) >  (Convert(Decimal(18,2),QTY))"
            'ElseIf strCondition = "P" Then

            '    strQuery += "and (Convert(Decimal(18,2),T0.OnHand)) > = (Convert(Decimal(18,2),QTY)) or ( (Convert(Decimal(18,2),QTY)) > (Convert(Decimal(18,2),T0.OnHand)) and (Convert(Decimal(18,2),T0.OnHand))>0) "
            'Else

            '    strQuery += " and 1=1"
            'End If
            oDt_Import.ExecuteQuery(strQuery)
            oGrid = oForm.Items.Item("68").Specific
            oGrid.DataTable = oDt_Import

            oDt_ErrorLog = oForm.DataSources.DataTables.Item("Dt_ErrorLog")
            strQuery = "Select RowID,CODE,T0.[DESCRIPTION] As [DESCRIPTION],ERROR As 'Error' From  Z_OPOR T0 left outer JOIN OITM T1 ON T0.CODE = T1.ItemCODE Where SUCCESS = '0' and RefCode='" & strCode & "'"
            oDt_ErrorLog.ExecuteQuery(strQuery)
            oGrid = oForm.Items.Item("69").Specific
            oGrid.DataTable = oDt_ErrorLog

            formatGrid(oForm)
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
        For intRow As Integer = 0 To oGrid.DataTable.Rows.Count - 1
            strQuery = "Update Z_OPOR set Flag='Y' where RowID='" & oGrid.DataTable.GetValue("RowID", intRow) & "'"
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
                'Try
                '    Dim dblQty As Double = oGrid.DataTable.GetValue("QTY", intIndex)
                '    Dim onhand As Double = 0 'oGrid.DataTable.GetValue("OnHand", intIndex)

                '    If onhand < dblQty Then
                '        ' oGrid.CommonSetting.SetRowBackColor(intIndex + 1, RGB(204, 255, 255))
                '        oGrid.CommonSetting.SetRowFontColor(intIndex + 1, RGB(255, 0, 0))
                '    Else
                '        oGrid.CommonSetting.SetRowFontColor(intIndex + 1, RGB(0, 0, 0))

                '    End If
                'Catch ex As Exception

                'End Try
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
            ' oGrid.Columns.Item("ORDNO").TitleObject.Caption = "Order ID"
            oGrid.Columns.Item("CODE").TitleObject.Caption = "Item Code"
            oGrid.Columns.Item("DESCRIPTION").TitleObject.Caption = "Description"
            oGrid.Columns.Item("QTY").TitleObject.Caption = "Quantity Required"
            '  oGrid.Columns.Item("Total").TitleObject.Caption = "Total"
            oGrid.Columns.Item("CODE").Type = SAPbouiCOM.BoGridColumnType.gct_EditText
            oEditColumn = oGrid.Columns.Item("CODE")
            oEditColumn.LinkedObjectType = "4"

            oGrid.Columns.Item("QTY").RightJustified = True
            oGrid.Columns.Item("QTY").Editable = False
            oGrid.Columns.Item("Price").Editable = False

            oGrid.Columns.Item("CODE").Editable = False
            oGrid.Columns.Item("DESCRIPTION").Editable = False
            oGrid.Columns.Item("RowID").Editable = False
            Try
                '  oGrid.Columns.Item("BaseRef").TitleObject.Caption = "Base Ref"
                '  oGrid.Columns.Item("BaseLine").TitleObject.Caption = "Base Line"
            Catch ex As Exception

            End Try

            ' oGrid.Columns.Item("BaseRef").Editable = False
            '  oGrid.Columns.Item("Total").Editable = False
            '   oGrid.Columns.Item("ORDNO").Editable = False
            '    oGrid.Columns.Item("BaseLine").Editable = False
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
