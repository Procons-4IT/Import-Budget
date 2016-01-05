Imports System.Xml
Imports System.Collections.Specialized
Imports System.IO


Public Class clsUtilities

    Private strThousSep As String = ","
    Private strDecSep As String = "."
    Private intQtyDec As Integer = 3
    Private FormNum As Integer
    Private strFilepath As String = String.Empty
    Private oRecordSet As SAPbobsCOM.Recordset
    Private sQuery As String = String.Empty
    Dim oform As SAPbouiCOM.Form
    Dim oGrid As SAPbouiCOM.Grid
    Dim oEditColumn As SAPbouiCOM.EditTextColumn

    Public Sub New()
        MyBase.New()
        FormNum = 1
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
            oCFL = oCFLs.Item("CFL_2")
            oCons = oCFL.GetConditions()
            oCon = oCons.Add()
            oCon.Alias = "CardType"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCon.CondVal = "C"
            oCFL.SetConditions(oCons)

         
           

        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Public Sub LoadPriceHistory(ByVal aCode As String, ByVal aItemCode As String)
        Try
            Dim oTest As SAPbobsCOM.Recordset
            oTest = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oform = LoadForm("frm_ItemPriceHisttory.xml", "frm_ItemPrice")
            oform = oApplication.SBO_Application.Forms.ActiveForm()
            oform.Freeze(True)
            cflFilter(oform)
            oform.DataSources.UserDataSources.Add("ItemName", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
            oform.DataSources.UserDataSources.Add("ItemCode", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
            oform.DataSources.UserDataSources.Add("SellPri", SAPbouiCOM.BoDataType.dt_PRICE)
            oform.DataSources.UserDataSources.Add("Stock", SAPbouiCOM.BoDataType.dt_QUANTITY)
            Dim oEdit As SAPbouiCOM.EditText
            oEdit = oform.Items.Item("3").Specific
            oEdit.DataBind.SetBound(True, "", "ItemName")
            oEdit.String = aItemCode
            oEdit = oform.Items.Item("5").Specific
            oEdit.DataBind.SetBound(True, "", "ItemCode")
            oEdit.String = aCode
            oEdit = oform.Items.Item("7").Specific
            oEdit.DataBind.SetBound(True, "", "SellPri")
            oEdit = oform.Items.Item("9").Specific
            oEdit.DataBind.SetBound(True, "", "Stock")
            oTest.DoQuery("Select Sum(OnHand) from OITW where ItemCode='" & aCode & "'")
            oEdit.String = oTest.Fields.Item(0).Value

            oTest.DoQuery("Select Price from ITM1 where pricelist=1 and  ItemCode='" & aCode & "'")
            oEdit = oform.Items.Item("7").Specific
            oEdit.String = oTest.Fields.Item(0).Value


            Dim strQry As String

            strQry = " SELECT Top 20 CardName 'Customer Name',T0.DocDate 'Document Date',Quantity,Price FROM OINV T0  INNER JOIN INV1 T1 ON T0.[DocEntry] = T1.[DocEntry] where  T1.ItemCode='" & aCode & "' order by T0.DocEntry Desc"
            oGrid = oform.Items.Item("10").Specific
            oGrid.DataTable.ExecuteQuery(strQry)
            oform.Items.Item("10").Enabled = False
            For intRow As Integer = 0 To oGrid.DataTable.Rows.Count - 1
                oGrid.RowHeaders.SetText(intRow, intRow + 1)
            Next
            oGrid.AutoResizeColumns()
            oform.PaneLevel = 1
            oform.Freeze(False)
        Catch ex As Exception
            oform.Freeze(False)
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Public Sub RefreshItemPriceHistory(ByVal aCode As String, ByVal aItemCode As String)
        Try
            Dim oTest As SAPbobsCOM.Recordset
            oTest = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Dim strQry As String
            If aCode = "" Then
                strQry = " SELECT Top 20 CardName 'Customer Name',T0.DocDate 'Document Date',Quantity,Price FROM OINV T0  INNER JOIN INV1 T1 ON T0.[DocEntry] = T1.[DocEntry] where   T1.ItemCode='" & aItemCode & "' order by T0.DocEntry Desc"
            Else
                strQry = " SELECT Top 20 CardName 'Customer Name',T0.DocDate 'Document Date',Quantity,Price FROM OINV T0  INNER JOIN INV1 T1 ON T0.[DocEntry] = T1.[DocEntry] where T0.CardCode='" & aCode & "' and   T1.ItemCode='" & aItemCode & "' order by T0.DocEntry Desc"
            End If

            oGrid = oform.Items.Item("10").Specific
            oGrid.DataTable.ExecuteQuery(strQry)
            oform.Items.Item("10").Enabled = False
            For intRow As Integer = 0 To oGrid.DataTable.Rows.Count - 1
                oGrid.RowHeaders.SetText(intRow, intRow + 1)
            Next
            oGrid.AutoResizeColumns()
            oform.PaneLevel = 1
            oform.Freeze(False)
        Catch ex As Exception
            oform.Freeze(False)
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

#Region "Connect to Company"
    Public Sub Connect()
        Dim strCookie As String
        Dim strConnectionContext As String

        Try
            strCookie = oApplication.Company.GetContextCookie
            strConnectionContext = oApplication.SBO_Application.Company.GetConnectionContext(strCookie)

            If oApplication.Company.SetSboLoginContext(strConnectionContext) <> 0 Then
                Throw New Exception("Wrong login credentials.")
            End If

            'Open a connection to company
            If oApplication.Company.Connect() <> 0 Then
                Throw New Exception("Cannot connect to company database. ")
            End If

        Catch ex As Exception
            Throw ex
        End Try
    End Sub
#End Region

#Region "Genral Functions"

#Region "Get MaxCode"
    Public Function getMaxCode(ByVal sTable As String, ByVal sColumn As String) As String
        Dim oRS As SAPbobsCOM.Recordset
        Dim MaxCode As Integer
        Dim sCode As String
        Dim strSQL As String
        Try
            strSQL = "SELECT MAX(CAST(" & sColumn & " AS Numeric)) FROM [" & sTable & "]"
            ExecuteSQL(oRS, strSQL)

            If Convert.ToString(oRS.Fields.Item(0).Value).Length > 0 Then
                MaxCode = oRS.Fields.Item(0).Value + 1
            Else
                MaxCode = 1
            End If

            sCode = Format(MaxCode, "00000000")
            Return sCode
        Catch ex As Exception
            Throw ex
        Finally
            oRS = Nothing
        End Try
    End Function
#End Region

#Region "Status Message"
    Public Sub Message(ByVal sMessage As String, ByVal StatusType As SAPbouiCOM.BoStatusBarMessageType)
        oApplication.SBO_Application.StatusBar.SetText(sMessage, SAPbouiCOM.BoMessageTime.bmt_Short, StatusType)
    End Sub
#End Region

#Region "Add Choose from List"
    Public Sub AddChooseFromList(ByVal FormUID As String, ByVal CFL_Text As String, ByVal CFL_Button As String, _
                                        ByVal ObjectType As SAPbouiCOM.BoLinkedObject, _
                                            Optional ByVal AliasName As String = "", Optional ByVal CondVal As String = "", _
                                                    Optional ByVal Operation As SAPbouiCOM.BoConditionOperation = SAPbouiCOM.BoConditionOperation.co_EQUAL)

        Dim oCFLs As SAPbouiCOM.ChooseFromListCollection
        Dim oCons As SAPbouiCOM.Conditions
        Dim oCon As SAPbouiCOM.Condition
        Dim oCFL As SAPbouiCOM.ChooseFromList
        Dim oCFLCreationParams As SAPbouiCOM.ChooseFromListCreationParams
        Try
            oCFLs = oApplication.SBO_Application.Forms.Item(FormUID).ChooseFromLists
            oCFLCreationParams = oApplication.SBO_Application.CreateObject( _
                                    SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)

            ' Adding 2 CFL, one for the button and one for the edit text.
            If ObjectType = SAPbouiCOM.BoLinkedObject.lf_Items Then
                oCFLCreationParams.MultiSelection = True
            Else
                oCFLCreationParams.MultiSelection = False
            End If

            oCFLCreationParams.ObjectType = ObjectType
            oCFLCreationParams.UniqueID = CFL_Text

            oCFL = oCFLs.Add(oCFLCreationParams)

            ' Adding Conditions to CFL1

            oCons = oCFL.GetConditions()

            If Not AliasName = "" Then
                oCon = oCons.Add()
                oCon.Alias = AliasName
                oCon.Operation = Operation
                oCon.CondVal = CondVal
                oCFL.SetConditions(oCons)
            End If

            oCFLCreationParams.UniqueID = CFL_Button
            oCFL = oCFLs.Add(oCFLCreationParams)

        Catch ex As Exception
            Throw ex
        End Try
    End Sub
#End Region

#Region "Get Linked Object Type"
    Public Function getLinkedObjectType(ByVal Type As SAPbouiCOM.BoLinkedObject) As String
        Return CType(Type, String)
    End Function

#End Region

#Region "Execute Query"
    Public Sub ExecuteSQL(ByRef oRecordSet As SAPbobsCOM.Recordset, ByVal SQL As String)
        Try
            If oRecordSet Is Nothing Then
                oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            End If

            oRecordSet.DoQuery(SQL)

        Catch ex As Exception
            Throw ex
        End Try
    End Sub
#End Region

#Region "Get Application path"
    Public Function getApplicationPath() As String

        Return Application.StartupPath.Trim

        'Return IO.Directory.GetParent(Application.StartupPath).ToString
    End Function
#End Region

#Region "Date Manipulation"

#Region "Convert SBO Date to System Date"
    '********************************************************************
    'Type		            :   Public Procedure     
    'Name               	:	ConvertStrToDate
    'Parameter          	:   ByVal oDate As String, ByVal strFormat As String
    'Return Value       	:	
    'Author             	:	Manu
    'Created Date       	:	07/12/05
    'Last Modified By	    :	
    'Modified Date        	:	
    'Purpose             	:	To convert Date according to current culture info
    '********************************************************************
    Public Function ConvertStrToDate(ByVal strDate As String, ByVal strFormat As String) As DateTime
        Try
            Dim oDate As DateTime
            Dim ci As New System.Globalization.CultureInfo("en-GB", False)
            Dim newCi As System.Globalization.CultureInfo = CType(ci.Clone(), System.Globalization.CultureInfo)

            System.Threading.Thread.CurrentThread.CurrentCulture = newCi
            oDate = oDate.ParseExact(strDate, strFormat, ci.DateTimeFormat)

            Return oDate
        Catch ex As Exception
            Throw ex
        End Try

    End Function
#End Region

#Region " Get SBO Date Format in String (ddmmyyyy)"
    '********************************************************************
    'Type		            :   Public Procedure     
    'Name               	:	StrSBODateFormat
    'Parameter          	:   none
    'Return Value       	:	
    'Author             	:	Manu
    'Created Date       	:	
    'Last Modified By	    :	
    'Modified Date        	:	
    'Purpose             	:	To get date Format(ddmmyy value) as applicable to SBO
    '********************************************************************
    Public Function StrSBODateFormat() As String
        Try
            Dim rsDate As SAPbobsCOM.Recordset
            Dim strsql As String, GetDateFormat As String
            Dim DateSep As Char

            strsql = "Select DateFormat,DateSep from OADM"
            rsDate = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            rsDate.DoQuery(strsql)
            DateSep = rsDate.Fields.Item(1).Value

            Select Case rsDate.Fields.Item(0).Value
                Case 0
                    GetDateFormat = "dd" & DateSep & "MM" & DateSep & "yy"
                Case 1
                    GetDateFormat = "dd" & DateSep & "MM" & DateSep & "yyyy"
                Case 2
                    GetDateFormat = "MM" & DateSep & "dd" & DateSep & "yy"
                Case 3
                    GetDateFormat = "MM" & DateSep & "dd" & DateSep & "yyyy"
                Case 4
                    GetDateFormat = "yyyy" & DateSep & "dd" & DateSep & "MM"
                Case 5
                    GetDateFormat = "dd" & DateSep & "MMM" & DateSep & "yyyy"
            End Select
            Return GetDateFormat

        Catch ex As Exception
            Throw ex
        End Try
    End Function
#End Region

#Region "Get SBO date Format in Number"
    '********************************************************************
    'Type		            :   Public Procedure     
    'Name               	:	IntSBODateFormat
    'Parameter          	:   none
    'Return Value       	:	
    'Author             	:	Manu
    'Created Date       	:	
    'Last Modified By	    :	
    'Modified Date        	:	
    'Purpose             	:	To get date Format(integer value) as applicable to SBO
    '********************************************************************
    Public Function NumSBODateFormat() As String
        Try
            Dim rsDate As SAPbobsCOM.Recordset
            Dim strsql As String
            Dim DateSep As Char

            strsql = "Select DateFormat,DateSep from OADM"
            rsDate = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            rsDate.DoQuery(strsql)
            DateSep = rsDate.Fields.Item(1).Value

            Select Case rsDate.Fields.Item(0).Value
                Case 0
                    NumSBODateFormat = 3
                Case 1
                    NumSBODateFormat = 103
                Case 2
                    NumSBODateFormat = 1
                Case 3
                    NumSBODateFormat = 120
                Case 4
                    NumSBODateFormat = 126
                Case 5
                    NumSBODateFormat = 130
            End Select
            Return NumSBODateFormat

        Catch ex As Exception
            Throw ex
        End Try
    End Function
#End Region

#End Region

#Region "Get Rental Period"
    Public Function getRentalDays(ByVal Date1 As String, ByVal Date2 As String, ByVal IsWeekDaysBilling As Boolean) As Integer
        Dim TotalDays, TotalDaysincSat, TotalBillableDays As Integer
        Dim TotalWeekEnds As Integer
        Dim StartDate As Date
        Dim EndDate As Date
        Dim oRecordset As SAPbobsCOM.Recordset

        StartDate = CType(Date1.Insert(4, "/").Insert(7, "/"), Date)
        EndDate = CType(Date2.Insert(4, "/").Insert(7, "/"), Date)

        TotalDays = DateDiff(DateInterval.Day, StartDate, EndDate)

        If IsWeekDaysBilling Then
            strSQL = " select dbo.WeekDays('" & Date1 & "','" & Date2 & "')"
            oApplication.Utilities.ExecuteSQL(oRecordset, strSQL)
            If oRecordset.RecordCount > 0 Then
                TotalBillableDays = oRecordset.Fields.Item(0).Value
            End If
            Return TotalBillableDays
        Else
            Return TotalDays + 1
        End If

    End Function

    Public Function WorkDays(ByVal dtBegin As Date, ByVal dtEnd As Date) As Long
        Try
            Dim dtFirstSunday As Date
            Dim dtLastSaturday As Date
            Dim lngWorkDays As Long

            ' get first sunday in range
            dtFirstSunday = dtBegin.AddDays((8 - Weekday(dtBegin)) Mod 7)

            ' get last saturday in range
            dtLastSaturday = dtEnd.AddDays(-(Weekday(dtEnd) Mod 7))

            ' get work days between first sunday and last saturday
            lngWorkDays = (((DateDiff(DateInterval.Day, dtFirstSunday, dtLastSaturday)) + 1) / 7) * 5

            ' if first sunday is not begin date
            If dtFirstSunday <> dtBegin Then

                ' assume first sunday is after begin date
                ' add workdays from begin date to first sunday
                lngWorkDays = lngWorkDays + (7 - Weekday(dtBegin))

            End If

            ' if last saturday is not end date
            If dtLastSaturday <> dtEnd Then

                ' assume last saturday is before end date
                ' add workdays from last saturday to end date
                lngWorkDays = lngWorkDays + (Weekday(dtEnd) - 1)

            End If

            WorkDays = lngWorkDays
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try


    End Function

#End Region

#Region "Get Item Price with Factor"
    Public Function getPrcWithFactor(ByVal CardCode As String, ByVal ItemCode As String, ByVal RntlDays As Integer, ByVal Qty As Double) As Double
        Dim oItem As SAPbobsCOM.Items
        Dim Price, Expressn As Double
        Dim oDataSet, oRecSet As SAPbobsCOM.Recordset

        oItem = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oItems)
        oApplication.Utilities.ExecuteSQL(oDataSet, "Select U_RentFac, U_NumDys From [@REN_FACT] order by U_NumDys ")
        If oItem.GetByKey(ItemCode) And oDataSet.RecordCount > 0 Then

            oApplication.Utilities.ExecuteSQL(oRecSet, "Select ListNum from OCRD where CardCode = '" & CardCode & "'")
            oItem.PriceList.SetCurrentLine(oRecSet.Fields.Item(0).Value - 1)
            Price = oItem.PriceList.Price
            Expressn = 0
            oDataSet.MoveFirst()

            While RntlDays > 0

                If oDataSet.EoF Then
                    oDataSet.MoveLast()
                End If

                If RntlDays < oDataSet.Fields.Item(1).Value Then
                    Expressn += (oDataSet.Fields.Item(0).Value * RntlDays * Price * Qty)
                    RntlDays = 0
                    Exit While
                End If
                Expressn += (oDataSet.Fields.Item(0).Value * oDataSet.Fields.Item(1).Value * Price * Qty)
                RntlDays -= oDataSet.Fields.Item(1).Value
                oDataSet.MoveNext()

            End While

        End If
        If oItem.UserFields.Fields.Item("U_Rental").Value = "Y" Then
            Return CDbl(Expressn / Qty)
        Else
            Return Price
        End If


    End Function
#End Region

#Region "Get WareHouse List"
    Public Function getUsedWareHousesList(ByVal ItemCode As String, ByVal Quantity As Double) As DataTable
        Dim oDataTable As DataTable
        Dim oRow As DataRow
        Dim rswhs As SAPbobsCOM.Recordset
        Dim LeftQty As Double
        Try
            oDataTable = New System.Data.DataTable
            oDataTable.Columns.Add(New System.Data.DataColumn("ItemCode"))
            oDataTable.Columns.Add(New System.Data.DataColumn("WhsCode"))
            oDataTable.Columns.Add(New System.Data.DataColumn("Quantity"))

            strSQL = "Select WhsCode, ItemCode, (OnHand + OnOrder - IsCommited) As Available From OITW Where ItemCode = '" & ItemCode & "' And " & _
                        "WhsCode Not In (Select Whscode From OWHS Where U_Reserved = 'Y' Or U_Rental = 'Y') Order By (OnHand + OnOrder - IsCommited) Desc "

            ExecuteSQL(rswhs, strSQL)
            LeftQty = Quantity

            While Not rswhs.EoF
                oRow = oDataTable.NewRow()

                oRow.Item("WhsCode") = rswhs.Fields.Item("WhsCode").Value
                oRow.Item("ItemCode") = rswhs.Fields.Item("ItemCode").Value

                LeftQty = LeftQty - CType(rswhs.Fields.Item("Available").Value, Double)

                If LeftQty <= 0 Then
                    oRow.Item("Quantity") = CType(rswhs.Fields.Item("Available").Value, Double) + LeftQty
                    oDataTable.Rows.Add(oRow)
                    Exit While
                Else
                    oRow.Item("Quantity") = CType(rswhs.Fields.Item("Available").Value, Double)
                End If

                oDataTable.Rows.Add(oRow)
                rswhs.MoveNext()
                oRow = Nothing
            End While

            'strSQL = ""
            'For count As Integer = 0 To oDataTable.Rows.Count - 1
            '    strSQL += oDataTable.Rows(count).Item("WhsCode") & " : " & oDataTable.Rows(count).Item("Quantity") & vbNewLine
            'Next
            'MessageBox.Show(strSQL)

            Return oDataTable

        Catch ex As Exception
            Throw ex
        Finally
            oRow = Nothing
        End Try
    End Function
#End Region

#End Region

#Region "Functions related to Load XML"

#Region "Add/Remove Menus "
    Public Sub AddRemoveMenus(ByVal sFileName As String)
        Dim oXMLDoc As New Xml.XmlDocument
        Dim sFilePath As String
        Try
            sFilePath = getApplicationPath() & "\XML Files\" & sFileName
            oXMLDoc.Load(sFilePath)
            oApplication.SBO_Application.LoadBatchActions(oXMLDoc.InnerXml)
        Catch ex As Exception
            Throw ex
        Finally
            oXMLDoc = Nothing
        End Try
    End Sub
#End Region

#Region "Load XML File "
    Private Function LoadXMLFiles(ByVal sFileName As String) As String
        Dim oXmlDoc As Xml.XmlDocument
        Dim oXNode As Xml.XmlNode
        Dim oAttr As Xml.XmlAttribute
        Dim sPath As String
        Dim FrmUID As String
        Try
            oXmlDoc = New Xml.XmlDocument

            sPath = getApplicationPath() & "\XML Files\" & sFileName

            oXmlDoc.Load(sPath)
            oXNode = oXmlDoc.GetElementsByTagName("form").Item(0)
            oAttr = oXNode.Attributes.GetNamedItem("uid")
            oAttr.Value = oAttr.Value & FormNum
            FormNum = FormNum + 1
            oApplication.SBO_Application.LoadBatchActions(oXmlDoc.InnerXml)
            FrmUID = oAttr.Value

            Return FrmUID

        Catch ex As Exception
            Throw ex
        Finally
            oXmlDoc = Nothing
        End Try
    End Function
#End Region

    Public Function LoadForm(ByVal XMLFile As String, ByVal FormType As String) As SAPbouiCOM.Form
        'Return LoadForm(XMLFile, FormType.ToString(), FormType & "_" & oApplication.SBO_Application.Forms.Count.ToString)
        LoadXMLFiles(XMLFile)
        Return Nothing
    End Function

    '*****************************************************************
    'Type               : Function   
    'Name               : LoadForm
    'Parameter          : XmlFile,FormType,FormUID
    'Return Value       : SBO Form
    'Author             : Senthil Kumar B Senthil Kumar B
    'Created Date       : 
    'Last Modified By   : 
    'Modified Date      : 
    'Purpose            : To Load XML file 
    '*****************************************************************

    Public Function LoadForm(ByVal XMLFile As String, ByVal FormType As String, ByVal FormUID As String) As SAPbouiCOM.Form

        Dim oXML As System.Xml.XmlDocument
        Dim objFormCreationParams As SAPbouiCOM.FormCreationParams
        Try
            oXML = New System.Xml.XmlDocument
            oXML.Load(XMLFile)
            objFormCreationParams = (oApplication.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams))
            objFormCreationParams.XmlData = oXML.InnerXml
            objFormCreationParams.FormType = FormType
            objFormCreationParams.UniqueID = FormUID
            Return oApplication.SBO_Application.Forms.AddEx(objFormCreationParams)
        Catch ex As Exception
            Throw ex

        End Try

    End Function



#Region "Load Forms"
    Public Sub LoadForm(ByRef oObject As Object, ByVal XmlFile As String)
        Try
            oObject.FrmUID = LoadXMLFiles(XmlFile)
            oObject.Form = oApplication.SBO_Application.Forms.Item(oObject.FrmUID)
            If Not oApplication.Collection.ContainsKey(oObject.FrmUID) Then
                oApplication.Collection.Add(oObject.FrmUID, oObject)
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
#End Region

#End Region

#Region "Functions related to System Initilization"

#Region "Create Tables"
    Public Sub CreateTables()
        Dim oCreateTable As clsTable
        Try
            oCreateTable = New clsTable
            oCreateTable.CreateTables()
        Catch ex As Exception
            Throw ex
        Finally
            oCreateTable = Nothing
        End Try
    End Sub
#End Region

#Region "Notify Alert"
    Public Sub NotifyAlert()
        'Dim oAlert As clsPromptAlert

        'Try
        '    oAlert = New clsPromptAlert
        '    oAlert.AlertforEndingOrdr()
        'Catch ex As Exception
        '    Throw ex
        'Finally
        '    oAlert = Nothing
        'End Try

    End Sub
#End Region

#End Region

#Region "Function related to Quantities"

#Region "Get Available Quantity"
    Public Function getAvailableQty(ByVal ItemCode As String) As Long
        Dim rsQuantity As SAPbobsCOM.Recordset

        strSQL = "Select SUM(T1.OnHand + T1.OnOrder - T1.IsCommited) From OITW T1 Left Outer Join OWHS T3 On T3.Whscode = T1.WhsCode " & _
                    "Where T1.ItemCode = '" & ItemCode & "'"
        Me.ExecuteSQL(rsQuantity, strSQL)

        If rsQuantity.Fields.Item(0) Is System.DBNull.Value Then
            Return 0
        Else
            Return CLng(rsQuantity.Fields.Item(0).Value)
        End If

    End Function
#End Region

#Region "Get Rented Quantity"
    Public Function getRentedQty(ByVal ItemCode As String, ByVal StartDate As String, ByVal EndDate As String) As Long
        Dim rsQuantity As SAPbobsCOM.Recordset
        Dim RentedQty As Long

        strSQL = " select Sum(U_ReqdQty) from [@REN_RDR1] Where U_ItemCode = '" & ItemCode & "' " & _
                    " And DocEntry IN " & _
                    " (Select DocEntry from [@REN_ORDR] Where U_Status = 'R') " & _
                    " and '" & StartDate & "' between [@REN_RDR1].U_ShipDt1 and [@REN_RDR1].U_ShipDt2 "
        '" and [@REN_RDR1].U_ShipDt1 between '" & StartDate & "' and '" & EndDate & "'"

        ExecuteSQL(rsQuantity, strSQL)
        If Not rsQuantity.Fields.Item(0).Value Is System.DBNull.Value Then
            RentedQty = rsQuantity.Fields.Item(0).Value
        End If

        Return RentedQty

    End Function
#End Region

#Region "Get Reserved Quantity"
    Public Function getReservedQty(ByVal ItemCode As String, ByVal StartDate As String, ByVal EndDate As String) As Long
        Dim rsQuantity As SAPbobsCOM.Recordset
        Dim ReservedQty As Long

        strSQL = " select Sum(U_ReqdQty) from [@REN_QUT1] Where U_ItemCode = '" & ItemCode & "' " & _
                    " And DocEntry IN " & _
                    " (Select DocEntry from [@REN_OQUT] Where U_Status = 'R' And Status = 'O') " & _
                    " and '" & StartDate & "' between [@REN_QUT1].U_ShipDt1 and [@REN_QUT1].U_ShipDt2"

        ExecuteSQL(rsQuantity, strSQL)
        If Not rsQuantity.Fields.Item(0).Value Is System.DBNull.Value Then
            ReservedQty = rsQuantity.Fields.Item(0).Value
        End If

        Return ReservedQty

    End Function
#End Region

#End Region

#Region "Functions related to Tax"

#Region "Get Tax Codes"
    Public Sub getTaxCodes(ByRef oCombo As SAPbouiCOM.ComboBox)
        Dim rsTaxCodes As SAPbobsCOM.Recordset

        strSQL = "Select Code, Name From OVTG Where Category = 'O' Order By Name"
        Me.ExecuteSQL(rsTaxCodes, strSQL)

        oCombo.ValidValues.Add("", "")
        If rsTaxCodes.RecordCount > 0 Then
            While Not rsTaxCodes.EoF
                oCombo.ValidValues.Add(rsTaxCodes.Fields.Item(0).Value, rsTaxCodes.Fields.Item(1).Value)
                rsTaxCodes.MoveNext()
            End While
        End If
        oCombo.ValidValues.Add("Define New", "Define New")
        'oCombo.Select("")
    End Sub
#End Region

#Region "Get Applicable Code"

    Public Function getApplicableTaxCode1(ByVal CardCode As String, ByVal ItemCode As String, ByVal Shipto As String) As String
        Dim oBP As SAPbobsCOM.BusinessPartners
        Dim oItem As SAPbobsCOM.Items
        Dim rsExempt As SAPbobsCOM.Recordset
        Dim TaxGroup As String
        oBP = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oBusinessPartners)

        If oBP.GetByKey(CardCode.Trim) Then
            If oBP.VatLiable = SAPbobsCOM.BoVatStatus.vLiable Or oBP.VatLiable = SAPbobsCOM.BoVatStatus.vEC Then
                If oBP.VatGroup.Trim <> "" Then
                    TaxGroup = oBP.VatGroup.Trim
                Else
                    strSQL = "select LicTradNum from CRD1 where Address ='" & Shipto & "' and CardCode ='" & CardCode & "'"
                    Me.ExecuteSQL(rsExempt, strSQL)
                    If rsExempt.RecordCount > 0 Then
                        rsExempt.MoveFirst()
                        TaxGroup = rsExempt.Fields.Item(0).Value
                    Else
                        TaxGroup = ""
                    End If
                    'TaxGroup = oBP.FederalTaxID
                End If
            ElseIf oBP.VatLiable = SAPbobsCOM.BoVatStatus.vExempted Then
                strSQL = "Select Code From OVTG Where Rate = 0 And Category = 'O' Order By Code"
                Me.ExecuteSQL(rsExempt, strSQL)
                If rsExempt.RecordCount > 0 Then
                    rsExempt.MoveFirst()
                    TaxGroup = rsExempt.Fields.Item(0).Value
                Else
                    TaxGroup = ""
                End If
            End If
        End If




        Return TaxGroup

    End Function


    Public Function getApplicableTaxCode(ByVal CardCode As String, ByVal ItemCode As String) As String
        Dim oBP As SAPbobsCOM.BusinessPartners
        Dim oItem As SAPbobsCOM.Items
        Dim rsExempt As SAPbobsCOM.Recordset
        Dim TaxGroup As String
        oBP = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oBusinessPartners)

        If oBP.GetByKey(CardCode.Trim) Then
            If oBP.VatLiable = SAPbobsCOM.BoVatStatus.vLiable Or oBP.VatLiable = SAPbobsCOM.BoVatStatus.vEC Then
                If oBP.VatGroup.Trim <> "" Then
                    TaxGroup = oBP.VatGroup.Trim
                Else
                    TaxGroup = oBP.FederalTaxID
                End If
            ElseIf oBP.VatLiable = SAPbobsCOM.BoVatStatus.vExempted Then
                strSQL = "Select Code From OVTG Where Rate = 0 And Category = 'O' Order By Code"
                Me.ExecuteSQL(rsExempt, strSQL)
                If rsExempt.RecordCount > 0 Then
                    rsExempt.MoveFirst()
                    TaxGroup = rsExempt.Fields.Item(0).Value
                Else
                    TaxGroup = ""
                End If
            End If
        End If

        'If oBP.GetByKey(CardCode.Trim) Then
        '    If oBP.VatLiable = SAPbobsCOM.BoVatStatus.vLiable Or oBP.VatLiable = SAPbobsCOM.BoVatStatus.vEC Then
        '        If oBP.VatGroup.Trim <> "" Then
        '            TaxGroup = oBP.VatGroup.Trim
        '        Else
        '            oItem = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oItems)
        '            If oItem.GetByKey(ItemCode.Trim) Then
        '                TaxGroup = oItem.SalesVATGroup.Trim
        '            End If
        '        End If
        '    ElseIf oBP.VatLiable = SAPbobsCOM.BoVatStatus.vExempted Then
        '        strSQL = "Select Code From OVTG Where Rate = 0 And Category = 'O' Order By Code"
        '        Me.ExecuteSQL(rsExempt, strSQL)
        '        If rsExempt.RecordCount > 0 Then
        '            rsExempt.MoveFirst()
        '            TaxGroup = rsExempt.Fields.Item(0).Value
        '        Else
        '            TaxGroup = ""
        '        End If
        '    End If
        'End If
        Return TaxGroup

    End Function
#End Region

#End Region

#Region "Log Transaction"
    Public Sub LogTransaction(ByVal DocNum As Integer, ByVal ItemCode As String, _
                                    ByVal FromWhs As String, ByVal TransferedQty As Double, ByVal ProcessDate As Date)
        Dim sCode As String
        Dim sColumns As String
        Dim sValues As String
        Dim rsInsert As SAPbobsCOM.Recordset

        sCode = Me.getMaxCode("@REN_PORDR", "Code")

        sColumns = "Code, Name, U_DocNum, U_WhsCode, U_ItemCode, U_Quantity, U_RetQty, U_Date"
        sValues = "'" & sCode & "','" & sCode & "'," & DocNum & ",'" & FromWhs & "','" & ItemCode & "'," & TransferedQty & ", 0, Convert(DateTime,'" & ProcessDate.ToString("yyyyMMdd") & "')"

        strSQL = "Insert into [@REN_PORDR] (" & sColumns & ") Values (" & sValues & ")"
        oApplication.Utilities.ExecuteSQL(rsInsert, strSQL)

    End Sub

    Public Sub LogCreatedDocument(ByVal DocNum As Integer, ByVal CreatedDocType As SAPbouiCOM.BoLinkedObject, ByVal CreatedDocNum As String, ByVal sCreatedDate As String)
        Dim oUserTable As SAPbobsCOM.UserTable
        Dim sCode As String
        Dim CreatedDate As DateTime
        Try
            oUserTable = oApplication.Company.UserTables.Item("REN_DORDR")

            sCode = Me.getMaxCode("@REN_DORDR", "Code")

            If Not oUserTable.GetByKey(sCode) Then
                oUserTable.Code = sCode
                oUserTable.Name = sCode

                With oUserTable.UserFields.Fields
                    .Item("U_DocNum").Value = DocNum
                    .Item("U_DocType").Value = CInt(CreatedDocType)
                    .Item("U_DocEntry").Value = CInt(CreatedDocNum)

                    If sCreatedDate <> "" Then
                        CreatedDate = CDate(sCreatedDate.Insert(4, "/").Insert(7, "/"))
                        .Item("U_Date").Value = CreatedDate
                    Else
                        .Item("U_Date").Value = CDate(Format(Now, "Long Date"))
                    End If

                End With

                If oUserTable.Add <> 0 Then
                    Throw New Exception(oApplication.Company.GetLastErrorDescription)
                End If
            End If

        Catch ex As Exception
            Throw ex
        Finally
            oUserTable = Nothing
        End Try
    End Sub
#End Region

    Public Function getLocalCurrency(ByVal strCurrency As String) As Double
        Dim oTemp As SAPbobsCOM.Recordset
        oTemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oTemp.DoQuery("Select Maincurrncy from OADM")
        Return oTemp.Fields.Item(0).Value
    End Function

#Region "Get ExchangeRate"
    Public Function getExchangeRate(ByVal strCurrency As String) As Double
        Dim oTemp As SAPbobsCOM.Recordset
        Dim dblExchange As Double
        If GetCurrency("Local") = strCurrency Then
            dblExchange = 1
        Else
            oTemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oTemp.DoQuery("Select isNull(Rate,0) from ORTT where convert(nvarchar(10),RateDate,101)=Convert(nvarchar(10),getdate(),101) and currency='" & strCurrency & "'")
            dblExchange = oTemp.Fields.Item(0).Value
        End If
        Return dblExchange
    End Function

    Public Function getExchangeRate(ByVal strCurrency As String, ByVal dtdate As Date) As Double
        Dim oTemp As SAPbobsCOM.Recordset
        Dim strSql As String
        Dim dblExchange As Double
        If GetCurrency("Local") = strCurrency Then
            dblExchange = 1
        Else
            oTemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            strSql = "Select isNull(Rate,0) from ORTT where ratedate='" & dtdate.ToString("yyyy-MM-dd") & "' and currency='" & strCurrency & "'"
            oTemp.DoQuery(strSql)
            dblExchange = oTemp.Fields.Item(0).Value
        End If
        Return dblExchange
    End Function
#End Region

    Public Function GetDateTimeValue(ByVal DateString As String) As DateTime
        Dim objBridge As SAPbobsCOM.SBObob
        objBridge = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoBridge)
        Return objBridge.Format_StringToDate(DateString).Fields.Item(0).Value
    End Function

#Region "Get DocCurrency"
    Public Function GetDocCurrency(ByVal aDocEntry As Integer) As String
        Dim oTemp As SAPbobsCOM.Recordset
        oTemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oTemp.DoQuery("Select DocCur from OINV where docentry=" & aDocEntry)
        Return oTemp.Fields.Item(0).Value
    End Function
#End Region

#Region "GetEditTextValues"
    Public Function getEditTextvalue(ByVal aForm As SAPbouiCOM.Form, ByVal strUID As String) As String
        Dim oEditText As SAPbouiCOM.EditText
        oEditText = aForm.Items.Item(strUID).Specific
        Return oEditText.Value
    End Function
#End Region

#Region "Get Currency"
    Public Function GetCurrency(ByVal strChoice As String, Optional ByVal aCardCode As String = "") As String
        Dim strCurrQuery, Currency As String
        Dim oTempCurrency As SAPbobsCOM.Recordset
        If strChoice = "Local" Then
            strCurrQuery = "Select MainCurncy from OADM"
        Else
            strCurrQuery = "Select Currency from OCRD where CardCode='" & aCardCode & "'"
        End If
        oTempCurrency = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oTempCurrency.DoQuery(strCurrQuery)
        Currency = oTempCurrency.Fields.Item(0).Value
        Return Currency
    End Function

#End Region

    Public Function FormatDataSourceValue(ByVal Value As String) As Double
        Dim NewValue As Double

        If Value <> "" Then
            If Value.IndexOf(".") > -1 Then
                Value = Value.Replace(".", CompanyDecimalSeprator)
            End If

            If Value.IndexOf(CompanyThousandSeprator) > -1 Then
                Value = Value.Replace(CompanyThousandSeprator, "")
            End If
        Else
            Value = "0"

        End If

        ' NewValue = CDbl(Value)
        NewValue = Val(Value)

        Return NewValue


        'Dim dblValue As Double
        'Value = Value.Replace(CompanyThousandSeprator, "")
        'Value = Value.Replace(CompanyDecimalSeprator, System.Globalization.NumberFormatInfo.CurrentInfo.NumberDecimalSeparator)
        'dblValue = Val(Value)
        'Return dblValue
    End Function

    Public Function FormatScreenValues(ByVal Value As String) As Double
        Dim NewValue As Double

        If Value <> "" Then
            If Value.IndexOf(".") > -1 Then
                Value = Value.Replace(".", CompanyDecimalSeprator)
            End If
        Else
            Value = "0"
        End If

        'NewValue = CDbl(Value)
        NewValue = Val(Value)

        Return NewValue

        'Dim dblValue As Double
        'Value = Value.Replace(CompanyThousandSeprator, "")
        'Value = Value.Replace(CompanyDecimalSeprator, System.Globalization.NumberFormatInfo.CurrentInfo.NumberDecimalSeparator)
        'dblValue = Val(Value)
        'Return dblValue

    End Function

    Public Function SetScreenValues(ByVal Value As String) As String

        If Value.IndexOf(CompanyDecimalSeprator) > -1 Then
            Value = Value.Replace(CompanyDecimalSeprator, ".")
        End If

        Return Value

    End Function

    Public Function SetDBValues(ByVal Value As String) As String

        If Value.IndexOf(CompanyDecimalSeprator) > -1 Then
            Value = Value.Replace(CompanyDecimalSeprator, ".")
        End If

        Return Value

    End Function

    Public Sub RemoveComboValues(ByVal oForm As SAPbouiCOM.Form, ByVal strID As String)
        Try
            Dim oComboBox As SAPbouiCOM.ComboBox
            oComboBox = oForm.Items.Item(strID).Specific
            Dim intCount As Integer = oComboBox.ValidValues.Count - 1
            While intCount >= 0
                oComboBox.ValidValues.Remove(intCount, SAPbouiCOM.BoSearchKey.psk_Index)
                intCount -= 1
            End While
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

#Region "AddControls"
    Public Sub AddControls(ByVal objForm As SAPbouiCOM.Form, ByVal ItemUID As String, ByVal SourceUID As String, ByVal ItemType As SAPbouiCOM.BoFormItemTypes, ByVal position As String, Optional ByVal fromPane As Integer = 1, Optional ByVal toPane As Integer = 1, Optional ByVal linkedUID As String = "", Optional ByVal strCaption As String = "", Optional ByVal dblWidth As Double = 0, Optional ByVal dblTop As Double = 0, Optional ByVal Hight As Double = 0, Optional ByVal Enable As Boolean = True)
        Dim objNewItem, objOldItem As SAPbouiCOM.Item
        Dim ostatic As SAPbouiCOM.StaticText
        Dim oButton As SAPbouiCOM.Button
        Dim oCheckbox As SAPbouiCOM.CheckBox
        Dim oEditText As SAPbouiCOM.EditText
        Dim ofolder As SAPbouiCOM.Folder
        objOldItem = objForm.Items.Item(SourceUID)
        objNewItem = objForm.Items.Add(ItemUID, ItemType)
        With objNewItem
            If ItemType = SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON Then
                .Left = objOldItem.Left - 15
                .Top = objOldItem.Top + 1
                .LinkTo = linkedUID
            Else
                If position.ToUpper = "RIGHT" Then
                    .Left = objOldItem.Left + objOldItem.Width + 5
                    .Top = objOldItem.Top
                ElseIf position.ToUpper = "DOWN" Then
                    If ItemUID = "edWork" Then
                        .Left = objOldItem.Left + 40
                    Else
                        .Left = objOldItem.Left
                    End If
                    .Top = objOldItem.Top + objOldItem.Height + 3

                    .Width = objOldItem.Width
                    .Height = objOldItem.Height
                ElseIf position.ToUpper = "COPY" Then
                    .Top = objOldItem.Top
                    .Left = objOldItem.Left
                    .Height = objOldItem.Height
                    .Width = objOldItem.Width
                End If
            End If
            .FromPane = fromPane
            .ToPane = toPane
            If ItemType = SAPbouiCOM.BoFormItemTypes.it_STATIC Then
                .LinkTo = linkedUID
            End If
            .LinkTo = linkedUID
        End With
        If (ItemType = SAPbouiCOM.BoFormItemTypes.it_EDIT Or ItemType = SAPbouiCOM.BoFormItemTypes.it_STATIC) Then
            objNewItem.Width = objOldItem.Width
        End If
        If ItemType = SAPbouiCOM.BoFormItemTypes.it_BUTTON Then
            objNewItem.Width = objOldItem.Width '+ 50
            oButton = objNewItem.Specific
            oButton.Caption = strCaption
        ElseIf ItemType = SAPbouiCOM.BoFormItemTypes.it_FOLDER Then
            ofolder = objNewItem.Specific
            ofolder.Caption = strCaption
            ofolder.GroupWith(linkedUID)
        ElseIf ItemType = SAPbouiCOM.BoFormItemTypes.it_STATIC Then
            ostatic = objNewItem.Specific
            ostatic.Caption = strCaption
        ElseIf ItemType = SAPbouiCOM.BoFormItemTypes.it_CHECK_BOX Then
            oCheckbox = objNewItem.Specific
            oCheckbox.Caption = strCaption

        End If
        If dblWidth <> 0 Then
            objNewItem.Width = dblWidth
        End If

        If dblTop <> 0 Then
            objNewItem.Top = objNewItem.Top + dblTop
        End If
        If Hight <> 0 Then
            objNewItem.Height = objNewItem.Height + Hight
        End If
    End Sub
#End Region

#Region "Set / Get Values from Matrix"
    Public Function getMatrixValues(ByVal aMatrix As SAPbouiCOM.Matrix, ByVal coluid As String, ByVal intRow As Integer) As String
        Return aMatrix.Columns.Item(coluid).Cells.Item(intRow).Specific.value
    End Function
    Public Sub SetMatrixValues(ByVal aMatrix As SAPbouiCOM.Matrix, ByVal coluid As String, ByVal intRow As Integer, ByVal strvalue As String)
        aMatrix.Columns.Item(coluid).Cells.Item(intRow).Specific.value = strvalue
    End Sub
#End Region

#Region "Add Condition CFL"
    Public Sub AddConditionCFL(ByVal FormUID As String, ByVal strQuery As String, ByVal strQueryField As String, ByVal sCFL As String)
        Dim oRecordSet As SAPbobsCOM.Recordset
        Dim oCFLs As SAPbouiCOM.ChooseFromListCollection
        Dim Conditions As SAPbouiCOM.Conditions
        Dim oCond As SAPbouiCOM.Condition
        Dim oCFL As SAPbouiCOM.ChooseFromList
        Dim oCFLCreationParams As SAPbouiCOM.ChooseFromListCreationParams
        Dim sDocEntry As New ArrayList()
        Dim sDocNum As ArrayList
        Dim MatrixItem As ArrayList
        sDocEntry = New ArrayList()
        sDocNum = New ArrayList()
        MatrixItem = New ArrayList()

        Try
            oCFLs = oApplication.SBO_Application.Forms.Item(FormUID).ChooseFromLists
            oCFLCreationParams = oApplication.SBO_Application.CreateObject( _
                                    SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)

            oCFL = oCFLs.Item(sCFL)

            Dim oRec As SAPbobsCOM.Recordset
            oRec = DirectCast(oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
            oRec.DoQuery(strQuery)
            oRec.MoveFirst()

            Try
                If oRec.EoF Then
                    sDocEntry.Add("")
                Else
                    While Not oRec.EoF
                        Dim DocNum As String = oRec.Fields.Item(strQueryField).Value.ToString()
                        If DocNum <> "" Then
                            sDocEntry.Add(DocNum)
                        End If
                        oRec.MoveNext()
                    End While
                End If
            Catch generatedExceptionName As Exception
                Throw
            End Try

            'If IsMatrixCondition = True Then
            '    Dim oMatrix As SAPbouiCOM.Matrix
            '    oMatrix = DirectCast(oForm.Items.Item(Matrixname).Specific, SAPbouiCOM.Matrix)

            '    For a As Integer = 1 To oMatrix.RowCount
            '        If a <> pVal.Row Then
            '            MatrixItem.Add(DirectCast(oMatrix.Columns.Item(columnname).Cells.Item(a).Specific, SAPbouiCOM.EditText).Value)
            '        End If
            '    Next
            '    If removelist = True Then
            '        For xx As Integer = 0 To MatrixItem.Count - 1
            '            Dim zz As String = MatrixItem(xx).ToString()
            '            If sDocEntry.Contains(zz) Then
            '                sDocEntry.Remove(zz)
            '            End If
            '        Next
            '    End If
            'End If

            'oCFLs = oForm.ChooseFromLists
            'oCFLCreationParams = DirectCast(SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams), SAPbouiCOM.ChooseFromListCreationParams)
            'If systemMatrix = True Then
            '    Dim oCFLEvento As SAPbouiCOM.IChooseFromListEvent = Nothing
            '    oCFLEvento = DirectCast(pVal, SAPbouiCOM.IChooseFromListEvent)
            '    Dim sCFL_ID As String = Nothing
            '    sCFL_ID = oCFLEvento.ChooseFromListUID
            '    oCFL = oForm.ChooseFromLists.Item(sCFL_ID)
            'Else
            '    oCFL = oForm.ChooseFromLists.Item(sCHUD)
            'End If

            Conditions = New SAPbouiCOM.Conditions()
            oCFL.SetConditions(Conditions)
            Conditions = oCFL.GetConditions()
            oCond = Conditions.Add()
            oCond.BracketOpenNum = 2
            For i As Integer = 0 To sDocEntry.Count - 1
                If i > 0 Then
                    oCond.Relationship = SAPbouiCOM.BoConditionRelationship.cr_OR
                    oCond = Conditions.Add()
                    oCond.BracketOpenNum = 1
                End If

                oCond.[Alias] = strQueryField
                oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                oCond.CondVal = sDocEntry(i).ToString()
                If i + 1 = sDocEntry.Count Then
                    oCond.BracketCloseNum = 2
                Else
                    oCond.BracketCloseNum = 1
                End If
            Next

            oCFL.SetConditions(Conditions)


        Catch ex As Exception
            Throw ex
        End Try
    End Sub
#End Region

    Public Function getFreightName(ByVal strExpCode As String) As String
        Dim oTemp As SAPbobsCOM.Recordset
        Try
            oTemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oTemp.DoQuery("Select ExpnsName From OEXD Where ExpnsCode = '" + strExpCode + "'")
            Return oTemp.Fields.Item(0).Value
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Public Function getDocumentQuantity(ByVal strQuantity As String) As Double
        Dim dblQuant As Double
        Dim strTemp, strTemp1 As String
        Dim oRec As SAPbobsCOM.Recordset
        oRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRec.DoQuery("Select CurrCode  from OCRN")
        For intRow As Integer = 0 To oRec.RecordCount - 1
            strQuantity = strQuantity.Replace(oRec.Fields.Item(0).Value, "")
            oRec.MoveNext()
        Next
        strTemp1 = strQuantity
        strTemp = CompanyDecimalSeprator
        If CompanyDecimalSeprator <> "." Then
            If CompanyThousandSeprator <> strTemp Then
            End If
            strQuantity = strQuantity.Replace(".", ",")
        End If
        If strQuantity = "" Then
            Return 0
        End If
        Try
            dblQuant = Convert.ToDouble(strQuantity)
        Catch ex As Exception
            dblQuant = Convert.ToDouble(strTemp1)
        End Try

        Return dblQuant
    End Function

    Public Sub OpenFileDialogBox(ByVal oForm As SAPbouiCOM.Form, ByVal strID As String)
        Dim _retVal As String = String.Empty
        Try
            FileOpen()
            Dim oEdit As SAPbouiCOM.EditText
            ' oEdit = oForm.Items.Item("27").Specific
            ' oEdit.string = strFilepath
            CType(oForm.Items.Item(strID).Specific, SAPbouiCOM.StaticText).Caption = (strFilepath)
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

#Region "FileOpen"
    Private Sub FileOpen()
        Try
            Dim mythr As New System.Threading.Thread(AddressOf ShowFileDialog)
            mythr.SetApartmentState(Threading.ApartmentState.STA)
            mythr.Start()
            mythr.Join()
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub ShowFileDialog()
        Try
            Dim oDialogBox As New OpenFileDialog
            Dim strMdbFilePath As String
            Dim oProcesses() As Process
            Try
                oProcesses = Process.GetProcessesByName("SAP Business One")
                If oProcesses.Length <> 0 Then
                    For i As Integer = 0 To oProcesses.Length - 1
                        Dim MyWindow As New clsListener.WindowWrapper(oProcesses(i).MainWindowHandle)
                        If oDialogBox.ShowDialog(MyWindow) = DialogResult.OK Then
                            strMdbFilePath = oDialogBox.FileName
                            strFilepath = oDialogBox.FileName
                        Else
                        End If
                    Next
                End If
            Catch ex As Exception
                oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Finally
            End Try
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
#End Region

    Public Function ValidateFile(ByVal oForm As SAPbouiCOM.Form, ByVal strID As String) As Boolean
        Dim _retVal As Boolean = True
        Try
            Dim strPath As String = CType(oForm.Items.Item(strID).Specific, SAPbouiCOM.StaticText).Caption
            ' If Path.GetExtension(strPath) <> ".txt" Then
            If Path.GetExtension(strPath) <> ".xls" And Path.GetExtension(strPath) <> ".xlsx" Then
                _retVal = False
                oApplication.Utilities.Message("In Valid File Format...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Else
                _retVal = True
            End If
            Return _retVal
        Catch ex As Exception
            Throw ex
        End Try
    End Function


    Public Sub createDatatable()
        Try
            _oDt.TableName = "Budget"
            _oDt.Columns.Add("RowID", GetType(String))
            _oDt.Columns.Add("Year", GetType(String))
            _oDt.Columns.Add("OcrCode", GetType(String))
            _oDt.Columns.Add("OcrName", GetType(String))
            _oDt.Columns.Add("AcctCode", GetType(String))
            _oDt.Columns.Add("AcctName", GetType(String))
            _oDt.Columns.Add("Jan", GetType(String))
            _oDt.Columns.Add("Feb", GetType(String))
            _oDt.Columns.Add("Mar", GetType(String))
            _oDt.Columns.Add("Apr", GetType(String))
            _oDt.Columns.Add("May", GetType(String))
            _oDt.Columns.Add("June", GetType(String))
            _oDt.Columns.Add("July", GetType(String))
            _oDt.Columns.Add("Aug", GetType(String))
            _oDt.Columns.Add("Sep", GetType(String))
            _oDt.Columns.Add("Oct", GetType(String))
            _oDt.Columns.Add("Nov", GetType(String))
            _oDt.Columns.Add("Dec", GetType(String))
            _oDt.Columns.Add("SUCCESS", GetType(String))
            _oDt.Columns.Add("ERROR", GetType(String))
            _oDt.Columns.Add("RefCode", GetType(String))

        Catch ex As Exception

        End Try
    End Sub
    Public Function GetExcelData(ByVal oForm As SAPbouiCOM.Form, ByVal strID As String) As Boolean
        Dim _retVal As Boolean = False
        '   Dim _oDt As New DataTable
        oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Try
            _oDt.TableName = "SIRAW"
            _oDt.Columns.Add("RowID", GetType(String))
            _oDt.Columns.Add("CODE", GetType(String))
            _oDt.Columns.Add("DESCRIPTION", GetType(String))
            _oDt.Columns.Add("QTY", GetType(String))
            _oDt.Columns.Add("Price", GetType(String))
            _oDt.Columns.Add("OnHand", GetType(String))
            _oDt.Columns.Add("AQTY", GetType(String))
            _oDt.Columns.Add("SUCCESS", GetType(String))
            _oDt.Columns.Add("ERROR", GetType(String))
            Dim oTemp As SAPbobsCOM.Recordset
            oTemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Dim strWhs As String = CType(oForm.Items.Item("30").Specific, SAPbouiCOM.EditText).Value
            Dim strPath As String = CType(oForm.Items.Item(strID).Specific, SAPbouiCOM.StaticText).Caption
            Dim strItemCode As String
            Dim strCardCode As String = CType(oForm.Items.Item("3").Specific, SAPbouiCOM.EditText).Value
            Dim dblAmt As Single
            If strPath.Length > 0 Then
                Dim intCol As Integer = 0
                Dim txtRows() As String
                Dim fields() As String
                Dim oDr As DataRow
                txtRows = System.IO.File.ReadAllLines(strPath)
                Dim intRow As Integer = 0
                'For Each txtrow As String In txtRows
                '    If intRow = 0 Then
                '        fields = txtrow.Split(vbTab)
                '        For index As Integer = 0 To fields.Length - 1
                '            _oDt.Columns.Add(fields(intCol).ToUpper(), GetType(String)).Caption = fields(intCol).ToUpper()
                '            intCol += 1
                '        Next
                '        Exit For
                '    End If
                'Next

                intRow = 0
                For Each txtrow As String In txtRows
                    Dim strError As String = String.Empty
                    Dim blnSuccess As Boolean = True
                    If intRow = 0 Then
                        'fields = txtrow.Split(vbTab)
                        'If fields.Length >= 5 Then
                        '    For index As Integer = 0 To fields.Length
                        '        If fields(0).ToUpper() <> "CODE" Or fields(1).ToUpper() <> "QTY" Then 'Or fields(fields.Length - 3).ToUpper() <> "PCS" Or fields(fields.Length - 2).ToUpper() <> "SP" Or fields(fields.Length - 1).ToUpper() <> "VALUE" Then
                        '            Throw New Exception("In Valid Columns Found...1-CODE,2-DESCRIPTION ")
                        '            Exit Function
                        '        End If
                        '    Next
                        'Else
                        '    Throw New Exception("In Valid File Please Check File Format.")
                        '    Exit Function
                        'End If
                        'For index As Integer = 0 To fields.Length
                        '    If fields(0).ToUpper() <> "CODE" Or fields(1).ToUpper() <> "QTY" Then ' Or fields(fields.Length - 3).ToUpper() <> "PCS" Or fields(fields.Length - 2).ToUpper() <> "SP" Or fields(fields.Length - 1).ToUpper() <> "VALUE" Then
                        '        Throw New Exception("In Valid Columns Found...1-CODE,2-DESCRIPTION and Final Columns Should be PCS,SP,VALUE")
                        '        Exit Function
                        '    End If
                        'Next
                    ElseIf intRow > 0 Then
                        fields = txtrow.Split(vbTab)
                        oDr = _oDt.NewRow()
                        oDr.Item("RowID") = intRow.ToString
                        oDr.Item("CODE") = fields(0)
                        oDr.Item("QTY") = fields(1)
                        oTemp.DoQuery("Select ItemName from OITM where ItemCode='" & fields(0) & "'")
                        If oTemp.RecordCount > 0 Then
                            oDr.Item("DESCRIPTION") = oTemp.Fields.Item(0).Value
                        Else
                            oDr.Item("DESCRIPTION") = fields(0) ' "-"
                            strError = " Item Code Not Found "
                            blnSuccess = False
                        End If
                        Dim strPrice As String
                        Dim oCombo As SAPbouiCOM.ComboBox
                        oCombo = oForm.Items.Item("14").Specific
                        strPrice = oCombo.Selected.Value
                        If strPrice = "" Then
                            oDr.Item("Price") = "0"
                        Else
                            oDr.Item("Price") = GetItemPrice_Pricelist(strCardCode, fields(0), 0, strPrice)
                            ' row("SP") = GetItemPrice(strCardCode, strItemCode, dblAmt, dtDate).ToString()
                        End If

                        oTemp.DoQuery("Select  OnHand,(OnHand-IsCommited+OnOrder) 'Available'  from OITW where whscode='" & strWhs & "' and  ItemCode='" & fields(0) & "'")
                        If oTemp.RecordCount > 0 Then
                            oDr.Item("OnHand") = oTemp.Fields.Item("OnHand").Value
                            oDr.Item("AQTY") = oTemp.Fields.Item("Available").Value
                        Else
                            oDr.Item("OnHand") = "0"
                            oDr.Item("AQTY") = "0"
                        End If

                        '   oDr.ItemArray = fields
                        If Not blnSuccess Then
                            oDr.Item("SUCCESS") = "0"
                            oDr.Item("ERROR") = strError
                        Else
                            oDr.Item("SUCCESS") = "1"
                        End If
                        _oDt.Rows.Add(oDr)
                    End If
                    intRow = intRow + 1
                Next
            End If

            If Not IsNothing(_oDt) Then
                If _oDt.Rows.Count > 0 Then
                    ' Dim _unPivot As DataTable = unPivotTable(oForm, _oDt)
                    Dim strDtXML As String = getXMLstring(_oDt) ' _unPivot)
                    oRecordSet.DoQuery("Exec [Insert_SQImport] '" + strDtXML + "'")
                    _retVal = True
                End If
            End If

            Return _retVal
        Catch ex As Exception
            Throw ex
        End Try
    End Function
    Public Function ReadExcelDat_GRPO(ByVal aORderNo As String, ByVal aItemCode As String, ByVal aItemName As String, ByVal aQty As String, ByVal aPrice As String, ByVal aValue As String, ByVal aForm As SAPbouiCOM.Form, ByVal aRowID As Integer, aRefCode As String) As Boolean
        Dim _retVal As Boolean = False
        '   Dim _oDt As New DataTable
        oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Try
            ' _oDt.Columns.Clear()
            If _oDtGRPO.Columns.Count <= 0 Then
                _oDtGRPO.TableName = "GRPO"
                _oDtGRPO.Columns.Add("RowID", GetType(String))
                _oDtGRPO.Columns.Add("ORDNO", GetType(String))
                _oDtGRPO.Columns.Add("CODE", GetType(String))
                _oDtGRPO.Columns.Add("DESCRIPTION", GetType(String))
                _oDtGRPO.Columns.Add("QTY", GetType(String))
                _oDtGRPO.Columns.Add("Price", GetType(String))
                _oDtGRPO.Columns.Add("Total", GetType(String))
                _oDtGRPO.Columns.Add("BaseRef", GetType(String))
                _oDtGRPO.Columns.Add("BaseLine", GetType(String))
                _oDtGRPO.Columns.Add("SUCCESS", GetType(String))
                _oDtGRPO.Columns.Add("ERROR", GetType(String))
                _oDtGRPO.Columns.Add("RefCode", GetType(String))
            End If
            Dim oTemp As SAPbobsCOM.Recordset
            oTemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Dim strWhs As String = CType(aForm.Items.Item("30").Specific, SAPbouiCOM.EditText).Value
            Dim strPath As String = "" 'CType(aForm.Items.Item(strID).Specific, SAPbouiCOM.StaticText).Caption
            Dim strItemCode As String
            Dim strCardCode As String = CType(aForm.Items.Item("3").Specific, SAPbouiCOM.EditText).Value
            Dim dblAmt As Single


            Dim oDr As DataRow

            Dim strError As String = String.Empty
            Dim blnSuccess As Boolean = True
            oDr = _oDtGRPO.NewRow()
            oDr.Item("RowID") = aRowID
            oDr.Item("ORDNO") = aORderNo
            oDr.Item("CODE") = aItemCode.Replace("-", "")
            oDr.Item("QTY") = aQty
            oDr.Item("Price") = aPrice
            oDr.Item("Total") = aValue
            oDr.Item("RefCode") = aRefCode
            strError = ""
            oTemp.DoQuery("Select ItemName from OITM where ItemCode='" & aItemCode.Replace("-", "") & "'")
            If oTemp.RecordCount > 0 Then
                oDr.Item("DESCRIPTION") = oTemp.Fields.Item(0).Value
            Else
                oDr.Item("DESCRIPTION") = aItemCode ' "-"
                strError = " Item Code Not Found "
                blnSuccess = False
            End If
            '  Dim s As String = "Select T0.DocEntry,T0.LineNum,* from POR1 T0 inner Join OPOR T1 On T1.DocEntry=T0.DocEntry where T1.CardCode='" & strCardCode & "' and T0.ItemCode='" & aItemCode.Replace("-", "") & "' and T0.LineStatus='O' and convert(Varchar,T1.DocNum)='" & aORderNo & "'"
            Dim s As String = "Select T0.DocEntry,T0.LineNum,* from POR1 T0 inner Join OPOR T1 On T1.DocEntry=T0.DocEntry where T1.CardCode='" & strCardCode & "' and T0.ItemCode='" & aItemCode.Replace("-", "") & "' and T0.LineStatus='O' and convert(Varchar,T1.NumAtCard)='" & aORderNo & "'"
            oTemp.DoQuery(s)
            If oTemp.RecordCount > 0 Then
                oDr.Item("BaseRef") = oTemp.Fields.Item(0).Value
                oDr.Item("BaseLine") = oTemp.Fields.Item(1).Value
            Else
                strError = "Purchase Order does not exists: DocNum : " & aORderNo & ": Customer Code: " & strCardCode
                blnSuccess = False
                oDr.Item("BaseRef") = 0
                oDr.Item("BaseLine") = 0
            End If
            If Not blnSuccess Then
                oDr.Item("SUCCESS") = "0"
                oDr.Item("ERROR") = strError
            Else
                oDr.Item("SUCCESS") = "1"
                oDr.Item("ERROR") = ""
            End If
            _oDtGRPO.Rows.Add(oDr)
            Return _retVal
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Public Function ReadExcelDat_PO(ByVal aItemCode As String, ByVal aItemName As String, ByVal aQty As String, ByVal aPrice As String, ByVal aForm As SAPbouiCOM.Form, ByVal aRowID As Integer, aRefCode As String) As Boolean
        Dim _retVal As Boolean = False
        '   Dim _oDt As New DataTable
        oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Try
            ' _oDt.Columns.Clear()
            If _oDtPO.Columns.Count <= 0 Then
                _oDtPO.TableName = "PO"
                _oDtPO.Columns.Add("RowID", GetType(String))
                _oDtPO.Columns.Add("CODE", GetType(String))
                _oDtPO.Columns.Add("DESCRIPTION", GetType(String))
                _oDtPO.Columns.Add("QTY", GetType(String))
                _oDtPO.Columns.Add("Price", GetType(String))
                _oDtPO.Columns.Add("SUCCESS", GetType(String))
                _oDtPO.Columns.Add("ERROR", GetType(String))
                _oDtPO.Columns.Add("RefCode", GetType(String))
            End If
            Dim oTemp As SAPbobsCOM.Recordset
            oTemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Dim strWhs As String = CType(aForm.Items.Item("30").Specific, SAPbouiCOM.EditText).Value
            Dim strPath As String = "" 'CType(aForm.Items.Item(strID).Specific, SAPbouiCOM.StaticText).Caption
            Dim strItemCode As String
            Dim strCardCode As String = CType(aForm.Items.Item("3").Specific, SAPbouiCOM.EditText).Value
            Dim dblAmt As Single


            Dim oDr As DataRow

            Dim strError As String = String.Empty
            Dim blnSuccess As Boolean = True
            oDr = _oDtPO.NewRow()
            oDr.Item("RowID") = aRowID
            oDr.Item("CODE") = aItemCode.Replace("-", "")
            oDr.Item("QTY") = aQty
            oDr.Item("Price") = aPrice
            'oDr.Item("Total") = aValue
            strError = ""
            oTemp.DoQuery("Select ItemName from OITM where ItemCode='" & aItemCode.Replace("-", "") & "'")
            If oTemp.RecordCount > 0 Then
                oDr.Item("DESCRIPTION") = oTemp.Fields.Item(0).Value
                strError = " "
            Else
                oDr.Item("DESCRIPTION") = aItemName ' "-"
                strError = "Item Code Not Found"
                blnSuccess = False
            End If
            'Dim s As String = "Select T0.DocEntry,T0.LineNum,* from POR1 T0 inner Join OPOR T1 On T1.DocEntry=T0.DocEntry where T1.CardCode='" & strCardCode & "' and T0.ItemCode='" & aItemCode & "' and T0.LineStatus='O' and convert(Varchar,T1.DocNum)='" & aORderNo & "'"
            'oTemp.DoQuery(s)
            'If oTemp.RecordCount > 0 Then
            '    oDr.Item("BaseRef") = oTemp.Fields.Item(0).Value
            '    oDr.Item("BaseLine") = oTemp.Fields.Item(1).Value
            'Else
            '    oDr.Item("BaseRef") = 0
            '    oDr.Item("BaseLine") = 0
            'End If
            If Not blnSuccess Then
                oDr.Item("SUCCESS") = "0"
                oDr.Item("ERROR") = strError
            Else
                oDr.Item("SUCCESS") = "1"
                oDr.Item("ERROR") = ""
            End If
            oDr.Item("RefCode") = aRefCode
            _oDtPO.Rows.Add(oDr)
            Return _retVal
        Catch ex As Exception
            Throw ex
        End Try
    End Function
    'Public Function ReadExcelDat(ByVal aItemCode As String, ByVal aQty As String, ByVal aForm As SAPbouiCOM.Form, ByVal aRowID As Integer, aRefCode As String, aWhsCode As String) As Boolean
    '    Dim _retVal As Boolean = False
    '    '   Dim _oDt As New DataTable
    '    oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
    '    Try
    '        If _oDt.Columns.Count <= 0 Then
    '            _oDt.TableName = "Budget"
    '            _oDt.Columns.Add("RowID", GetType(String))
    '            _oDt.Columns.Add("Year", GetType(String))
    '            _oDt.Columns.Add("AcctCode", GetType(String))
    '            _oDt.Columns.Add("AcctName", GetType(String))
    '            _oDt.Columns.Add("Jan", GetType(String))
    '            _oDt.Columns.Add("Feb", GetType(String))
    '            _oDt.Columns.Add("Mar", GetType(String))
    '            _oDt.Columns.Add("Apr", GetType(String))
    '            _oDt.Columns.Add("May", GetType(String))
    '            _oDt.Columns.Add("June", GetType(String))
    '            _oDt.Columns.Add("July", GetType(String))
    '            _oDt.Columns.Add("Aug", GetType(String))
    '            _oDt.Columns.Add("Sep", GetType(String))
    '            _oDt.Columns.Add("Oct", GetType(String))
    '            _oDt.Columns.Add("Nov", GetType(String))
    '            _oDt.Columns.Add("Dec", GetType(String))
    '            _oDt.Columns.Add("SUCCESS", GetType(String))
    '            _oDt.Columns.Add("ERROR", GetType(String))
    '            _oDt.Columns.Add("RefCode", GetType(String))
    '        End If
    '        Dim oTemp As SAPbobsCOM.Recordset
    '        oTemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
    '        Dim oDr As DataRow
    '        Dim strError As String = String.Empty
    '        Dim blnSuccess As Boolean = True
    '        oDr = _oDt.NewRow()
    '        oDr.Item("RowID") = aRowID
    '        oDr.Item("Year") = Year() '
    '        oDr.Item("QTY") = aQty
    '        oTemp.DoQuery("Select ItemName from OITM where ItemCode='" & aItemCode.Replace("-", "") & "'")
    '        If oTemp.RecordCount > 0 Then
    '            oDr.Item("DESCRIPTION") = oTemp.Fields.Item(0).Value
    '        Else
    '            oDr.Item("DESCRIPTION") = aItemCode ' "-"
    '            strError = " Item Code Not Found "
    '            blnSuccess = False
    '        End If
    '        Dim strPrice As String
    '        Dim oCombo As SAPbouiCOM.ComboBox
    '        oCombo = aForm.Items.Item("14").Specific
    '        strPrice = oCombo.Selected.Value
    '        If strPrice = "" Then
    '            oDr.Item("Price") = "0"
    '        Else
    '            oDr.Item("Price") = GetItemPrice_Pricelist(strCardCode, aItemCode.Replace("-", ""), 0, strPrice)
    '            ' row("SP") = GetItemPrice(strCardCode, strItemCode, dblAmt, dtDate).ToString()
    '        End If
    '        oTemp.DoQuery("Select  isnull(Sum(OnHand),0) from OITW where   ItemCode='" & aItemCode.Replace("-", "") & "'")
    '        dblTotalOnHand = oTemp.Fields.Item(0).Value

    '        oTemp.DoQuery("Select  OnHand,(OnHand-IsCommited+OnOrder) 'Available'  from OITW where whscode='" & strWhs & "' and  ItemCode='" & aItemCode.Replace("-", "") & "'")
    '        If oTemp.RecordCount > 0 Then
    '            oDr.Item("OnHand") = oTemp.Fields.Item("OnHand").Value
    '            oDr.Item("AQTY") = dblTotalOnHand.ToString  ' oTemp.Fields.Item("Available").Value
    '        Else
    '            oDr.Item("OnHand") = "0"
    '            oDr.Item("AQTY") = dblTotalOnHand.ToString ' "0"
    '        End If

    '        '   oDr.ItemArray = fields
    '        If Not blnSuccess Then
    '            oDr.Item("SUCCESS") = "0"
    '            oDr.Item("ERROR") = strError
    '        Else
    '            oDr.Item("SUCCESS") = "1"
    '        End If
    '        oDr.Item("RefCode") = aRefCode.ToString
    '        oDr.Item("WhsCode") = aWhsCode.ToString
    '        _oDt.Rows.Add(oDr)

    '        Return _retVal
    '    Catch ex As Exception
    '        Throw ex
    '    End Try
    'End Function


    Public Function ReadExcelDat_Budget(aForm As SAPbouiCOM.Form, ByVal aRowID As Integer, aYear As String, ocrCode As String, ocrName As String, AcctCode As String, AcctName As String, Jan As String, Feb As String, Mar As String, Apr As String, May As String, June As String, July As String, Aug As String, Sep As String, Oct As String, Nov As String, Dec As String, aRefCode As String) As Boolean
        Dim _retVal As Boolean = False
        '   Dim _oDt As New DataTable
        'Dim ocrName As String = ocrCode
        oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Try
            If _oDt.Columns.Count <= 0 Then
                _oDt.TableName = "Budget"
                _oDt.Columns.Add("RowID", GetType(String))
                _oDt.Columns.Add("Year", GetType(String))
                _oDt.Columns.Add("OcrCode", GetType(String))
                _oDt.Columns.Add("OcrName", GetType(String))
                _oDt.Columns.Add("AcctCode", GetType(String))
                _oDt.Columns.Add("AcctName", GetType(String))
                _oDt.Columns.Add("Jan", GetType(String))
                _oDt.Columns.Add("Feb", GetType(String))
                _oDt.Columns.Add("Mar", GetType(String))
                _oDt.Columns.Add("Apr", GetType(String))
                _oDt.Columns.Add("May", GetType(String))
                _oDt.Columns.Add("June", GetType(String))
                _oDt.Columns.Add("July", GetType(String))
                _oDt.Columns.Add("Aug", GetType(String))
                _oDt.Columns.Add("Sep", GetType(String))
                _oDt.Columns.Add("Oct", GetType(String))
                _oDt.Columns.Add("Nov", GetType(String))
                _oDt.Columns.Add("Dec", GetType(String))
                _oDt.Columns.Add("SUCCESS", GetType(String))
                _oDt.Columns.Add("ERROR", GetType(String))
                _oDt.Columns.Add("RefCode", GetType(String))
            End If
            Dim oTemp As SAPbobsCOM.Recordset
            oTemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Dim strWhs As String = "" 'CType(aForm.Items.Item("30").Specific, SAPbouiCOM.EditText).Value
            Dim strPath As String = "" 'CType(aForm.Items.Item(strID).Specific, SAPbouiCOM.StaticText).Caption
            Dim strItemCode As String
            Dim strCardCode As String = "" ' CType(aForm.Items.Item("3").Specific, SAPbouiCOM.EditText).Value
            Dim dblAmt, dblTotalOnHand As Single
            Dim oDr As DataRow
            Dim strError As String = String.Empty
            Dim blnSuccess As Boolean = True
            oDr = _oDt.NewRow()
            oDr.Item("RowID") = aRowID
            oDr.Item("Year") = aYear
            oDr.Item("AcctCode") = AcctCode
            oDr.Item("AcctName") = AcctName
            oDr.Item("OcrCode") = ocrCode
            oDr.Item("OcrName") = ocrName
            If Jan = "" Then
                Jan = "0"
            End If
            If Feb = "" Then
                Feb = "0"
            End If
            If Mar = "" Then
                Mar = "0"
            End If
            If Apr = "" Then
                Apr = "0"
            End If
            If May = "" Then
                May = "0"
            End If
            If June = "" Then
                June = "0"
            End If
            If July = "0" Then
                July = ""
            End If
            If Aug = "" Then
                Aug = "0"
            End If
            If Sep = "" Then
                Sep = "0"
            End If
            If Oct = "" Then
                Oct = "0"
            End If
            If Nov = "" Then
                Nov = "0"
            End If
            If Dec = "" Then
                Dec = "0"
            End If

            oDr.Item("Jan") = Jan
            oDr.Item("Feb") = Feb
            oDr.Item("Mar") = Mar
            oDr.Item("Apr") = Apr
            oDr.Item("May") = May
            oDr.Item("June") = June
            oDr.Item("July") = July
            oDr.Item("Aug") = Aug
            oDr.Item("Sep") = Sep
            oDr.Item("Oct") = Oct
            oDr.Item("Nov") = Nov
            oDr.Item("Dec") = Dec
            oDr.Item("RefCode") = aRefCode

            oTemp.DoQuery("Select * from OACT where FormatCode='" & AcctCode & "'")
            If oTemp.RecordCount > 0 Then
                oDr.Item("AcctName") = oTemp.Fields.Item("AcctName").Value
            Else
                oDr.Item("AcctName") = AcctCode 'AcctName  ' "-"
                strError = " Account Code Not Found "
                blnSuccess = False
            End If

            oTemp.DoQuery("Select * from OPRC where PrcCode='" & ocrCode & "'")
            If oTemp.RecordCount > 0 Then
                oDr.Item("OcrName") = oTemp.Fields.Item("prcName").Value
            Else
                oDr.Item("OcrCode") = ocrCode  'AcctName  ' "-"
                strError = " Country dimension Not Found "
                blnSuccess = False
            End If

            '   oDr.ItemArray = fields
            If Not blnSuccess Then
                oDr.Item("SUCCESS") = "0"
                oDr.Item("ERROR") = strError
            Else
                oDr.Item("SUCCESS") = "1"
            End If

            _oDt.Rows.Add(oDr)

            Return _retVal
        Catch ex As Exception
            Throw ex
        End Try
    End Function
    Private Function unPivotTable(ByVal oForm As SAPbouiCOM.Form, ByVal dt As DataTable) As DataTable
        Dim dtNew As New DataTable("SIIMPORT")
        Dim sQuery As String
        Dim oRecordSet As SAPbobsCOM.Recordset
        oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Dim strWhs As String = CType(oForm.Items.Item("30").Specific, SAPbouiCOM.EditText).Value
        Try
            If Not (dt.Columns.Count > 0 AndAlso dt.Rows.Count > 0) Then
                Return dtNew
            End If

            dtNew.Columns.Add("CODE", GetType(String))
            dtNew.Columns.Add("DESCRIPTION", GetType(String))
            dtNew.Columns.Add("BIN", GetType(String))
            dtNew.Columns.Add("QTY", GetType(String))
            dtNew.Columns.Add("SP", GetType(String))
            dtNew.Columns.Add("PCS", GetType(String))
            dtNew.Columns.Add("TOTAL", GetType(String))
            dtNew.Columns.Add("AQTY", GetType(String))
            dtNew.Columns.Add("SUCCESS", GetType(String))
            dtNew.Columns.Add("ERROR", GetType(String))

            Dim row As DataRow = Nothing
            For k As Integer = 0 To dt.Rows.Count - 1
                For j As Integer = 2 To dt.Columns.Count - 4
                    If dt.Rows(k)("CODE").ToString() <> "TOTAL" And CDbl(IIf(dt.Rows(k)(j).ToString() = "", "0", dt.Rows(k)(j).ToString())) > 0 Then
                        CType(oForm.Items.Item("73").Specific, SAPbouiCOM.StaticText).Caption = "Importing Item : " + dt.Rows(k)("CODE").ToString() + "..."
                        Dim strError As String = String.Empty
                        Dim blnSuccess As Boolean = True

                        If dt.Rows(k)("CODE").ToString() <> "" Then

                            row = dtNew.NewRow()
                            row("CODE") = dt.Rows(k)("CODE").ToString()
                            row("DESCRIPTION") = dt.Rows(k)("DESCRIPTION").ToString().Replace("'", "")
                            row("BIN") = dt.Columns(j).ColumnName.ToString()
                            row("QTY") = IIf(dt.Rows(k)(j).ToString() = "", "0", dt.Rows(k)(j).ToString())
                            row("PCS") = dt.Rows(k)("PCS").ToString()
                            'row("SP") = IIf(dt.Rows(k)("SP").ToString().Replace("""", "").Replace(",", "") = "", "0", dt.Rows(k)("SP").ToString().Replace("""", "").Replace(",", ""))
                            Dim strItemCode As String = dt.Rows(k)("CODE").ToString()
                            Dim strCardCode As String = CType(oForm.Items.Item("3").Specific, SAPbouiCOM.EditText).Value
                            Dim dblAmt As Single
                            Dim dtDate As DateTime = System.DateTime.Now
                            ' row("SP") = GetItemPrice(strCardCode, strItemCode, dblAmt, dtDate).ToString()
                            row("TOTAL") = dt.Rows(k)("VALUE").ToString().Replace("""", "").Replace(",", "")

                            'Item Not Availble
                            Dim strManSerial As String = String.Empty
                            Dim strManBatch As String = String.Empty
                            sQuery = " Select ItemCode,ManSerNum,ManBtchNum From OITM T0 "
                            sQuery += " Where ItemCode = '" + dt.Rows(k)("CODE").ToString() + "'"
                            oRecordSet.DoQuery(sQuery)
                            If oRecordSet.EoF Then
                                strError = " Item Code Not Found "
                                blnSuccess = False
                                row("SP") = 0 ' GetItemPrice(strCardCode, strItemCode, dblAmt, dtDate).ToString()
                            Else
                                strManSerial = oRecordSet.Fields.Item("ManSerNum").Value
                                strManBatch = oRecordSet.Fields.Item("ManBtchNum").Value
                                Dim strPrice As String
                                Dim oCombo As SAPbouiCOM.ComboBox
                                oCombo = oForm.Items.Item("14").Specific
                                strPrice = oCombo.Selected.Value
                                If strPrice = "" Then
                                    row("SP") = dt.Rows(k)("SP").ToString().Replace("""", "").Replace(",", "")
                                Else
                                    row("SP") = GetItemPrice_Pricelist(strCardCode, strItemCode, dblAmt, strPrice)
                                    ' row("SP") = GetItemPrice(strCardCode, strItemCode, dblAmt, dtDate).ToString()
                                End If

                            End If

                            'Bin Not Available
                            sQuery = " Select SL1Code From OBIN T0 "
                            sQuery += " Where SL1Code = '" + dt.Columns(j).ColumnName.ToString() + "'"
                            oRecordSet.DoQuery(sQuery)
                            If oRecordSet.EoF Then
                                blnSuccess = False
                                strError += " Bin Not Found "
                            End If

                            'Stock Availablity foR Batch Managed Item.
                            If strManBatch = "Y" Then
                                'Stock Not Available
                                If CInt(row("QTY")) > 0 Then
                                    sQuery = " Select SUM(T0.OnHandQty) As OnHandQty From OBBQ T0 "
                                    sQuery += " JOIN OBIN T1 ON T0.BinAbs = T1.AbsEntry "
                                    sQuery += " Where ItemCode = '" + dt.Rows(k)("CODE").ToString() + "'"
                                    sQuery += " And T0.WhsCode = '" + strWhs + "'"
                                    sQuery += " And T1.SL1Code = '" + dt.Columns(j).ColumnName.ToString() + "'"
                                    oRecordSet.DoQuery(sQuery)
                                    If Not oRecordSet.EoF Then
                                        If CInt(oRecordSet.Fields.Item(0).Value) < CInt(row("QTY")) Then
                                            row("AQTY") = oRecordSet.Fields.Item(0).Value
                                            blnSuccess = False
                                            strError += " Stock Not Available In Bin"
                                        Else
                                            row("AQTY") = oRecordSet.Fields.Item(0).Value
                                        End If
                                    Else
                                        row("AQTY") = "0"
                                    End If
                                Else
                                    row("AQTY") = "0"
                                End If
                            ElseIf (strManBatch = "N" And strManSerial = "N") Then   'Stock Availablity foR Non Batch & Serail Managd.
                                If CInt(row("QTY")) > 0 Then

                                    'sQuery = " Select "
                                    'sQuery += " ( "
                                    'sQuery += " ISNULL((Select SUM(T0.Quantity) As 'In' "
                                    'sQuery += " From OILM T0 JOIN OBTL T1 On T0.MessageID = T1.MessageID  "
                                    'sQuery += " JOIN OBIN T2 ON T1.BinAbs = T2.AbsEntry "
                                    'sQuery += " Where ActionType = 1 "
                                    'sQuery += " And T2.SL1Code = '" + dt.Columns(j).ColumnName.ToString() + "' "
                                    'sQuery += " And T0.ItemCode = '" + dt.Rows(k)("CODE").ToString() + "'"
                                    'sQuery += " And T0.LocCode = '" + strWhs + "'"
                                    'sQuery += "),0) "
                                    'sQuery += " - "
                                    'sQuery += " ISNULL((Select SUM(T0.Quantity) As 'OUT' "
                                    'sQuery += " From OILM T0 JOIN OBTL T1 On T0.MessageID = T1.MessageID  "
                                    'sQuery += " JOIN OBIN T2 ON T1.BinAbs = T2.AbsEntry "
                                    'sQuery += " Where ActionType = 2  "
                                    'sQuery += " And T2.SL1Code = '" + dt.Columns(j).ColumnName.ToString() + "' "
                                    'sQuery += " And T0.ItemCode = '" + dt.Rows(k)("CODE").ToString() + "'"
                                    'sQuery += " And T0.LocCode = '" + strWhs + "'"
                                    'sQuery += " ),0) "
                                    'sQuery += " ) As 'OnHand' "

                                    sQuery = " Select SUM(T0.OnHandQty) As OnHandQty From OIBQ T0 "
                                    sQuery += " JOIN OBIN T1 ON T0.BinAbs = T1.AbsEntry "
                                    sQuery += " Where ItemCode = '" + dt.Rows(k)("CODE").ToString() + "'"
                                    sQuery += " And T0.WhsCode = '" + strWhs + "'"
                                    sQuery += " And T1.SL1Code = '" + dt.Columns(j).ColumnName.ToString() + "'"
                                    oRecordSet.DoQuery(sQuery)

                                    If Not oRecordSet.EoF Then
                                        If CInt(oRecordSet.Fields.Item(0).Value) < CInt(row("QTY")) Then
                                            row("AQTY") = oRecordSet.Fields.Item(0).Value
                                            blnSuccess = False
                                            strError += " Stock Not Available In Bin"
                                        Else
                                            row("AQTY") = oRecordSet.Fields.Item(0).Value
                                        End If
                                    Else
                                        row("AQTY") = "0"
                                    End If
                                Else
                                    row("AQTY") = "0"
                                End If
                            End If

                            If Not blnSuccess Then
                                row("SUCCESS") = "0"
                                row("ERROR") = strError
                            Else
                                row("SUCCESS") = "1"
                            End If

                            dtNew.Rows.Add(row)

                        End If
                    End If
                Next
            Next
        Catch ex As Exception
            Throw ex
        End Try
        Return dtNew
    End Function

    Public Function getXMLstring(ByVal oDt As System.Data.DataTable) As String
        Dim _retVal As String = String.Empty
        Try
            Dim sr As New System.IO.StringWriter()
            oDt.WriteXml(sr, False)
            _retVal = sr.ToString()
            Return _retVal
        Catch ex As Exception
            Throw ex
        End Try
    End Function


    Public Function saveAsDraft_GRPO(ByVal oForm As SAPbouiCOM.Form, strBatchNum As String) As Boolean
        Try
            Dim _retVal As String = False
            Dim sQuery As String
            Dim oIRecordSet, oRecordSet, oISBatchSerial, oBatch, oSerial, oBin As SAPbobsCOM.Recordset
            Dim oSIDraft As SAPbobsCOM.Documents
            Dim intCurrentLine As Integer = 0
            Dim intBatchNo As Integer = 0
            Dim intSerialBatchNo As Integer = 0
            Dim oHashTable As Hashtable


            oSIDraft = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseDeliveryNotes)
            oIRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oISBatchSerial = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oBatch = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oSerial = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oBin = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

            'oSIDraft.DocObjectCode = SAPbobsCOM.BoObjectTypes.oQuotations

            'oSIDraft.Series = CType(oForm.Items.Item("20").Specific, SAPbouiCOM.ComboBox).Selected.Value
            ' oSIDraft.UserFields.Fields.Item("U_IsImport").Value = "Y"
            oSIDraft.CardCode = CType(oForm.Items.Item("3").Specific, SAPbouiCOM.EditText).Value
            oSIDraft.CardName = CType(oForm.Items.Item("5").Specific, SAPbouiCOM.EditText).Value

            'If CType(oForm.Items.Item("7").Specific, SAPbouiCOM.ComboBox).Selected.Value <> "" Then
            '    oSIDraft.ContactPersonCode = CType(oForm.Items.Item("7").Specific, SAPbouiCOM.ComboBox).Selected.Value
            'End If

            oSIDraft.NumAtCard = CType(oForm.Items.Item("9").Specific, SAPbouiCOM.EditText).Value
            oSIDraft.DocDate = GetDateTimeValue(CType(oForm.Items.Item("24").Specific, SAPbouiCOM.EditText).Value)
            oSIDraft.DocDueDate = GetDateTimeValue(CType(oForm.Items.Item("26").Specific, SAPbouiCOM.EditText).Value)
            oSIDraft.TaxDate = GetDateTimeValue(CType(oForm.Items.Item("28").Specific, SAPbouiCOM.EditText).Value)

            '    oSIDraft.ShipToCode = CType(oForm.Items.Item("33").Specific, SAPbouiCOM.ComboBox).Selected.Description
            'oSIDraft.Address = CType(oForm.Items.Item("34").Specific, SAPbouiCOM.EditText).Value
            'oSIDraft.Address2 = CType(oForm.Items.Item("39").Specific, SAPbouiCOM.EditText).Value

            'oSIDraft.Comments = CType(oForm.Items.Item("18").Specific, SAPbouiCOM.EditText).Value

            'oSIDraft.ControlAccount = CType(oForm.Items.Item("49").Specific, SAPbouiCOM.EditText).Value
            '   oSIDraft.JournalMemo = CType(oForm.Items.Item("42").Specific, SAPbouiCOM.EditText).Value
            '  oSIDraft.PayToCode = CType(oForm.Items.Item("44").Specific, SAPbouiCOM.ComboBox).Selected.Value
            '  oSIDraft.PaymentMethod = CType(oForm.Items.Item("51").Specific, SAPbouiCOM.ComboBox).Selected.Value
            Dim strCode As String
            Dim oEdit As SAPbouiCOM.EditText
            oEdit = oForm.Items.Item("81").Specific
            strCode = oEdit.String
            sQuery = "Select Code,Sum(Convert(Decimal(18,2),Qty)) As Quantity,Price From Z_SIIM1 Group By Code,Success "
            sQuery += " Having Success = '1' And Sum(Convert(Decimal(18,2),Qty)) > 0 and Flag='Y'  "

            sQuery = "Select Code,Sum(Convert(Decimal(18,2),Qty)) As Quantity,Avg(Convert(Decimal(18,2),Price)) 'Price',BaseRef,BaseLine From Z_ODLN where "
            sQuery += "  Success = '1' And (Convert(Decimal(18,2),Qty)) > 0 and Flag='Y' and RefCode='" & strCode & "'  group by Code,BaseRef,BaseLine  "
            oIRecordSet.DoQuery(sQuery)
            If Not oIRecordSet.EoF Then
                Dim oItem As SAPbobsCOM.Items
                While Not oIRecordSet.EoF
                    oHashTable = New Hashtable()
                    intBatchNo = 0
                    intSerialBatchNo = 0
                    If oIRecordSet.Fields.Item("Quantity").Value > 0 Then
                        If intCurrentLine > 0 Then
                            oSIDraft.Lines.Add()
                        End If
                        oSIDraft.Lines.SetCurrentLine(intCurrentLine)
                        If oIRecordSet.Fields.Item("BaseRef").Value > 0 Then

                            oItem = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oItems)
                            If oItem.GetByKey(oIRecordSet.Fields.Item("Code").Value) Then

                                oSIDraft.Lines.ItemCode = oIRecordSet.Fields.Item("Code").Value
                                '  oSIDraft.Lines.WarehouseCode = CType(oForm.Items.Item("30").Specific, SAPbouiCOM.EditText).Value
                                oSIDraft.Lines.BaseType = SAPbobsCOM.BoObjectTypes.oPurchaseOrders
                                oSIDraft.Lines.BaseEntry = oIRecordSet.Fields.Item("BaseRef").Value
                                oSIDraft.Lines.BaseLine = oIRecordSet.Fields.Item("BaseLine").Value
                                oSIDraft.Lines.Quantity = oIRecordSet.Fields.Item("Quantity").Value
                                ' oSIDraft.Lines.UnitPrice = oIRecordSet.Fields.Item("Price").Value
                                If CType(oForm.Items.Item("30").Specific, SAPbouiCOM.EditText).Value <> "" Then
                                    oSIDraft.Lines.WarehouseCode = CType(oForm.Items.Item("30").Specific, SAPbouiCOM.EditText).Value
                                End If
                                'oTest.DoQuery("Select isnull(DfltWH,'') from OITM where ItemCode=" & oIRecordSet.Fields.Item("Code").Value & "'")
                                'If oTest.Fields.Item(0).Value = "" Then
                                '    oSIDraft.Lines.WarehouseCode = CType(oForm.Items.Item("30").Specific, SAPbouiCOM.EditText).Value
                                'Else
                                '    oSIDraft.Lines.WarehouseCode = oTest.Fields.Item(0).Value
                                'End If
                                If oItem.ManageBatchNumbers = SAPbobsCOM.BoYesNoEnum.tYES Then
                                    oSIDraft.Lines.BatchNumbers.BatchNumber = strBatchNum
                                    oSIDraft.Lines.BatchNumbers.Quantity = oIRecordSet.Fields.Item("Quantity").Value
                                    '  oSIDraft.Lines.BatchNumbers.Add()
                                End If

                            End If

                        Else
                            oItem = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oItems)
                            If oItem.GetByKey(oIRecordSet.Fields.Item("Code").Value) Then


                                oSIDraft.Lines.ItemCode = oIRecordSet.Fields.Item("Code").Value
                                'oSIDraft.Lines.WarehouseCode = CType(oForm.Items.Item("30").Specific, SAPbouiCOM.EditText).Value
                                Dim oTest As SAPbobsCOM.Recordset
                                oTest = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                If CType(oForm.Items.Item("30").Specific, SAPbouiCOM.EditText).Value <> "" Then
                                    oSIDraft.Lines.WarehouseCode = CType(oForm.Items.Item("30").Specific, SAPbouiCOM.EditText).Value
                                Else
                                    oTest.DoQuery("Select isnull(DfltWH,'') from OITM where ItemCode=" & oIRecordSet.Fields.Item("Code").Value & "'")
                                    If oTest.Fields.Item(0).Value = "" Then
                                        oSIDraft.Lines.WarehouseCode = CType(oForm.Items.Item("30").Specific, SAPbouiCOM.EditText).Value
                                    Else
                                        oSIDraft.Lines.WarehouseCode = oTest.Fields.Item(0).Value
                                    End If
                                End If
                                If oItem.ManageBatchNumbers = SAPbobsCOM.BoYesNoEnum.tYES Then
                                    oSIDraft.Lines.BatchNumbers.BatchNumber = strBatchNum
                                    oSIDraft.Lines.BatchNumbers.Quantity = oIRecordSet.Fields.Item("Quantity").Value
                                    '  oSIDraft.Lines.BatchNumbers.Add()
                                End If



                                oSIDraft.Lines.Quantity = oIRecordSet.Fields.Item("Quantity").Value
                                oSIDraft.Lines.UnitPrice = oIRecordSet.Fields.Item("Price").Value
                            End If
                        End If

                        ' oSIDraft.Lines.DiscountPercent = CType(oForm.Items.Item("75").Specific, SAPbouiCOM.EditText).Value
                        intCurrentLine += 1
                    End If
                    oIRecordSet.MoveNext()

                End While
            End If
            Dim strData As String = String.Empty

            'oSIDraft.GetAsXML()
            'MessageBox.Show(oSIDraft.GetAsXML())

            Dim iRet As Integer = oSIDraft.Add()
            If iRet = 0 Then
                _retVal = True
            Else
                _retVal = False
            End If
            Return _retVal

        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Public Function CreateItemCode(ByVal oform As SAPbouiCOM.Form) As Boolean
        Dim oRec, oRec1 As SAPbobsCOM.Recordset
        Dim oItem As SAPbobsCOM.Items
        Dim strItemCode, strDescription, strQuery As String

        If oApplication.Company.InTransaction() Then
            oApplication.Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
        End If
        oApplication.Company.StartTransaction()
        oItem = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oItems)
        oRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRec1 = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRec.DoQuery("Select * from Z_OPOR where ERROR='Item Code Not Found'")
        For intRow As Integer = 0 To oRec.RecordCount - 1

            strItemCode = oRec.Fields.Item("CODE").Value
            oRec1.DoQuery("Select * from OITM where ItemCode='" & strItemCode & "'")
            If oRec1.RecordCount <= 0 Then


                strDescription = oRec.Fields.Item("DESCRIPTION").Value
                oItem = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oItems)
                oItem.ItemCode = strItemCode
                oItem.ItemName = strDescription
                Dim strGroup As String
                If strItemCode.Length > 5 Then
                    strGroup = strItemCode.Substring(0, 5)
                    oRec1.DoQuery("Select * from OITB where ItmsGrpNam='" & strGroup & "'")
                    If oRec1.RecordCount > 0 Then
                        oItem.ItemsGroupCode = oRec1.Fields.Item("ItmsGrpCod").Value

                    End If
                Else
                    strGroup = strItemCode
                    oRec1.DoQuery("Select * from OITB where ItmsGrpNam='" & strGroup & "'")
                    If oRec1.RecordCount > 0 Then
                        oItem.ItemsGroupCode = oRec1.Fields.Item("ItmsGrpCod").Value
                    End If
                End If
                oItem.ManageBatchNumbers = SAPbobsCOM.BoYesNoEnum.tYES
                If oItem.Add <> 0 Then
                    oApplication.Utilities.Message(oApplication.Company.GetLastErrorDescription, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    If oApplication.Company.InTransaction() Then
                        oApplication.Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                    End If
                    Return False
                Else
                    strQuery = "Update Z_OPOR set Flag='Y', SUCCESS = '1',ERROR='' where RowID='" & oRec.Fields.Item("RowID").Value & "'"
                    oRec1.DoQuery(strQuery)
                End If
            Else
                strQuery = "Update Z_OPOR set Flag='Y', SUCCESS = '1',ERROR='' where RowID='" & oRec.Fields.Item("RowID").Value & "'"
                oRec1.DoQuery(strQuery)
            End If
            oRec.MoveNext()
        Next
        If oApplication.Company.InTransaction() Then
            oApplication.Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
        End If
        Return True
    End Function
    Public Function saveAsDraft_PO(ByVal oForm As SAPbouiCOM.Form) As Boolean
        Try
            Dim _retVal As String = False
            Dim sQuery As String
            Dim oIRecordSet, oRecordSet, oISBatchSerial, oBatch, oSerial, oBin As SAPbobsCOM.Recordset
            Dim oSIDraft As SAPbobsCOM.Documents
            Dim intCurrentLine As Integer = 0
            Dim intBatchNo As Integer = 0
            Dim intSerialBatchNo As Integer = 0
            Dim oHashTable As Hashtable


            oSIDraft = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseOrders)
            oIRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oISBatchSerial = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oBatch = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oSerial = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oBin = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

            'oSIDraft.DocObjectCode = SAPbobsCOM.BoObjectTypes.oQuotations

            'oSIDraft.Series = CType(oForm.Items.Item("20").Specific, SAPbouiCOM.ComboBox).Selected.Value
            ' oSIDraft.UserFields.Fields.Item("U_IsImport").Value = "Y"
            oSIDraft.CardCode = CType(oForm.Items.Item("3").Specific, SAPbouiCOM.EditText).Value
            oSIDraft.CardName = CType(oForm.Items.Item("5").Specific, SAPbouiCOM.EditText).Value

            'If CType(oForm.Items.Item("7").Specific, SAPbouiCOM.ComboBox).Selected.Value <> "" Then
            '    oSIDraft.ContactPersonCode = CType(oForm.Items.Item("7").Specific, SAPbouiCOM.ComboBox).Selected.Value
            'End If

            oSIDraft.NumAtCard = CType(oForm.Items.Item("9").Specific, SAPbouiCOM.EditText).Value
            oSIDraft.DocDate = GetDateTimeValue(CType(oForm.Items.Item("24").Specific, SAPbouiCOM.EditText).Value)
            oSIDraft.DocDueDate = GetDateTimeValue(CType(oForm.Items.Item("26").Specific, SAPbouiCOM.EditText).Value)
            oSIDraft.TaxDate = GetDateTimeValue(CType(oForm.Items.Item("28").Specific, SAPbouiCOM.EditText).Value)

            '    oSIDraft.ShipToCode = CType(oForm.Items.Item("33").Specific, SAPbouiCOM.ComboBox).Selected.Description
            'oSIDraft.Address = CType(oForm.Items.Item("34").Specific, SAPbouiCOM.EditText).Value
            'oSIDraft.Address2 = CType(oForm.Items.Item("39").Specific, SAPbouiCOM.EditText).Value

            'oSIDraft.Comments = CType(oForm.Items.Item("18").Specific, SAPbouiCOM.EditText).Value

            'oSIDraft.ControlAccount = CType(oForm.Items.Item("49").Specific, SAPbouiCOM.EditText).Value
            '   oSIDraft.JournalMemo = CType(oForm.Items.Item("42").Specific, SAPbouiCOM.EditText).Value
            '  oSIDraft.PayToCode = CType(oForm.Items.Item("44").Specific, SAPbouiCOM.ComboBox).Selected.Value
            '  oSIDraft.PaymentMethod = CType(oForm.Items.Item("51").Specific, SAPbouiCOM.ComboBox).Selected.Value

            'sQuery = "Select Code,Sum(Convert(Decimal(18,2),Qty)) As Quantity,Price From Z_SIIM1 Group By Code,Success "
            'sQuery += " Having Success = '1' And Sum(Convert(Decimal(18,2),Qty)) > 0 and Flag='Y'  "

            Dim oEdit As SAPbouiCOM.EditText
            Dim strCode As String
            oEdit = oForm.Items.Item("81").Specific
            strCode = oEdit.String

            sQuery = "Select Code,sum((Convert(Decimal(18,2),Qty))) As Quantity,avg(convert(decimal(18,2),price)) 'Price' From Z_OPOR where "
            sQuery += "  Success = '1' And (Convert(Decimal(18,2),Qty)) > 0 and Flag='Y' and RefCode='" & strCode & "' group by Code  "
            oIRecordSet.DoQuery(sQuery)
            Dim oTest As SAPbobsCOM.Recordset
            oTest = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            If Not oIRecordSet.EoF Then
                While Not oIRecordSet.EoF
                    oHashTable = New Hashtable()
                    intBatchNo = 0
                    intSerialBatchNo = 0
                    If oIRecordSet.Fields.Item("Quantity").Value > 0 Then
                        If intCurrentLine > 0 Then
                            oSIDraft.Lines.Add()
                        End If
                        oSIDraft.Lines.SetCurrentLine(intCurrentLine)
                        If 1 = 2 Then 'oIRecordSet.Fields.Item("BaseRef").Value > 0 Then
                            oSIDraft.Lines.ItemCode = oIRecordSet.Fields.Item("Code").Value
                            '  oSIDraft.Lines.WarehouseCode = CType(oForm.Items.Item("30").Specific, SAPbouiCOM.EditText).Value
                            oSIDraft.Lines.BaseType = SAPbobsCOM.BoObjectTypes.oPurchaseOrders
                            oSIDraft.Lines.BaseEntry = oIRecordSet.Fields.Item("BaseRef").Value
                            oSIDraft.Lines.BaseLine = oIRecordSet.Fields.Item("BaseLine").Value
                            oSIDraft.Lines.Quantity = oIRecordSet.Fields.Item("Quantity").Value
                            ' oSIDraft.Lines.UnitPrice = oIRecordSet.Fields.Item("Price").Value
                        Else
                            oSIDraft.Lines.ItemCode = oIRecordSet.Fields.Item("Code").Value
                            oSIDraft.Lines.WarehouseCode = CType(oForm.Items.Item("30").Specific, SAPbouiCOM.EditText).Value
                            oSIDraft.Lines.Quantity = oIRecordSet.Fields.Item("Quantity").Value
                            oSIDraft.Lines.UnitPrice = oIRecordSet.Fields.Item("Price").Value
                        End If

                        ' oSIDraft.Lines.DiscountPercent = CType(oForm.Items.Item("75").Specific, SAPbouiCOM.EditText).Value

                        intCurrentLine += 1
                    End If
                    oIRecordSet.MoveNext()

                End While
            End If
            Dim strData As String = String.Empty

            'oSIDraft.GetAsXML()
            'MessageBox.Show(oSIDraft.GetAsXML())

            Dim iRet As Integer = oSIDraft.Add()
            If iRet = 0 Then
                _retVal = True
            Else
                _retVal = False
            End If
            Return _retVal

        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Private Function SaveUpdateClaim(ByVal aYear As String, aRefCode As String) As Boolean
        Try
            Dim oRecSet, oTemp As SAPbobsCOM.Recordset
            oRecSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oTemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Dim oGeneralService As SAPbobsCOM.GeneralService
            Dim oGeneralData1 As SAPbobsCOM.GeneralData
            Dim oGeneralParams As SAPbobsCOM.GeneralDataParams
            Dim oCompanyService As SAPbobsCOM.CompanyService
            Dim oChildren1 As SAPbobsCOM.GeneralDataCollection
            Dim sQuery As String
            oCompanyService = oApplication.Company.GetCompanyService()
            Dim oChild As SAPbobsCOM.GeneralData
            oGeneralService = oCompanyService.GetGeneralService("BUDGET_BY_COUNTRY")
            oGeneralData1 = oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)
            oGeneralParams = oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams)
            oTemp.DoQuery("select * from [@BUDGET_COUNTRY_DOC] where U_Year=" & aYear)

            If oTemp.RecordCount <= 0 Then
                '   oGeneralData1 = oGeneralService.GetByParams(oGeneralParams)
                oGeneralData1.SetProperty("U_Year", aYear)
                oChildren1 = oGeneralData1.Child("BUDGET_COUNTRY_ROW")
                sQuery = "Select * From Z_OBDG where "
                sQuery += "  Success = '1' And   year='" & aYear & "' and  RefCode='" & aRefCode & "'  Order by convert(numeric,RowID )"
                '  dbcon.strQuery = "Select * from ""U_OPRQ""   where ""SessionId""='" & objEN.SessionId & "' AND ""EmpId""='" & objEN.EmpId & "'"
                oRecSet.DoQuery(sQuery)
                For inlloop As Integer = 0 To oRecSet.RecordCount - 1
                    oChild = oChildren1.Add()
                    oChild.SetProperty("U_OcrCode1", oRecSet.Fields.Item("OcrCode").Value)
                    oChild.SetProperty("U_OcrName1", oRecSet.Fields.Item("OcrName").Value)
                    oChild.SetProperty("U_AcctCode", oRecSet.Fields.Item("AcctCode").Value)
                    oChild.SetProperty("U_AcctName", oRecSet.Fields.Item("AcctName").Value)
                    oChild.SetProperty("U_Budget_Jan", oRecSet.Fields.Item("Jan").Value)
                    oChild.SetProperty("U_Budget_Feb", oRecSet.Fields.Item("Feb").Value)
                    oChild.SetProperty("U_Budget_Mar", oRecSet.Fields.Item("Mar").Value)
                    oChild.SetProperty("U_Budget_Apr", oRecSet.Fields.Item("Apr").Value)
                    oChild.SetProperty("U_Budget_May", oRecSet.Fields.Item("May").Value)
                    oChild.SetProperty("U_Budget_Jun", oRecSet.Fields.Item("June").Value)
                    oChild.SetProperty("U_Budget_Jul", oRecSet.Fields.Item("July").Value)
                    oChild.SetProperty("U_Budget_Aug", oRecSet.Fields.Item("Aug").Value)
                    oChild.SetProperty("U_Budget_Sep", oRecSet.Fields.Item("Sep").Value)
                    oChild.SetProperty("U_Budget_Oct", oRecSet.Fields.Item("Oct").Value)
                    oChild.SetProperty("U_Budget_Nov", oRecSet.Fields.Item("Nov").Value)
                    oChild.SetProperty("U_Budget_Dec", oRecSet.Fields.Item("Dec").Value)


                    oRecSet.MoveNext()
                Next
                oGeneralParams = oGeneralService.Add(oGeneralData1)
            Else
                Dim intDocEntry As Integer = oTemp.Fields.Item("DocEntry").Value
                'oGeneralParams.SetProperty("DocEntry", oTemp.Fields.Item("DocEntry").Value)
                'oGeneralData1 = oGeneralService.GetByParams(oGeneralParams)
                'oGeneralData1.SetProperty("U_Year", aYear)
                ''oGeneralData1.SetProperty("U_Z_Destination", ddldestination.SelectedValue)
                ''oGeneralData1.SetProperty("U_Z_OrdPatient", ddlPatients.SelectedValue)

                'oChildren1 = oGeneralData1.Child("BUDGET_COUNTRY_ROW")

                sQuery = "Select * From Z_OBDG where "
                sQuery += "  Success = '1' And   year='" & aYear & "' and  RefCode='" & aRefCode & "' Order by convert(numeric,RowID ) "
                oRecSet.DoQuery(sQuery)
                Dim stQuery As String
                For intRow As Integer = 0 To oRecSet.RecordCount - 1
                    stQuery = "select * from [@BUDGET_COUNTRY_ROW] where   U_OcrCode1='" & oRecSet.Fields.Item("OcrCode").Value & "' and U_AcctCode='" & oRecSet.Fields.Item("AcctCode").Value & "'"
                    oTemp.DoQuery(stQuery)
                    If oTemp.RecordCount > 0 Then
                        Dim intLineID As Integer = oTemp.Fields.Item("LineID").Value
                        If oRecSet.Fields.Item("Jan").Value > 0 Then
                            oTemp.DoQuery("Update [@BUDGET_COUNTRY_ROW] set U_Budget_Jan='" & oRecSet.Fields.Item("Jan").Value & "' where LineID=" & intLineID & " and DocEntry=" & intDocEntry)
                        End If
                        If oRecSet.Fields.Item("Feb").Value > 0 Then
                            oTemp.DoQuery("Update [@BUDGET_COUNTRY_ROW] set U_Budget_Feb='" & oRecSet.Fields.Item("Feb").Value & "' where LineID=" & intLineID & " and DocEntry=" & intDocEntry)

                        End If
                        If oRecSet.Fields.Item("Mar").Value > 0 Then
                            oTemp.DoQuery("Update [@BUDGET_COUNTRY_ROW] set U_Budget_Mar='" & oRecSet.Fields.Item("Mar").Value & "' where LineID=" & intLineID & " and DocEntry=" & intDocEntry)

                        End If
                        If oRecSet.Fields.Item("Apr").Value > 0 Then
                            oTemp.DoQuery("Update [@BUDGET_COUNTRY_ROW] set U_Budget_Apr='" & oRecSet.Fields.Item("Apr").Value & "' where LineID=" & intLineID & " and DocEntry=" & intDocEntry)

                        End If
                        If oRecSet.Fields.Item("May").Value > 0 Then
                            oTemp.DoQuery("Update [@BUDGET_COUNTRY_ROW] set U_Budget_May='" & oRecSet.Fields.Item("May").Value & "' where LineID=" & intLineID & " and DocEntry=" & intDocEntry)

                        End If
                        If oRecSet.Fields.Item("June").Value > 0 Then
                            oTemp.DoQuery("Update [@BUDGET_COUNTRY_ROW] set U_Budget_Jun='" & oRecSet.Fields.Item("June").Value & "' where LineID=" & intLineID & " and DocEntry=" & intDocEntry)

                        End If
                        If oRecSet.Fields.Item("July").Value > 0 Then
                            oTemp.DoQuery("Update [@BUDGET_COUNTRY_ROW] set U_Budget_Jul='" & oRecSet.Fields.Item("July").Value & "' where LineID=" & intLineID & " and DocEntry=" & intDocEntry)

                        End If
                        If oRecSet.Fields.Item("Aug").Value > 0 Then
                            oTemp.DoQuery("Update [@BUDGET_COUNTRY_ROW] set U_Budget_Aug='" & oRecSet.Fields.Item("Aug").Value & "' where LineID=" & intLineID & " and DocEntry=" & intDocEntry)

                        End If
                        If oRecSet.Fields.Item("Sep").Value > 0 Then
                            oTemp.DoQuery("Update [@BUDGET_COUNTRY_ROW] set U_Budget_Sep='" & oRecSet.Fields.Item("Sep").Value & "' where LineID=" & intLineID & " and DocEntry=" & intDocEntry)

                        End If
                        If oRecSet.Fields.Item("Oct").Value > 0 Then
                            oTemp.DoQuery("Update [@BUDGET_COUNTRY_ROW] set U_Budget_Oct='" & oRecSet.Fields.Item("Oct").Value & "' where LineID=" & intLineID & " and DocEntry=" & intDocEntry)

                        End If
                        If oRecSet.Fields.Item("Nov").Value > 0 Then
                            oTemp.DoQuery("Update [@BUDGET_COUNTRY_ROW] set U_Budget_Nov='" & oRecSet.Fields.Item("Nov").Value & "' where LineID=" & intLineID & " and DocEntry=" & intDocEntry)

                        End If
                        If oRecSet.Fields.Item("Dec").Value > 0 Then
                            oTemp.DoQuery("Update [@BUDGET_COUNTRY_ROW] set U_Budget_Dec='" & oRecSet.Fields.Item("Dec").Value & "' where LineID=" & intLineID & " and DocEntry=" & intDocEntry)
                        End If
                        'oChild.SetProperty("U_Budget_Jan", oRecSet.Fields.Item("Jan").Value)
                        'oChild.SetProperty("U_Budget_Feb", oRecSet.Fields.Item("Feb").Value)
                        'oChild.SetProperty("U_Budget_Mar", oRecSet.Fields.Item("Mar").Value)
                        'oChild.SetProperty("U_Budget_Apr", oRecSet.Fields.Item("Apr").Value)
                        'oChild.SetProperty("U_Budget_May", oRecSet.Fields.Item("May").Value)
                        'oChild.SetProperty("U_Budget_Jun", oRecSet.Fields.Item("June").Value)
                        'oChild.SetProperty("U_Budget_Jul", oRecSet.Fields.Item("July").Value)
                        'oChild.SetProperty("U_Budget_Aug", oRecSet.Fields.Item("Aug").Value)
                        'oChild.SetProperty("U_Budget_Sep", oRecSet.Fields.Item("Sep").Value)
                        'oChild.SetProperty("U_Budget_Oct", oRecSet.Fields.Item("Oct").Value)
                        'oChild.SetProperty("U_Budget_Nov", oRecSet.Fields.Item("Nov").Value)
                        'oChild.SetProperty("U_Budget_Dec", oRecSet.Fields.Item("Dec").Value)


                    Else
                        oGeneralParams.SetProperty("DocEntry", intDocEntry)
                        oGeneralData1 = oGeneralService.GetByParams(oGeneralParams)
                        oGeneralData1.SetProperty("U_Year", aYear)
                        oChildren1 = oGeneralData1.Child("BUDGET_COUNTRY_ROW")
                        oChild = oChildren1.Add()
                        oChild.SetProperty("U_OcrCode1", oRecSet.Fields.Item("OcrCode").Value)
                        oChild.SetProperty("U_OcrName1", oRecSet.Fields.Item("OcrName").Value)
                        oChild.SetProperty("U_AcctCode", oRecSet.Fields.Item("AcctCode").Value)
                        oChild.SetProperty("U_AcctName", oRecSet.Fields.Item("AcctName").Value)
                        oChild.SetProperty("U_Budget_Jan", oRecSet.Fields.Item("Jan").Value)
                        oChild.SetProperty("U_Budget_Feb", oRecSet.Fields.Item("Feb").Value)
                        oChild.SetProperty("U_Budget_Mar", oRecSet.Fields.Item("Mar").Value)
                        oChild.SetProperty("U_Budget_Apr", oRecSet.Fields.Item("Apr").Value)
                        oChild.SetProperty("U_Budget_May", oRecSet.Fields.Item("May").Value)
                        oChild.SetProperty("U_Budget_Jun", oRecSet.Fields.Item("June").Value)
                        oChild.SetProperty("U_Budget_Jul", oRecSet.Fields.Item("July").Value)
                        oChild.SetProperty("U_Budget_Aug", oRecSet.Fields.Item("Aug").Value)
                        oChild.SetProperty("U_Budget_Sep", oRecSet.Fields.Item("Sep").Value)
                        oChild.SetProperty("U_Budget_Oct", oRecSet.Fields.Item("Oct").Value)
                        oChild.SetProperty("U_Budget_Nov", oRecSet.Fields.Item("Nov").Value)
                        oChild.SetProperty("U_Budget_Dec", oRecSet.Fields.Item("Dec").Value)
                        oGeneralService.Update(oGeneralData1)

                    End If
                    oRecSet.MoveNext()
                Next
                '  sQuery = "select * from [@BUDGET_COUNTRY_ROW] where   U_OcrCode1='"
                'oRecSet.DoQuery(sQuery)
                'For inlloop As Integer = 0 To oRecSet.RecordCount - 1
                '    oChild = oChildren1.Add()
                '    oChild.SetProperty("U_Z_ItemCode", oRecSet.Fields.Item("ItemCode").Value)
                '    oChild.SetProperty("U_Z_ItemName", oRecSet.Fields.Item("ItemName").Value)
                '    oChild.SetProperty("U_Z_OrdQty", oRecSet.Fields.Item("OrderQty").Value)
                '    oChild.SetProperty("U_Z_OrdUom", oRecSet.Fields.Item("OrderUom").Value)
                '    oChild.SetProperty("U_Z_OrdUomDesc", oRecSet.Fields.Item("OrderUomDesc").Value)
                '    oChild.SetProperty("U_Z_AltItemCode", oRecSet.Fields.Item("AltItemCode").Value)
                '    oChild.SetProperty("U_Z_AltItemName", oRecSet.Fields.Item("AltItemDesc").Value)
                '    oChild.SetProperty("U_Z_DeliQty", oRecSet.Fields.Item("DelQty").Value)
                '    oChild.SetProperty("U_Z_DelUom", oRecSet.Fields.Item("DelUom").Value)
                '    oChild.SetProperty("U_Z_DelUomDesc", oRecSet.Fields.Item("DelUomDesc").Value)
                '    oChild.SetProperty("U_Z_RecQty", oRecSet.Fields.Item("ReceivedQty").Value)
                '    oChild.SetProperty("U_Z_RecUom", oRecSet.Fields.Item("ReceivedUom").Value)
                '    oChild.SetProperty("U_Z_RecUomDesc", oRecSet.Fields.Item("ReceivedUomDesc").Value)
                '    oChild.SetProperty("U_Z_BarCode", oRecSet.Fields.Item("Barcode").Value)
                '    oChild.SetProperty("U_Z_LineStatus", oRecSet.Fields.Item("LineStatus").Value)
                '    oChild.SetProperty("U_Z_AppStatus", dbcon.DocApproval("PR", lbldept.Text.Trim()))
                '    oRecSet.MoveNext()
                'Next

            End If
            'Dim strDocEntry As String = oGeneralParams.GetProperty("DocEntry")
            'dbcon.strQuery = "Delete from  ""@Z_PRQ1""  where ""DocEntry""='" & strDocEntry & "' and ""U_Z_ItemCode"" Like '%D'"
            'oRecSet.DoQuery(dbcon.strQuery)
            'If ddlNewStatus.SelectedValue = "S" Then
            '    intTempID = dbcon.GetTemplateID("PRD", lbldept.Text.Trim())
            '    If intTempID <> "0" Then
            '        dbcon.UpdateApprovalRequired("@Z_PRQ1", "DocEntry", strDocEntry, "Y", intTempID)
            '    Else
            '        dbcon.UpdateApprovalRequired("@Z_PRQ1", "DocEntry", strDocEntry, "N", intTempID)
            '    End If
            '    If strDocEntry <> "" Then
            '        dbcon.InitialMessage("Purchase Requisition", strDocEntry, dbcon.DocApproval("PRD", lbldept.Text.Trim()), intTempID, lblempname.Text.Trim(), "PRD", dbcon.objMainCompany, strDocEntry)
            '    End If
            'End If
            'dbcon.strmsg = "Success"
            Return True
        Catch ex As Exception
            Return False
            'Page.ClientScript.RegisterStartupScript(Me.GetType(), "js", "<script>alert('" & ex.Message & "')</script>")
        End Try
    End Function


    Public Function saveAsDraft(ByVal oForm As SAPbouiCOM.Form) As Boolean
        Try
            Dim _retVal As String = False
            Dim sQuery As String
            Dim oIRecordSet, oRecordSet, oISBatchSerial, oBatch, oSerial, oBin As SAPbobsCOM.Recordset
            '  Dim oSIDraft As SAPbobsCOM.Documents
            Dim intCurrentLine As Integer = 0
            Dim intBatchNo As Integer = 0
            Dim intSerialBatchNo As Integer = 0
            Dim oHashTable As Hashtable


            'oSIDraft = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oQuotations)
            oIRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oISBatchSerial = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oBatch = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oSerial = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oBin = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

            'oSIDraft.DocObjectCode = SAPbobsCOM.BoObjectTypes.oQuotations

          
            Dim oEdit As SAPbouiCOM.EditText
            Dim strCode As String
            oEdit = oForm.Items.Item("81").Specific
            strCode = oEdit.String
            Dim aYear As String
            oBatch.DoQuery("Select Year,Count(*) from Z_OBDG where Success = '1' And     RefCode='" & strCode & "'  Group  by Year ")
            For intRow As Integer = 0 To oBatch.RecordCount - 1
                aYear = oBatch.Fields.Item(0).Value
                SaveUpdateClaim(aYear, strCode)
                'sQuery = "Select * From Z_OBDG where "
                'sQuery += "  Success = '1' And  and year='" & aYear & "' and  RefCode='" & strCode & "'  Order by Code "
                'oIRecordSet.DoQuery(sQuery)
                'If Not oIRecordSet.EoF Then

                '    While Not oIRecordSet.EoF
                '        oHashTable = New Hashtable()
                '        intBatchNo = 0
                '        intSerialBatchNo = 0

                '        If oIRecordSet.Fields.Item("Quantity").Value > 0 Then

                '            intCurrentLine += 1
                '        End If
                '        oIRecordSet.MoveNext()

                '    End While
                'End If
                'Dim strData As String = String.Empty

                ''oSIDraft.GetAsXML()
                ''MessageBox.Show(oSIDraft.GetAsXML())

                'Dim iRet As Integer = oSIDraft.Add()
                oBatch.MoveNext()
            Next
            'If iRet = 0 Then
            '    _retVal = True
            'Else
            '    _retVal = False
            'End If
            Return True

        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Public Function GetItemPrice(ByVal cardCode As String, ByVal itemCode As String, ByVal amount As Single, ByVal refDate As Date) As Double
        Try
            Dim vObj As SAPbobsCOM.SBObob
            Dim rs As SAPbobsCOM.Recordset
            vObj = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoBridge)
            rs = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            rs = vObj.GetItemPrice(cardCode, itemCode, amount, refDate)
            If Not rs.EoF Then
                Return CDbl(rs.Fields.Item(0).Value)
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Public Function GetItemPrice_Pricelist(ByVal cardCode As String, ByVal itemCode As String, ByVal amount As Single, ByVal PriceList As String) As Double
        Try
            Dim vObj As SAPbobsCOM.SBObob
            Dim rs As SAPbobsCOM.Recordset
            vObj = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoBridge)
            rs = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            rs.DoQuery("select Price   from ITM1 where PriceList=" & PriceList & " and ItemCode='" & itemCode & "'")
            If Not rs.EoF Then
                Return CDbl(rs.Fields.Item(0).Value)
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Function


End Class

