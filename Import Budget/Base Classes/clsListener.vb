Public Class clsListener
    Inherits Object

    Private ThreadClose As New Threading.Thread(AddressOf CloseApp)
    Private WithEvents _SBO_Application As SAPbouiCOM.Application
    Private _Company As SAPbobsCOM.Company
    Private _Utilities As clsUtilities
    Private _Collection As Hashtable
    Private _LookUpCollection As Hashtable
    Private _FormUID As String
    Private _Log As clsLog_Error
    Private oMenuObject As Object
    Private oItemObject As Object
    Private oSystemForms As Object
    Dim objFilters As SAPbouiCOM.EventFilters
    Dim objFilter As SAPbouiCOM.EventFilter

#Region "New"

    Public Sub New()
        MyBase.New()
        Try
            _Company = New SAPbobsCOM.Company
            _Utilities = New clsUtilities
            _Collection = New Hashtable(10, 0.5)
            _LookUpCollection = New Hashtable(10, 0.5)
            oSystemForms = New clsSystemForms
            _Log = New clsLog_Error

            SetApplication()

        Catch ex As Exception
            Throw ex
        End Try
    End Sub

#End Region

#Region "Public Properties"

    Public ReadOnly Property SBO_Application() As SAPbouiCOM.Application
        Get
            Return _SBO_Application
        End Get
    End Property

    Public ReadOnly Property Company() As SAPbobsCOM.Company
        Get
            Return _Company
        End Get
    End Property

    Public ReadOnly Property Utilities() As clsUtilities
        Get
            Return _Utilities
        End Get
    End Property

    Public ReadOnly Property Collection() As Hashtable
        Get
            Return _Collection
        End Get
    End Property

    Public ReadOnly Property LookUpCollection() As Hashtable
        Get
            Return _LookUpCollection
        End Get
    End Property

    Public ReadOnly Property Log() As clsLog_Error
        Get
            Return _Log
        End Get
    End Property


#Region "Filter"

    Public Sub SetFilter(ByVal Filters As SAPbouiCOM.EventFilters)
        oApplication.SBO_Application.SetFilter(Filters)
    End Sub

    Public Sub SetFilter()
        Try
            objFilters = New SAPbouiCOM.EventFilters()

            objFilter = objFilters.Add(SAPbouiCOM.BoEventTypes.et_FORM_LOAD)
            objFilter.AddEx(frm_SalImp) 'Import
           

            objFilter.AddEx("frm_PriceHistory")
            objFilter = objFilters.Add(SAPbouiCOM.BoEventTypes.et_MENU_CLICK)
            objFilter.AddEx(frm_SalImp) 'Import
           

           
            objFilter = objFilters.Add(SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED)
            objFilter.AddEx(frm_SalImp) 'Import

            objFilter = objFilters.Add(SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED)
            objFilter.AddEx(frm_SalImp) 'Import
          
            objFilter = objFilters.Add(SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST)
            objFilter.AddEx(frm_SalImp) 'Import
          

            objFilter = objFilters.Add(SAPbouiCOM.BoEventTypes.et_COMBO_SELECT)
            objFilter.AddEx(frm_SalImp) 'Import
            

            objFilter = objFilters.Add(SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK)
            objFilter.AddEx(frm_SalImp) 'Import
           


            SetFilter(objFilters)
        Catch ex As Exception
            Throw ex
        End Try

    End Sub
#End Region

#End Region

#Region "Data Event"
    Private Sub _SBO_Application_FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean) Handles _SBO_Application.FormDataEvent
        Try
            Select Case BusinessObjectInfo.FormTypeEx

            End Select
            If _Collection.ContainsKey(_FormUID) Then
                Dim objform As SAPbouiCOM.Form
                objform = oApplication.SBO_Application.Forms.ActiveForm()
                If 1 = 1 Then
                    oMenuObject = _Collection.Item(_FormUID)
                    oMenuObject.FormDataEvent(BusinessObjectInfo, BubbleEvent)
                End If

            End If
        Catch ex As Exception
            Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
#End Region

#Region "Menu Event"

    Private Sub SBO_Application_MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean) Handles _SBO_Application.MenuEvent
        Try
            If pVal.BeforeAction = False Then
                Select Case pVal.MenuUID
                    Case mnu_SalImp
                        oMenuObject = New clsImport_Sales
                        oMenuObject.MenuEvent(pVal, BubbleEvent)
                   
                End Select
            Else
                
            End If
        Catch ex As Exception
            Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        Finally
            oMenuObject = Nothing
        End Try
    End Sub

#End Region

#Region "Item Event"

    Private Sub SBO_Application_ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean) Handles _SBO_Application.ItemEvent
        Try
            _FormUID = FormUID
            If pVal.BeforeAction = False And pVal.EventType = SAPbouiCOM.BoEventTypes.et_FORM_LOAD Then
                Select Case pVal.FormTypeEx
                  
                    Case frm_SalImp
                        If Not _Collection.ContainsKey(FormUID) Then
                            oItemObject = New clsImport_Sales
                            oItemObject.FrmUID = FormUID
                            _Collection.Add(FormUID, oItemObject)
                        End If
                    
                End Select
            End If

            If pVal.FormTypeEx = "frm_ItemPrice" And pVal.ItemUID = "_12" And pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED And pVal.Before_Action = False Then
                Dim oform As SAPbouiCOM.Form
                oform = oApplication.SBO_Application.Forms.Item(FormUID)
                If oApplication.Utilities.getEditTextvalue(oform, "_10") <> "" Then
                    oApplication.Utilities.RefreshItemPriceHistory(oApplication.Utilities.getEditTextvalue(oform, "_11"), oApplication.Utilities.getEditTextvalue(oform, "5"))
                Else
                    oApplication.Utilities.RefreshItemPriceHistory("", oApplication.Utilities.getEditTextvalue(oform, "5"))
                End If

            End If


            Select Case pVal.EventType
                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                    Dim oform As SAPbouiCOM.Form
                    oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                    Dim oCFLEvento As SAPbouiCOM.IChooseFromListEvent
                    Dim oDataTable As SAPbouiCOM.DataTable
                    Dim strCustomer, strName, strCurrency, strWareHouse, strTaxCode, strAcctCode As String
                    Try
                        oCFLEvento = pVal
                        oDataTable = oCFLEvento.SelectedObjects
                        If pVal.ItemUID = "_10" And Not IsNothing(oDataTable) And pVal.Before_Action = False Then
                            strCustomer = oDataTable.GetValue("CardCode", 0)
                            strName = oDataTable.GetValue("CardName", 0)
                            strCurrency = oDataTable.GetValue("Currency", 0)
                            Try
                                CType(oform.Items.Item("_11").Specific, SAPbouiCOM.EditText).Value = strCustomer
                                CType(oform.Items.Item("_10").Specific, SAPbouiCOM.EditText).Value = strName
                            Catch ex As Exception

                            End Try
                        End If
                    Catch ex As Exception

                    End Try
            End Select
            If pVal.FormTypeEx = "frm_PriceHistory" And pVal.ItemUID = "7" Then
                Dim oform As SAPbouiCOM.Form
                oform = oApplication.SBO_Application.Forms.ActiveForm()
                oform.PaneLevel = 2

            End If

            If pVal.FormTypeEx = "frm_PriceHistory" And pVal.ItemUID = "6" Then
                Dim oform As SAPbouiCOM.Form
                oform = oApplication.SBO_Application.Forms.ActiveForm()
                oform.PaneLevel = 1

            End If
            If _Collection.ContainsKey(FormUID) Then
                oItemObject = _Collection.Item(FormUID)
                If oItemObject.IsLookUpOpen And pVal.BeforeAction = True Then
                    _SBO_Application.Forms.Item(oItemObject.LookUpFormUID).Select()
                    BubbleEvent = False
                    Exit Sub
                End If
                _Collection.Item(FormUID).ItemEvent(FormUID, pVal, BubbleEvent)
            End If
            If pVal.EventType = SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD And pVal.BeforeAction = False Then
                If _LookUpCollection.ContainsKey(FormUID) Then
                    oItemObject = _Collection.Item(_LookUpCollection.Item(FormUID))
                    If Not oItemObject Is Nothing Then
                        oItemObject.IsLookUpOpen = False
                    End If
                    _LookUpCollection.Remove(FormUID)
                End If

                If _Collection.ContainsKey(FormUID) Then
                    _Collection.Item(FormUID) = Nothing
                    _Collection.Remove(FormUID)
                End If
            End If
        Catch ex As Exception
            Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        Finally
            GC.WaitForPendingFinalizers()
            GC.Collect()
        End Try
    End Sub

#End Region

#Region "Right Click Event"

    Private Sub _SBO_Application_RightClickEvent(ByRef eventInfo As SAPbouiCOM.ContextMenuInfo, ByRef BubbleEvent As Boolean) Handles _SBO_Application.RightClickEvent
        Dim oform As SAPbouiCOM.Form
        oform = oApplication.SBO_Application.Forms.Item(eventInfo.FormUID)
        oform = oApplication.SBO_Application.Forms.Item(eventInfo.FormUID)
        'If eventInfo.FormUID = "RightClk" Then
        'If oform.TypeEx = frm_ITEM_MASTER Then
        '    If (eventInfo.BeforeAction = True) Then
        '        Dim oMenuItem As SAPbouiCOM.MenuItem
        '        Dim oMenus As SAPbouiCOM.Menus
        '        Try
        '            If oform.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
        '                Dim oCreationPackage As SAPbouiCOM.MenuCreationParams
        '                oCreationPackage = oApplication.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)
        '                oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING
        '                oCreationPackage.UniqueID = "mnu_IPrice"
        '                oCreationPackage.String = "Price Histrory"
        '                oCreationPackage.Enabled = True
        '                oMenuItem = oApplication.SBO_Application.Menus.Item("1280") 'Data'
        '                oMenus = oMenuItem.SubMenus
        '                Try
        '                    oMenus.AddEx(oCreationPackage)

        '                Catch ex As Exception

        '                End Try


        '            End If
        '        Catch ex As Exception
        '            MessageBox.Show(ex.Message)
        '        End Try
        '    Else
        '        Dim oMenuItem As SAPbouiCOM.MenuItem
        '        Dim oMenus As SAPbouiCOM.Menus
        '        Try
        '            If oform.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
        '                Try
        '                    oApplication.SBO_Application.Menus.RemoveEx("mnu_IPrice")
        '                Catch ex As Exception

        '                End Try

        '            End If
        '        Catch ex As Exception
        '            MessageBox.Show(ex.Message)
        '        End Try
        '   End If
        ' End If
    End Sub

#End Region

#Region "Application Event"

    Private Sub SBO_Application_AppEvent(ByVal EventType As SAPbouiCOM.BoAppEventTypes) Handles _SBO_Application.AppEvent
        Try
            Select Case EventType
                Case SAPbouiCOM.BoAppEventTypes.aet_ShutDown, SAPbouiCOM.BoAppEventTypes.aet_ServerTerminition, SAPbouiCOM.BoAppEventTypes.aet_CompanyChanged
                    _Utilities.AddRemoveMenus("RemoveMenus.xml")
                    CloseApp()
            End Select
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Termination Error", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly)
        End Try
    End Sub

#End Region

#Region "Close Application"

    Private Sub CloseApp()
        Try
            If Not _SBO_Application Is Nothing Then
                System.Runtime.InteropServices.Marshal.ReleaseComObject(_SBO_Application)
            End If

            If Not _Company Is Nothing Then
                If _Company.Connected Then
                    _Company.Disconnect()
                End If
                System.Runtime.InteropServices.Marshal.ReleaseComObject(_Company)
            End If

            _Utilities = Nothing
            _Collection = Nothing
            _LookUpCollection = Nothing

            ThreadClose.Sleep(10)
            System.Windows.Forms.Application.Exit()
        Catch ex As Exception
            Throw ex
        Finally
            oApplication = Nothing
            GC.WaitForPendingFinalizers()
            GC.Collect()
        End Try
    End Sub

#End Region

#Region "Set Application"

    Private Sub SetApplication()
        Dim SboGuiApi As SAPbouiCOM.SboGuiApi
        Dim sConnectionString As String

        Try
            If Environment.GetCommandLineArgs.Length > 1 Then
                sConnectionString = Environment.GetCommandLineArgs.GetValue(1)
                SboGuiApi = New SAPbouiCOM.SboGuiApi
                SboGuiApi.Connect(sConnectionString)
                _SBO_Application = SboGuiApi.GetApplication()
            Else
                Throw New Exception("Connection string missing.")
            End If

        Catch ex As Exception
            Throw ex
        Finally
            SboGuiApi = Nothing
        End Try
    End Sub

#End Region

#Region "Finalize"

    Protected Overrides Sub Finalize()
        Try
            MyBase.Finalize()
            '            CloseApp()

            oMenuObject = Nothing
            oItemObject = Nothing
            oSystemForms = Nothing

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Addon Termination Error", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly)
        Finally
            GC.WaitForPendingFinalizers()
            GC.Collect()
        End Try
    End Sub

#End Region

    Public Class WindowWrapper
        Implements System.Windows.Forms.IWin32Window
        Private _hwnd As IntPtr

        Public Sub New(ByVal handle As IntPtr)
            _hwnd = handle
        End Sub

        Public ReadOnly Property Handle() As System.IntPtr Implements System.Windows.Forms.IWin32Window.Handle
            Get
                Return _hwnd
            End Get
        End Property

    End Class

    Private Sub _SBO_Application_ServerInvokeCompletedEvent(ByRef b1iEventArgs As SAPbouiCOM.B1iEvent, ByRef BubbleEvent As Boolean) Handles _SBO_Application.ServerInvokeCompletedEvent

    End Sub
End Class
