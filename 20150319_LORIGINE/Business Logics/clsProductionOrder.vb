Imports System.IO
Imports System.Text.RegularExpressions
Imports System.Text
Imports System

Public Class clsProductionOrder
    Inherits clsBase
    Dim oDBDataSource As SAPbouiCOM.DBDataSource
    Private strQuery As String

    Public Sub New()
        MyBase.New()
    End Sub

#Region "Menu Event"
    Public Overrides Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        Try
            oForm = oApplication.SBO_Application.Forms.ActiveForm
            Select Case pVal.MenuUID
                Case mnu_FIRST, mnu_LAST, mnu_NEXT, mnu_PREVIOUS
                Case mnu_ADD
                Case mnu_GenerateSplit
                    oDBDataSource = oForm.DataSources.DBDataSources.Item(0)
                    If Not oDBDataSource.GetValue("DocEntry", 0).ToString = "" Then
                        If oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                            callSplipProductionList(oForm)
                        End If
                    End If
                    Dim oMenuItem As SAPbouiCOM.MenuItem
                    oMenuItem = oApplication.SBO_Application.Menus.Item("1280") 'Data'
                    If oMenuItem.SubMenus.Exists(pVal.MenuUID) Then
                        oApplication.SBO_Application.Menus.RemoveEx(pVal.MenuUID)
                    End If
            End Select
        Catch ex As Exception

        End Try
    End Sub
#End Region

#Region "Item Event"
    Public Overrides Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.FormTypeEx = frm_Production Then
                Select Case pVal.BeforeAction
                    Case True
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_CLICK

                        End Select
                    Case False
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                        End Select
                End Select
            End If
        Catch ex As Exception
            oForm = oApplication.SBO_Application.Forms.ActiveForm()
            oForm.Freeze(False)
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
#End Region

#Region "Data Events"
    Public Sub FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
        Try

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
#End Region

#Region "Right Click"
    Public Sub RightClickEvent(ByRef eventInfo As SAPbouiCOM.ContextMenuInfo, ByRef BubbleEvent As Boolean)
        oForm = oApplication.SBO_Application.Forms.Item(eventInfo.FormUID)
        If oForm.TypeEx = frm_Production Then
            If (eventInfo.BeforeAction = True) Then
                Dim oMenuItem As SAPbouiCOM.MenuItem
                Dim oMenus As SAPbouiCOM.Menus
                Try
                    oMenuItem = oApplication.SBO_Application.Menus.Item("1280") 'Data'
                    oDBDataSource = oForm.DataSources.DBDataSources.Item(0)
                    If oDBDataSource.GetValue("Status", 0).ToString = "P" And oDBDataSource.GetValue("U_Split", 0).ToString <> "Y" And oDBDataSource.GetValue("U_BaseProd", 0).ToString = "" Then
                        If Not oMenuItem.SubMenus.Exists(mnu_GenerateSplit) Then
                            Dim oCreationPackage As SAPbouiCOM.MenuCreationParams
                            oCreationPackage = oApplication.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)
                            oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING
                            oCreationPackage.UniqueID = mnu_GenerateSplit
                            oCreationPackage.String = "Split Production Order"
                            oCreationPackage.Enabled = True
                            oMenus = oMenuItem.SubMenus
                            oMenus.AddEx(oCreationPackage)
                        End If
                    ElseIf oDBDataSource.GetValue("Status", 0).ToString <> "P" Or oDBDataSource.GetValue("U_BaseProd", 0).ToString <> "" Then
                        If oMenuItem.SubMenus.Exists(mnu_GenerateSplit) Then
                            Try
                                oMenuItem.SubMenus.RemoveEx(mnu_GenerateSplit)
                            Catch ex As Exception

                            End Try

                        End If
                    End If
                Catch ex As Exception
                    MessageBox.Show(ex.Message)
                End Try
            End If
        End If
    End Sub
#End Region

#Region "Function"

    Private Sub initializeControls(ByVal oForm As SAPbouiCOM.Form)
        Try

        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub callSplipProductionList(ByVal oForm As SAPbouiCOM.Form)
        Try
            Dim objCommCharge As clsSplitList
            objCommCharge = New clsSplitList
            Dim strPONo As String = (oForm.Items.Item("18").Specific.value)
            oDBDataSource = oForm.DataSources.DBDataSources.Item(0)
            strPONo = oDBDataSource.GetValue("DocEntry", 0)

            Dim dblPlannedQty As Double = CDbl(oForm.Items.Item("12").Specific.value)
            Dim oRecordset As SAPbobsCOM.Recordset
            oRecordset = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            strQuery = "Select OrdrMulti,ItemCode,ItemName,U_MinPer From OITM Where ItemCode = '" + oApplication.Utilities.getEditTextvalue(oForm, 6) + "'"
            oRecordset.DoQuery(strQuery)
            If oRecordset.Fields.Item(0).Value <= 0 Then
                oApplication.Utilities.Message("Order Multiple Qty is not defined for this Item.", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Exit Sub
            End If
            objCommCharge.LoadForm(strPONo, dblPlannedQty)
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

#End Region

End Class
