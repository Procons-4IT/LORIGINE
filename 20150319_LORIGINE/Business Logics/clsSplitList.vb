Imports System.IO
Imports System.Text.RegularExpressions
Imports System.Text
Imports System

Public Class clsSplitList
    Inherits clsBase
    Private oGrid As SAPbouiCOM.Grid
    Private oDtSplitList As SAPbouiCOM.DataTable
    Private strQuery As String
    Dim oRecordSet As SAPbobsCOM.Recordset

    Public Sub New()
        MyBase.New()
    End Sub

    Public Sub LoadForm(ByVal strPOCode As String, ByVal dblPlnQty As Double)
        Try
            oForm = oApplication.Utilities.LoadForm(xml_SL, frm_SL)
            oForm = oApplication.SBO_Application.Forms.ActiveForm()
            oForm.Freeze(True)
            CType(oForm.Items.Item("4").Specific, SAPbouiCOM.EditText).Value = strPOCode
            CType(oForm.Items.Item("5").Specific, SAPbouiCOM.EditText).Value = CDbl(dblPlnQty)
            initialize(oForm, strPOCode, dblPlnQty)
            oForm.Freeze(False)
        Catch ex As Exception
            Throw ex
        Finally
            oForm.Freeze(False)
        End Try
    End Sub

#Region "Item Event"

    Public Overrides Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.FormTypeEx = frm_SL Then
                Select Case pVal.BeforeAction
                    Case True

                    Case False
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                            Case SAPbouiCOM.BoEventTypes.et_CLICK
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "btnSelect" Then
                                    oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                    SelectAll(oForm, True)
                                End If
                                If pVal.ItemUID = "btnClear" Then
                                    oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                    SelectAll(oForm, False)
                                End If
                                If pVal.ItemUID = "_1" Then
                                    Dim blnVStatus As Boolean = validate(pVal, oForm)
                                    If blnVStatus = False Then
                                        BubbleEvent = False
                                        Exit Sub
                                    End If
                                    Dim _retVal As Integer '
                                    If oApplication.SBO_Application.MessageBox("Do you want to create the new productions orders?", , "Continue", "Cancel") = 2 Then
                                        Exit Sub
                                    Else
                                        _retVal = 1
                                    End If
                                    '  = oApplication.SBO_Application.MessageBox("Do you want to Create Split production Automatically ?...", "2", "Yes", "No")
                                    If _retVal = 1 And blnVStatus Then
                                        If (addProductionOrder(pVal, oForm)) Then
                                            _retVal = oApplication.SBO_Application.MessageBox("Do you want to Cancel the Source production Automatically ?...", "2", "Yes", "No")
                                            If _retVal = 1 Then
                                                cancelProdOrder(pVal, oForm)
                                            End If
                                            oForm.Close()
                                        End If
                                    Else
                                        oApplication.SBO_Application.SetStatusBarMessage("Cannot Create if Split Qty is Less than Min Percentage...", SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                                        BubbleEvent = False
                                    End If
                                End If
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

#Region "Function"

    Private Sub initialize(ByVal oForm As SAPbouiCOM.Form, ByVal strPOCode As String, ByVal dblPlnQty As Double)
        Try
            oGrid = oForm.Items.Item("3").Specific
            oForm.DataSources.DataTables.Add("dtSplitList")
            oDtSplitList = oForm.DataSources.DataTables.Item(0)
            oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            strQuery = "Select T0.DueDate,T0.Status,ISNULL(U_Split,'N') As U_Split,ItemCode From OWOR T0 Where DocEntry =  '" + strPOCode + "'"
            oRecordSet.DoQuery(strQuery)
            If Not oRecordSet.EoF Then
                Dim strItemCode As String = oRecordSet.Fields.Item("ItemCode").Value
                Dim dblMinOrder As Double = 0
                Dim dblAppOrderQty As Double = 0
                Dim dblMinPer As Double = 0
                Dim strStatus As String = oRecordSet.Fields.Item("Status").Value
                Dim dtduedate As Date = oRecordSet.Fields.Item("DueDate").Value
                If oRecordSet.Fields.Item("U_Split").Value = "N" Then
                    strQuery = " Select T0.DocNum,T0.DueDate,T0.Status,T1.ItemCode,T1.ItemName,PlannedQty,'N' As 'Check','N' As 'Success',T0.DueDate 'Due' From OWOR T0 JOIN OITM T1 ON T0.ItemCode = T1.ItemCode "
                    strQuery += " Where 1 = 2 "
                    oDtSplitList.ExecuteQuery(strQuery)
                    strQuery = "Select OrdrMulti,ItemCode,ItemName,isnull(U_MinPer,100) 'U_MinPer' From OITM Where ItemCode = '" + strItemCode + "'"
                    oRecordSet.DoQuery(strQuery)
                    If Not oRecordSet.EoF Then
                        dblMinOrder = CDbl(oRecordSet.Fields.Item(0).Value)
                        dblAppOrderQty = dblPlnQty
                        dblMinPer = CDbl(oRecordSet.Fields.Item("U_MinPer").Value)
                        If dblMinOrder > 0 Then
                            Dim intRow As Integer = 0
                            If dblAppOrderQty > dblMinOrder Then
                                While dblAppOrderQty > dblMinOrder
                                    oDtSplitList.Rows.Add()
                                    oDtSplitList.SetValue("DocNum", intRow, strPOCode)
                                    oDtSplitList.SetValue("DueDate", intRow, dtduedate)
                                    oDtSplitList.SetValue("Due", intRow, dtduedate)
                                    oDtSplitList.SetValue("Status", intRow, strStatus)
                                    oDtSplitList.SetValue("ItemCode", intRow, oRecordSet.Fields.Item("ItemCode").Value)
                                    oDtSplitList.SetValue("ItemName", intRow, oRecordSet.Fields.Item("ItemName").Value)
                                    oDtSplitList.SetValue("PlannedQty", intRow, dblMinOrder)
                                    oDtSplitList.SetValue("Check", intRow, "Y")
                                    oDtSplitList.SetValue("Success", intRow, "Y")
                                    intRow = intRow + 1
                                    dblAppOrderQty = dblAppOrderQty - dblMinOrder
                                End While
                                If dblAppOrderQty <> 0 Then
                                    If dblAppOrderQty >= (dblMinOrder * (dblMinPer / 100)) Then
                                        oDtSplitList.Rows.Add()
                                        oDtSplitList.SetValue("DocNum", intRow, strPOCode)
                                        oDtSplitList.SetValue("DueDate", intRow, dtduedate)
                                        oDtSplitList.SetValue("Due", intRow, dtduedate)
                                        oDtSplitList.SetValue("Status", intRow, strStatus)
                                        oDtSplitList.SetValue("ItemCode", intRow, oRecordSet.Fields.Item("ItemCode").Value)
                                        oDtSplitList.SetValue("ItemName", intRow, oRecordSet.Fields.Item("ItemName").Value)
                                        oDtSplitList.SetValue("PlannedQty", intRow, dblAppOrderQty)
                                        oDtSplitList.SetValue("Check", intRow, "Y")
                                        oDtSplitList.SetValue("Success", intRow, "Y")
                                        intRow = intRow + 1
                                        dblAppOrderQty = dblAppOrderQty - dblAppOrderQty
                                    ElseIf (dblAppOrderQty < (dblMinOrder * (dblMinPer / 100))) Then
                                        oDtSplitList.Rows.Add()
                                        oDtSplitList.SetValue("DocNum", intRow, strPOCode)
                                        oDtSplitList.SetValue("DueDate", intRow, dtduedate)
                                        oDtSplitList.SetValue("Due", intRow, dtduedate)
                                        oDtSplitList.SetValue("Status", intRow, strStatus)
                                        oDtSplitList.SetValue("ItemCode", intRow, oRecordSet.Fields.Item("ItemCode").Value)
                                        oDtSplitList.SetValue("ItemName", intRow, oRecordSet.Fields.Item("ItemName").Value)
                                        oDtSplitList.SetValue("PlannedQty", intRow, dblAppOrderQty)
                                        oDtSplitList.SetValue("Check", intRow, "Y")
                                        oDtSplitList.SetValue("Success", intRow, "N")
                                        intRow = intRow + 1
                                        dblAppOrderQty = dblAppOrderQty - dblAppOrderQty
                                    End If
                                End If
                            End If
                        End If
                    End If
                    oGrid.DataTable = oDtSplitList
                    oGrid.Columns.Item("Check").Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox
                    oGrid.Columns.Item("Check").Editable = True
                Else
                    Dim strPONo As String = CType(oForm.Items.Item("4").Specific, SAPbouiCOM.EditText).Value
                    strQuery = " Select T0.DocNum,T1.ItemCode,T1.ItemName,PlannedQty,'Y' As 'Check','Y' As 'Success' From OWOR T0 JOIN OITM T1 ON T0.ItemCode = T1.ItemCode "
                    strQuery += " Where U_BaseProd = '" + strPOCode + "'"
                    oDtSplitList.ExecuteQuery(strQuery)
                    oGrid.DataTable = oDtSplitList
                    oGrid.Columns.Item("Check").Visible = False
                    oForm.Items.Item("_1").Enabled = False
                End If
            End If

            'Format
            oGrid.Columns.Item("DocNum").TitleObject.Caption = "Production Order No."
            oGrid.Columns.Item("DueDate").TitleObject.Caption = "Due Date"
            oGrid.Columns.Item("Status").TitleObject.Caption = "Status"
            oGrid.Columns.Item("Status").Type = SAPbouiCOM.BoGridColumnType.gct_ComboBox
            Dim ocombo As SAPbouiCOM.ComboBoxColumn
            ocombo = oGrid.Columns.Item("Status")
            ocombo.ValidValues.Add("P", "Planned")
            ocombo.ValidValues.Add("R", "Release")
            ocombo.DisplayType = SAPbouiCOM.BoComboDisplayType.cdt_both
            ocombo.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly
            oGrid.Columns.Item("Status").Editable = True
            oGrid.Columns.Item("DueDate").Editable = True
            oGrid.Columns.Item("Due").Visible = False

            oGrid.Columns.Item("ItemCode").TitleObject.Caption = "Finished Item"
            oGrid.Columns.Item("ItemName").TitleObject.Caption = "Finished Name"
            oGrid.Columns.Item("PlannedQty").TitleObject.Caption = "Planned Qty"
            oGrid.Columns.Item("Check").TitleObject.Caption = "Select"
            oGrid.Columns.Item("PlannedQty").RightJustified = True

            oGrid.Columns.Item("DocNum").Editable = False
            oGrid.Columns.Item("ItemCode").Editable = False
            oGrid.Columns.Item("ItemName").Editable = False
            oGrid.Columns.Item("PlannedQty").Editable = True
            oGrid.Columns.Item("Success").Visible = False
            For intRow As Integer = 0 To oGrid.DataTable.Rows.Count - 1
                oGrid.RowHeaders.SetText(intRow, intRow + 1)
            Next

        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub SelectAll(ByVal aForm As SAPbouiCOM.Form, ByVal aflag As Boolean)
        oGrid = aForm.Items.Item("3").Specific
        aForm.Freeze(True)
        Dim oCheckBox As SAPbouiCOM.CheckBoxColumn
        For intRow As Integer = 0 To oGrid.DataTable.Rows.Count - 1
            oCheckBox = oGrid.Columns.Item("Check")
            oCheckBox.Check(intRow, aflag)
        Next
        aForm.Freeze(False)
    End Sub




    Public Function addProductionOrder(ByRef pVal As SAPbouiCOM.ItemEvent, ByRef objForm As SAPbouiCOM.Form) As Boolean
        Dim _retVal As Boolean = False
        Dim objSourceProductionOrder As SAPbobsCOM.ProductionOrders
        Dim objProductionOrder As SAPbobsCOM.ProductionOrders
        Dim objDesignation As SAPbobsCOM.ProductionOrders
        objProductionOrder = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oProductionOrders)
        objDesignation = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oProductionOrders)

        oGrid = objForm.Items.Item("3").Specific
        Dim blnAddStatus As Boolean = True
        If oApplication.Company.InTransaction Then
            oApplication.Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
        End If
        oApplication.Company.StartTransaction()
        Dim strPONo As String = CType(objForm.Items.Item("4").Specific, SAPbouiCOM.EditText).Value
        objSourceProductionOrder = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oProductionOrders)
        If objSourceProductionOrder.GetByKey(strPONo) Then
            Dim strXML As String = String.Empty
            Dim strPath As String = String.Empty

            strXML = objSourceProductionOrder.GetAsXML()
            For i As Integer = 0 To oGrid.Rows.Count - 1
                oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                If oGrid.DataTable.GetValue("Check", i) = "Y" And CDbl(oGrid.DataTable.GetValue("PlannedQty", i)) > 0 Then

                    Try
                        objDesignation = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oProductionOrders)

                        objDesignation.CustomerCode = objSourceProductionOrder.CustomerCode
                        objDesignation.DistributionRule = objSourceProductionOrder.DistributionRule
                        objDesignation.DistributionRule2 = objSourceProductionOrder.DistributionRule2
                        objDesignation.DistributionRule3 = objSourceProductionOrder.DistributionRule3
                        objDesignation.DistributionRule4 = objSourceProductionOrder.DistributionRule4
                        objDesignation.DistributionRule5 = objSourceProductionOrder.DistributionRule5
                        objDesignation.DueDate = (oGrid.DataTable.GetValue("DueDate", i))
                        objDesignation.ItemNo = objSourceProductionOrder.ItemNo
                        objDesignation.JournalRemarks = objSourceProductionOrder.JournalRemarks
                        objDesignation.PlannedQuantity = CDbl(oGrid.DataTable.GetValue("PlannedQty", i))
                        objDesignation.PostingDate = objSourceProductionOrder.PostingDate
                        objDesignation.ProductionOrderOrigin = objSourceProductionOrder.ProductionOrderOrigin
                        If objSourceProductionOrder.ProductionOrderOrigin = SAPbobsCOM.BoProductionOrderOriginEnum.bopooSalesOrder Then
                            objDesignation.ProductionOrderOriginEntry = objSourceProductionOrder.ProductionOrderOriginEntry
                        End If


                        objDesignation.ProductionOrderType = objSourceProductionOrder.ProductionOrderType
                        objDesignation.Project = objSourceProductionOrder.Project
                        objDesignation.Series = objSourceProductionOrder.Series
                        objDesignation.Warehouse = objSourceProductionOrder.Warehouse
                        objDesignation.UserFields.Fields.Item("U_BaseProd").Value = objSourceProductionOrder.DocumentNumber.ToString
                        For intRow As Integer = 0 To objSourceProductionOrder.Lines.Count - 1
                            objSourceProductionOrder.Lines.SetCurrentLine(intRow)
                            If intRow > 0 Then
                                objDesignation.Lines.Add()
                            End If
                            objDesignation.Lines.SetCurrentLine(intRow)
                            objDesignation.Lines.BaseQuantity = objSourceProductionOrder.Lines.BaseQuantity
                            objDesignation.Lines.DistributionRule = objSourceProductionOrder.Lines.DistributionRule
                            objDesignation.Lines.DistributionRule2 = objSourceProductionOrder.Lines.DistributionRule2
                            objDesignation.Lines.DistributionRule3 = objSourceProductionOrder.Lines.DistributionRule3
                            objDesignation.Lines.DistributionRule4 = objSourceProductionOrder.Lines.DistributionRule4
                            objDesignation.Lines.DistributionRule5 = objSourceProductionOrder.Lines.DistributionRule5
                            objDesignation.Lines.ItemNo = objSourceProductionOrder.Lines.ItemNo
                            objDesignation.Lines.ProductionOrderIssueType = objSourceProductionOrder.Lines.ProductionOrderIssueType
                            If objSourceProductionOrder.Lines.Project <> "" Then
                                objDesignation.Lines.Project = objSourceProductionOrder.Lines.Project
                            End If


                            objDesignation.Lines.PlannedQuantity = CDbl(oGrid.DataTable.GetValue("PlannedQty", i))
                            For intLoop As Integer = 0 To objSourceProductionOrder.Lines.UserFields.Fields.Count - 1
                                objDesignation.Lines.UserFields.Fields.Item(intLoop).Value = objSourceProductionOrder.Lines.UserFields.Fields.Item(intLoop).Value

                            Next
                        Next
                        Dim intRetStatus As Integer = objDesignation.Add()
                        If intRetStatus <> 0 Then
                            oApplication.SBO_Application.MessageBox(oApplication.Company.GetLastErrorDescription())
                            blnAddStatus = False
                        Else
                            Dim strDocEntry As String
                            oApplication.Company.GetNewObjectCode(strDocEntry)
                            'If File.Exists(strPath) Then
                            '    File.Delete(strPath)
                            'End If
                            If 1 = 1 Then
                                objProductionOrder.GetByKey(CInt(strDocEntry))
                                objProductionOrder.DueDate = oGrid.DataTable.GetValue("DueDate", i)
                                Dim oCombo As SAPbouiCOM.ComboBoxColumn
                                oCombo = oGrid.Columns.Item("Status")
                                Try

                              
                                If oCombo.GetSelectedValue(i).Value = "R" Then
                                    objProductionOrder.ProductionOrderStatus = SAPbobsCOM.BoProductionOrderStatusEnum.boposReleased
                                Else
                                    '   objProductionOrder.ProductionOrderStatus = SAPbobsCOM.BoProductionOrderStatusEnum.boposPlanned
                                End If

                                    objProductionOrder.Update()
                                Catch ex As Exception
                                    MsgBox(ex.Message)
                                End Try
                            End If
                        End If

                        'Dim oXMLDocument As New Xml.XmlDocument
                        'Dim oXMLElement As Xml.XmlElement
                        'Dim oXMLHeaderNodes As Xml.XmlNode
                        'Dim oXMLRowsNodes As Xml.XmlNode
                        'Dim oXMLDocEntry As Xml.XmlNode
                        'Dim oXMLDocNum As Xml.XmlNode

                        'Dim oXMLDueDate As Xml.XmlNode
                        'Dim oXMLStatus As Xml.XmlNode
                        'Dim oXMLPlnQty As Xml.XmlNode
                        'Dim oXMLRDocEntry As Xml.XmlNode
                        'Dim oXMLRPlanned As Xml.XmlNode
                        'Dim oXMLOriginAbs As Xml.XmlNode
                        'Dim oXMLOriginNum As Xml.XmlNode
                        'Dim oXMLBaseProd As Xml.XmlNode


                        'oXMLDocument.LoadXml(strXML)
                        'oXMLElement = oXMLDocument.DocumentElement

                        'oXMLHeaderNodes = oXMLElement.ChildNodes.Item(0).ChildNodes(1)
                        'oXMLDocEntry = oXMLDocument.SelectSingleNode("/BOM/BO/OWOR/row/DocEntry")
                        'oXMLDocNum = oXMLDocument.SelectSingleNode("/BOM/BO/OWOR/row/DocNum")

                        'oXMLDueDate = oXMLDocument.SelectSingleNode("/BOM/BO/OWOR/row/DueDate")
                        'oXMLStatus = oXMLDocument.SelectSingleNode("/BOM/BO/OWOR/row/Status")

                        'oXMLPlnQty = oXMLDocument.SelectSingleNode("/BOM/BO/OWOR/row/PlannedQty")
                        'oXMLOriginAbs = oXMLDocument.SelectSingleNode("/BOM/BO/OWOR/row/OriginAbs")
                        'oXMLOriginNum = oXMLDocument.SelectSingleNode("/BOM/BO/OWOR/row/OriginNum")
                        'oXMLBaseProd = oXMLDocument.SelectSingleNode("/BOM/BO/OWOR/row/U_BaseProd")

                        'If IsNothing(oXMLDocEntry) Then
                        '    oXMLDocEntry = oXMLDocument.SelectSingleNode("/BOM/BO/ProductionOrders/row/AbsoluteEntry")
                        'End If

                        'If IsNothing(oXMLDocNum) Then
                        '    oXMLDocNum = oXMLDocument.SelectSingleNode("/BOM/BO/ProductionOrders/row/DocumentNumber")
                        'End If

                        'If IsNothing(oXMLDueDate) Then
                        '    oXMLDueDate = oXMLDocument.SelectSingleNode("/BOM/BO/ProductionOrders/row/DueDate")
                        'End If

                        'If IsNothing(oXMLStatus) Then
                        '    oXMLStatus = oXMLDocument.SelectSingleNode("/BOM/BO/ProductionOrders/row/ProductionOrderStatus")
                        'End If

                        'If IsNothing(oXMLPlnQty) Then
                        '    oXMLPlnQty = oXMLDocument.SelectSingleNode("/BOM/BO/ProductionOrders/row/PlannedQuantity")
                        'End If

                        'If IsNothing(oXMLOriginAbs) Then
                        '    oXMLOriginAbs = oXMLDocument.SelectSingleNode("/BOM/BO/ProductionOrders/row/ProductionOrderOriginEntry")
                        'End If

                        'If IsNothing(oXMLOriginNum) Then
                        '    oXMLOriginNum = oXMLDocument.SelectSingleNode("/BOM/BO/ProductionOrders/row/ProductionOrderOriginNumber")
                        'End If

                        'If IsNothing(oXMLBaseProd) Then
                        '    oXMLBaseProd = oXMLDocument.SelectSingleNode("/BOM/BO/ProductionOrders/row/U_BaseProd")
                        'End If



                        'oXMLDocEntry.InnerText = String.Empty
                        'oXMLDocNum.InnerText = String.Empty
                        'oXMLPlnQty.InnerText = CDbl(oGrid.DataTable.GetValue("PlannedQty", i))
                        '' oXMLDueDate.InnerText = (oGrid.DataTable.GetValue("DueDate", i))
                        ''                        oXMLPlnQty.InnerText = CDbl(oGrid.DataTable.GetValue("PlannedQty", i))


                        'If oXMLOriginAbs.InnerText = "0" Or oXMLOriginAbs.InnerText = "" Then
                        '    oXMLOriginAbs.ParentNode.RemoveChild(oXMLOriginAbs)
                        'End If
                        'If oXMLOriginNum.InnerText = "0" Or oXMLOriginNum.InnerText = "" Then
                        '    oXMLOriginNum.ParentNode.RemoveChild(oXMLOriginNum)
                        'End If

                        'oXMLBaseProd.InnerText = strPONo

                        ''Clearing All DocEntry
                        'oXMLRowsNodes = oXMLElement.ChildNodes.Item(0).ChildNodes(2)
                        'For index As Integer = 0 To oXMLRowsNodes.ChildNodes.Count - 1
                        '    oXMLRDocEntry = oXMLRowsNodes.ChildNodes(index).ChildNodes(0)
                        '    oXMLRDocEntry.InnerText = ""
                        '    'Planned Qty
                        '    oXMLRPlanned = oXMLRowsNodes.ChildNodes(index).ChildNodes(4)
                        '    oXMLRPlanned.InnerText = CDbl(oGrid.DataTable.GetValue("PlannedQty", i))
                        'Next

                        'strPath = strPONo.ToString() + "_" + i.ToString() + ".xml"
                        'oXMLDocument.Save(strPath)
                        'oApplication.Company.XmlExportType = SAPbobsCOM.BoXmlExportTypes.xet_NodesAsProperties
                        'objProductionOrder = oApplication.Company.GetBusinessObjectFromXML(strPath, 0)
                        'Dim intRetStatus As Integer = objProductionOrder.Add()
                        'If intRetStatus <> 0 Then
                        '    oApplication.SBO_Application.MessageBox(oApplication.Company.GetLastErrorDescription())
                        '    blnAddStatus = False
                        'Else
                        '    Dim strDocEntry As String
                        '    oApplication.Company.GetNewObjectCode(strDocEntry)
                        '    If File.Exists(strPath) Then
                        '        File.Delete(strPath)
                        '    End If
                        '    If 1 = 1 Then
                        '        objProductionOrder.GetByKey(CInt(strDocEntry))
                        '        objProductionOrder.DueDate = oGrid.DataTable.GetValue("DueDate", i)
                        '        objProductionOrder.ProductionOrderStatus = SAPbobsCOM.BoProductionOrderStatusEnum.boposReleased
                        '        objProductionOrder.Update()
                        '    End If
                        'End If
                    Catch ex As Exception
                        oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        blnAddStatus = False
                    End Try
                End If
            Next

            If blnAddStatus Then
                '  objProductionOrder.GetByKey(CInt(strPONo))
                objSourceProductionOrder.UserFields.Fields.Item("U_Split").Value = "Y"
                Dim intRetStatus As Integer = objSourceProductionOrder.Update()
                If intRetStatus <> 0 Then
                    blnAddStatus = False
                    oApplication.SBO_Application.MessageBox(oApplication.Company.GetLastErrorDescription())
                End If
            End If

        End If

        If blnAddStatus Then
            If oApplication.Company.InTransaction Then
                oApplication.Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
            End If
            oApplication.SBO_Application.MessageBox("Split Production Orders Created Successfully....")
        Else
            If oApplication.Company.InTransaction Then
                oApplication.Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
            End If
        End If
        Return blnAddStatus
    End Function

    Public Sub cancelProdOrder(ByRef pVal As SAPbouiCOM.ItemEvent, ByRef objForm As SAPbouiCOM.Form)
        Try
            Dim objSourceProductionOrder As SAPbobsCOM.ProductionOrders
            Dim strPONo As String = CType(oForm.Items.Item("4").Specific, SAPbouiCOM.EditText).Value
            objSourceProductionOrder = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oProductionOrders)
            If objSourceProductionOrder.GetByKey(strPONo) Then
                Dim intStatua As Integer = objSourceProductionOrder.Cancel()
                If intStatua = 0 Then
                    oApplication.SBO_Application.MessageBox("Source Production Orders Cancelled Successfully....")
                End If
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Public Function validate(ByRef pVal As SAPbouiCOM.ItemEvent, ByRef objForm As SAPbouiCOM.Form)
        Dim _retVal As Boolean = True
        oGrid = objForm.Items.Item("3").Specific
        Try
            For i As Integer = 0 To oGrid.Rows.Count - 1
                If oGrid.DataTable.GetValue("Check", i) = "Y" And oGrid.DataTable.GetValue("Success", i) = "N" Then
                    oApplication.Utilities.Message("Cannot Create if Split Qty is Less than Min Percentage", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Return False
                End If
                If oGrid.DataTable.GetValue("Check", i) = "Y" Then
                    If oGrid.DataTable.GetValue("DueDate", i) < oGrid.DataTable.GetValue("Due", i) Then
                        oApplication.Utilities.Message("Due date should be greater than or equal to Source produce order due date", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Return False
                    End If
                End If

            Next
            Return True
        Catch ex As Exception
            Throw ex
        End Try
        Return _retVal
    End Function
#End Region

End Class
