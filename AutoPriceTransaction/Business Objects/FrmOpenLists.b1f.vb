Option Strict Off
Option Explicit On

Imports SAPbouiCOM.Framework

Namespace AutoPriceTransaction
    <FormAttribute("OpenList", "Business Objects/FrmOpenLists.b1f")>
    Friend Class FrmOpenLists
        Inherits UserFormBase
        Public WithEvents objform, objformUDF, objformNew As SAPbouiCOM.Form
        Public WithEvents objText As SAPbouiCOM.EditText
        Private WithEvents objMatrix As SAPbouiCOM.Matrix
        Private WithEvents odbdsDetails As SAPbouiCOM.DBDataSource
        Private WithEvents objDTable As SAPbouiCOM.DataTable
        Public UDFFormID, StrQuery As String
        Dim objRS As SAPbobsCOM.Recordset
        Dim GetEntry() As String
        Dim ColumnNum As Integer

        Public Sub New()
        End Sub

        Public Overrides Sub OnInitializeComponent()
            Me.Button0 = CType(Me.GetItem("101").Specific, SAPbouiCOM.Button)
            Me.Button1 = CType(Me.GetItem("2").Specific, SAPbouiCOM.Button)
            Me.Matrix0 = CType(Me.GetItem("mtxdata").Specific, SAPbouiCOM.Matrix)
            Me.StaticText0 = CType(Me.GetItem("lblfind").Specific, SAPbouiCOM.StaticText)
            Me.EditText0 = CType(Me.GetItem("tfind").Specific, SAPbouiCOM.EditText)
            Me.OnCustomInitialize()

        End Sub

        Public Overrides Sub OnInitializeFormEvents()
            AddHandler CloseBefore, AddressOf Me.Form_CloseBefore

        End Sub

        Private Sub OnCustomInitialize()
            Try
                objform = objaddon.objapplication.Forms.GetForm("OpenList", 0)
                bModal = True
                objform.Settings.Enabled = True
            Catch ex As Exception
                objform.Freeze(False)
            End Try
        End Sub

#Region "Fields"
        Private WithEvents Button0 As SAPbouiCOM.Button
        Private WithEvents Button1 As SAPbouiCOM.Button
        Private WithEvents Matrix0 As SAPbouiCOM.Matrix
        Private WithEvents StaticText0 As SAPbouiCOM.StaticText
        Private WithEvents EditText0 As SAPbouiCOM.EditText
#End Region

        Private Sub Button0_ClickAfter(sboObject As Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles Button0.ClickAfter
            Try
                GetEntry = GetSelectedEntry()
                objform.Close()
                If GetEntry(0) = "" Then Exit Sub
                If DType = "GR" Then
                    Load_GoodsIssue_GoodsReceipt(GetEntry(0), GetEntry(1))
                ElseIf DType = "AR" Then
                    Load_Delivery_ARInvoice(GetEntry(0), GetEntry(1))
                ElseIf DType = "GRPO" Then
                    If VendorCode = "" Then Exit Sub
                    Load_ARInvoice_GRPO(GetEntry(0), GetEntry(1), VendorCode)
                Else
                    Exit Sub
                End If

            Catch ex As Exception

            End Try

        End Sub

        Private Sub Button0_ClickBefore(sboObject As Object, pVal As SAPbouiCOM.SBOItemEventArg, ByRef BubbleEvent As Boolean) Handles Button0.ClickBefore
            Try

            Catch ex As Exception

            End Try

        End Sub

        Public Sub Load_Data(ByVal Query As String, ByVal DocType As String)
            Try
                If objform.DataSources.DataTables.Count.Equals(0) Then
                    objform.DataSources.DataTables.Add("DT_Data")
                Else
                    objform.DataSources.DataTables.Item("DT_Data").Clear()
                End If
                If DocType = "GR" Then
                    objform.Title = "Goods Issue Open Entries"
                ElseIf DocType = "AR" Then
                    objform.Title = "Delivery Open Entries"
                ElseIf DocType = "GRPO" Then
                    objform.Title = "A/R Invoice Open Entries"
                Else
                    objform.Title = "Open Entries"
                End If

                objform.Freeze(True)
                objDTable = objform.DataSources.DataTables.Item("DT_Data")
                objDTable.Clear()
                objDTable.ExecuteQuery(Query)
                objRS = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                objRS.DoQuery(Query)
                If objRS.RecordCount = 0 Then
                    objaddon.objapplication.StatusBar.SetText("No Entries Found...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    Exit Sub
                End If
                Matrix0.Clear()
                Matrix0.LoadFromDataSourceEx()
                If Matrix0.VisualRowCount > 0 Then
                    Matrix0.SelectRow(1, True, False)
                    objaddon.objapplication.StatusBar.SetText("Successfully Loaded Entries...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                Else
                    'objaddon.objapplication.StatusBar.SetText("No Entries Found...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                End If
                Matrix0.AutoResizeColumns()
                objform.Settings.Enabled = True
                'objDTable = Nothing
                'objform.Freeze(False)
            Catch ex As Exception
                'objform.Freeze(False)
            Finally
                objform.Freeze(False)
            End Try
        End Sub

        Private Sub Matrix0_DoubleClickAfter(sboObject As Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles Matrix0.DoubleClickAfter
            Try
                If pVal.ActionSuccess = True And pVal.Row <> 0 Then
                    GetEntry = GetSelectedEntry()
                    objform.Close()
                    If DType = "GR" Then
                        Load_GoodsIssue_GoodsReceipt(GetEntry(0), GetEntry(1))
                    ElseIf DType = "AR" Then
                        Load_Delivery_ARInvoice(GetEntry(0), GetEntry(1))
                    ElseIf DType = "GRPO" Then
                        If VendorCode = "" Then Exit Sub
                        Load_ARInvoice_GRPO(GetEntry(0), GetEntry(1), VendorCode)
                    Else
                        Exit Sub
                    End If
                Else
                    Dim colname As String = Matrix0.Columns.Item(pVal.ColUID).TitleObject.Caption
                    For i As Integer = 0 To Matrix0.Columns.Count - 1
                        If Matrix0.Columns.Item(i).TitleObject.Caption = colname Then
                            ColumnNum = i
                            Exit For
                        End If
                    Next

                End If
            Catch ex As Exception

            End Try

        End Sub

        Private Sub EditText0_KeyDownAfter(sboObject As Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles EditText0.KeyDownAfter
            Try
                If pVal.CharPressed <> 9 Then Exit Sub
                Dim FindString As String
                FindString = EditText0.Value
                Dim Flag As Boolean = False


                For j As Integer = 1 To Matrix0.VisualRowCount
                    Dim vv As String = Trim(Matrix0.Columns.Item(ColumnNum).Cells.Item(j).Specific.String)
                    If Trim(Matrix0.Columns.Item(ColumnNum).Cells.Item(j).Specific.String) Like FindString & "*" Or Trim(Matrix0.Columns.Item(ColumnNum).Cells.Item(j).Specific.String) = FindString Then
                        Matrix0.SelectRow(j, True, False)
                        Exit For
                    End If
                Next
            Catch ex As Exception

            End Try

        End Sub

        Private Sub Matrix0_ClickAfter(sboObject As Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles Matrix0.ClickAfter
            Try
                If pVal.Row <> 0 Then
                    Matrix0.SelectRow(pVal.Row, True, False)
                Else
                    Matrix0.Columns.Item(pVal.ColUID).TitleObject.Sortable = True
                End If
            Catch ex As Exception
            End Try
        End Sub

        Private Function GetSelectedEntry()
            Dim DocEntry As String = "", DocNum As String = ""
            Try
                For i As Integer = 1 To Matrix0.VisualRowCount
                    If Matrix0.IsRowSelected(i) Then
                        DocEntry = Matrix0.Columns.Item("docentry").Cells.Item(i).Specific.String
                        DocNum = Matrix0.Columns.Item("docnum").Cells.Item(i).Specific.String
                        Exit For
                    End If
                Next
                Return {DocEntry, DocNum}
            Catch ex As Exception
                Return ""
            End Try
        End Function

        Private Sub Form_CloseBefore(pVal As SAPbouiCOM.SBOItemEventArg, ByRef BubbleEvent As Boolean)
            Try
                If bModal Then
                    bModal = False
                End If
            Catch ex As Exception

            End Try

        End Sub

        Private Sub Load_GoodsIssue_GoodsReceipt(ByVal DocEntry As String, ByVal DocNum As String)
            'Dim FormID, series As String
            Dim objRS1 As SAPbobsCOM.Recordset
            Dim objGRform As SAPbouiCOM.Form
            Dim objGRMatrix As SAPbouiCOM.Matrix
            Dim objCombo As SAPbouiCOM.ComboBox
            Dim Row As Integer = 0
            Dim FormType As Integer
            'Dim m_oProgBar As SAPbouiCOM.ProgressBar
            Try
                If DocEntry = "" Then
                    Exit Sub
                End If
                'GRPODocEntry = DocEntry
                'objaddon.objapplication.SetStatusBarMessage("Goods Issue Loading to Goods Receipt Please Wait...", SAPbouiCOM.BoMessageTime.bmt_Short, False)
                FormType = objaddon.objapplication.Forms.ActiveForm.TypeCount
                objGRform = objaddon.objapplication.Forms.GetForm("721", FormType)
                objGRMatrix = objGRform.Items.Item("13").Specific
                objRS1 = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                'objRS = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                'objaddon.objapplication.SetStatusBarMessage("Goods Issue Loading to Goods Receipt Please wait... DocumentNumber-> " & DocNum, SAPbouiCOM.BoMessageTime.bmt_Medium, False)

                StrQuery = "Select T0.""BPLId"",T0.""U_TRANSTYPE"",T0.""U_TOWHS"" as ""Whse"",T0.""Comments"",T1.""DocEntry"",T1.""U_MRP"",T1.""Project"",T1.""LineNum"",T1.""ItemCode"",T1.""Quantity"",T1.""WhsCode"",T1.""AcctCode"",T1.""StockPrice"",T1.""LocCode"""
                StrQuery += vbCrLf + " from OIGE T0 join IGE1 T1 on T0.""DocEntry""=T1.""DocEntry"" where ifnull(T1.""LineStatus"",'O')='O' and T0.""DocEntry"" in (" & DocEntry & ") order by T0.""DocNum"",T1.""LineNum""  "
                objRS1.DoQuery(StrQuery)
                objGRMatrix.Clear()
                objGRMatrix.AddRow()
                'If objaddon.HANA Then
                '    series = objaddon.objglobalmethods.getSingleValue("select ""Series"" From NNM1 where ""ObjectCode""='59' and ""Indicator""=(select Top 1 ""Indicator""  from OFPR where '" & Now.Date.ToString("yyyyMMdd") & "' between ""F_RefDate"" and ""T_RefDate"") " &
                '                                                                " and ""BPLId""=(Select ""BPLid"" from OWHS Where ""WhsCode""='" & objRS1.Fields.Item("Whse").Value.ToString & "')") ''" & objRS1.Fields.Item("BPLId").Value.ToString & "'
                'Else
                '    series = objaddon.objglobalmethods.getSingleValue("select Series From NNM1 where ObjectCode='59' and Indicator=(select Top 1 Indicator  from OFPR where '" & Now.Date.ToString("yyyyMMdd") & "' between F_RefDate and T_RefDate) " &
                '                                                                " and BPLId=(Select BPLid from OWHS Where WhsCode='" & objRS1.Fields.Item("Whse").Value.ToString & "')") ''" & objRS1.Fields.Item("BPLId").Value.ToString & "' 
                'End If
                'objCombo = objGRform.Items.Item("30").Specific
                'If series <> "" Then objCombo.Select(series, SAPbouiCOM.BoSearchKey.psk_ByValue)
                objGRform.Items.Item("trefno").Specific.String = DocEntry
                objaddon.objapplication.SetStatusBarMessage("Goods Issue DocumentNumber-> " & DocNum, SAPbouiCOM.BoMessageTime.bmt_Medium, False)
                'If objRS1.RecordCount > 0 Then
                '    m_oProgBar = objaddon.objapplication.StatusBar.CreateProgressBar("My Progress", objRS1.RecordCount, True)
                '    m_oProgBar.Text = "Goods Issue Loading to Goods Receipt Please wait... DocumentNumber-> " & DocNum
                '    m_oProgBar.Value = 0
                '    Dim oUDFForm, tempform As SAPbouiCOM.Form
                '    oUDFForm = objaddon.objapplication.Forms.Item(objGRform.UDFFormUID)
                '    tempform = objaddon.objapplication.Forms.GetForm("0", 0) 'objaddon.objapplication.Forms.Item("0")
                '    oUDFForm.Items.Item("U_TOWHS").Specific.String = objRS1.Fields.Item("Whse").Value.ToString
                '    oUDFForm.Items.Item("U_RefNo").Enabled = False
                '    objCombo = oUDFForm.Items.Item("U_TRANSTYPE").Specific
                '    objCombo.Select("STOCK TRANSFER", SAPbouiCOM.BoSearchKey.psk_ByDescription)
                '    If tempform.Visible = True Then
                '        tempform.Items.Item("1").Click()
                '    End If
                '    If objGRMatrix.Columns.Item("U_MRP").Editable = False Or objGRMatrix.Columns.Item("U_DocLine").Editable = False Then
                '        objGRMatrix.Columns.Item("U_MRP").Editable = True
                '        objGRMatrix.Columns.Item("U_DocLine").Editable = True
                '    End If
                '    For i As Integer = 0 To objRS1.RecordCount - 1
                '        Row += 1
                '        objGRMatrix.Columns.Item("1").Cells.Item(Row).Specific.String = Trim(objRS1.Fields.Item("ItemCode").Value.ToString)
                '        objGRMatrix.Columns.Item("11").Cells.Item(Row).Specific.String = Trim(CDbl(objRS1.Fields.Item("Quantity").Value.ToString))
                '        objGRMatrix.Columns.Item("15").Cells.Item(Row).Specific.String = Trim(objRS1.Fields.Item("Whse").Value.ToString)
                '        objGRMatrix.Columns.Item("21").Cells.Item(Row).Specific.String = Trim(objRS1.Fields.Item("Project").Value.ToString)
                '        objGRMatrix.Columns.Item("U_DocLine").Cells.Item(Row).Specific.String = Trim(objRS1.Fields.Item("LineNum").Value.ToString)
                '        objGRMatrix.Columns.Item("U_MRP").Cells.Item(Row).Specific.String = Trim(objRS1.Fields.Item("U_MRP").Value.ToString)
                '        objRS1.MoveNext()
                '        m_oProgBar.Value = i
                '    Next
                '    objGRMatrix.AutoResizeColumns()
                '    objGRMatrix.Columns.Item("11").Cells.Item(1).Click()
                '    'objGRMatrix.Columns.Item("U_MRP").Editable = False
                '    objGRMatrix.Columns.Item("U_DocLine").Editable = False
                '    objaddon.objapplication.StatusBar.SetText("Goods Issue Loaded to Goods Receipt Successfully!!! ", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                '    m_oProgBar.Text = "Goods Issue Loaded to Goods Receipt Successfully!!! "
                '    objRS1 = Nothing
                'End If


            Catch ex As Exception
                objaddon.objapplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Finally
                'm_oProgBar.Stop()
                'System.Runtime.InteropServices.Marshal.ReleaseComObject(m_oProgBar)
                'm_oProgBar = Nothing
                GC.Collect()
                GC.WaitForPendingFinalizers()
            End Try
        End Sub

        Private Sub Load_Delivery_ARInvoice(ByVal DocEntry As String, ByVal DocNum As String)
            'Dim FormID, series As String
            Dim objRS1 As SAPbobsCOM.Recordset
            Dim objARform As SAPbouiCOM.Form
            Dim objARMatrix As SAPbouiCOM.Matrix
            Dim objCombo As SAPbouiCOM.ComboBox
            Dim Row As Integer = 0
            Dim FormType As Integer
            'Dim m_oProgBar As SAPbouiCOM.ProgressBar
            Try
                If DocEntry = "" Then
                    Exit Sub
                End If

                'GRPODocEntry = DocEntry
                'objaddon.objapplication.SetStatusBarMessage("Delivery Loading to A/R Invoice Please Wait...", SAPbouiCOM.BoMessageTime.bmt_Short, False)
                FormType = objaddon.objapplication.Forms.ActiveForm.TypeCount
                objARform = objaddon.objapplication.Forms.GetForm("133", FormType)
                objARMatrix = objARform.Items.Item("38").Specific
                'odbdsDetails = objARform.DataSources.DBDataSources.Item("INV1")
                objRS1 = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                'objRS = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                'objaddon.objapplication.SetStatusBarMessage("Delivery Loading to A/R Invoice Please wait... DocumentNumber-> " & DocNum, SAPbouiCOM.BoMessageTime.bmt_Medium, False)

                StrQuery = "Select T0.""BPLId"",(Select ""BPLid"" from OWHS Where ""WhsCode""=T0.""U_TOWHS"") as ""To_Branch"",T0.""U_FROMWHS"",T0.""U_TRANSTYPE"",T0.""U_TOWHS"","
                StrQuery += vbCrLf + "T0.""Comments"",T0.""DocNum"",T1.""DocEntry"",T1.""LineNum"",T1.""ItemCode"",T1.""U_MRP"",T1.""Project"",T1.""Quantity"",T1.""TaxCode"",T1.""WhsCode"","
                StrQuery += vbCrLf + "T1.""AcctCode"",T1.""StockPrice"",T1.""LocCode"" "
                StrQuery += vbCrLf + "from ODLN T0 join DLN1 T1 on T0.""DocEntry""=T1.""DocEntry"" where T0.""DocStatus""='O' and ifnull(T1.""LineStatus"",'O')='O'  "
                StrQuery += vbCrLf + "and T0.""DocType""='I' and T0.""DocEntry"" in (" & DocEntry & ") order by T0.""DocNum"",T1.""LineNum""   "
                objRS1.DoQuery(StrQuery)
                objARMatrix.Clear()
                'odbdsDetails.Clear()
                objARMatrix.AddRow()

                objARform.Items.Item("trefno").Specific.String = DocEntry
                objaddon.objapplication.SetStatusBarMessage("Delivery Loading to A/R Invoice Please wait... DocumentNumber-> " & DocNum, SAPbouiCOM.BoMessageTime.bmt_Short, False)
                'objCombo = objARform.Items.Item("2001").Specific
                'objCombo.Select(Trim(objRS1.Fields.Item("BPLId").Value.ToString), SAPbouiCOM.BoSearchKey.psk_ByValue)
                'If objRS1.RecordCount > 0 Then
                '    m_oProgBar = objaddon.objapplication.StatusBar.CreateProgressBar("My Progress", objRS1.RecordCount, True)
                '    m_oProgBar.Text = "Delivery Loading to A/R Invoice Please wait... DocumentNumber-> " & DocNum
                '    m_oProgBar.Value = 0
                '    Dim oUDFForm, tempform As SAPbouiCOM.Form
                '    oUDFForm = objaddon.objapplication.Forms.Item(objARform.UDFFormUID)
                '    tempform = objaddon.objapplication.Forms.GetForm("0", 0) 'objaddon.objapplication.Forms.Item("0")
                '    oUDFForm.Items.Item("U_FROMWHS").Specific.String = objRS1.Fields.Item("U_FROMWHS").Value.ToString
                '    oUDFForm.Items.Item("U_TOWHS").Specific.String = objRS1.Fields.Item("U_TOWHS").Value.ToString

                '    oUDFForm.Items.Item("U_RefNo").Enabled = False
                '    objCombo = oUDFForm.Items.Item("U_TRANSTYPE").Specific
                '    objCombo.Select("STOCK TRANSFER", SAPbouiCOM.BoSearchKey.psk_ByDescription)
                '    If tempform.Visible = True Then
                '        tempform.Items.Item("1").Click()
                '    End If
                '    If objARMatrix.Columns.Item("U_MRP").Editable = False Or objARMatrix.Columns.Item("U_DocLine").Editable = False Then
                '        objARMatrix.Columns.Item("U_MRP").Editable = True
                '        objARMatrix.Columns.Item("U_DocLine").Editable = True
                '    End If
                '    For i As Integer = 0 To objRS1.RecordCount - 1
                '        Row += 1
                '        objARMatrix.Columns.Item("1").Cells.Item(Row).Specific.String = Trim(objRS1.Fields.Item("ItemCode").Value.ToString)
                '        objARMatrix.Columns.Item("11").Cells.Item(Row).Specific.String = Trim(CDbl(objRS1.Fields.Item("Quantity").Value.ToString))
                '        objARMatrix.Columns.Item("160").Cells.Item(Row).Specific.String = Trim(objRS1.Fields.Item("TaxCode").Value.ToString)
                '        objARMatrix.Columns.Item("24").Cells.Item(Row).Specific.String = Trim(objRS1.Fields.Item("WhsCode").Value.ToString) 'Trim(objRS1.Fields.Item("U_TOWHS").Value.ToString)
                '        objARMatrix.Columns.Item("31").Cells.Item(Row).Specific.String = Trim(objRS1.Fields.Item("Project").Value.ToString)
                '        objARMatrix.Columns.Item("U_DocLine").Cells.Item(Row).Specific.String = Trim(objRS1.Fields.Item("LineNum").Value.ToString)
                '        objARMatrix.Columns.Item("U_MRP").Cells.Item(Row).Specific.String = Trim(objRS1.Fields.Item("U_MRP").Value.ToString)
                '        objRS1.MoveNext()
                '        m_oProgBar.Value = i
                '    Next
                '    'While Not objRS1.EoF  ' Not working
                '    '    If Trim(objRS1.Fields.Item("ItemCode").Value) <> "" Then
                '    '        If objARMatrix.Columns.Item("1").Cells.Item(objARMatrix.VisualRowCount).Specific.String <> "" Then
                '    '            objARMatrix.AddRow()
                '    '        End If

                '    '        'odbdsDetails.Clear()
                '    '        objARMatrix.GetLineData(objARMatrix.VisualRowCount)
                '    '        odbdsDetails.SetValue("ItemCode", 0, Trim(objRS1.Fields.Item("ItemCode").Value.ToString))
                '    '        odbdsDetails.SetValue("Quantity", 0, Trim(CDbl(objRS1.Fields.Item("Quantity").Value.ToString)))
                '    '        odbdsDetails.SetValue("TaxCode", 0, Trim(objRS1.Fields.Item("TaxCode").Value.ToString))
                '    '        odbdsDetails.SetValue("WhsCode", 0, Trim(objRS1.Fields.Item("WhsCode").Value.ToString))
                '    '        odbdsDetails.SetValue("Project", 0, Trim(objRS1.Fields.Item("Project").Value.ToString))
                '    '        odbdsDetails.SetValue("U_MRP", 0, Trim(objRS1.Fields.Item("U_MRP").Value.ToString))
                '    '        objARMatrix.SetLineData(objARMatrix.VisualRowCount)
                '    '    End If
                '    '    objRS1.MoveNext()
                '    'End While
                '    objARMatrix.AutoResizeColumns()
                '    objARMatrix.Columns.Item("11").Cells.Item(1).Click()
                '    'objARMatrix.Columns.Item("U_MRP").Editable = False
                '    objARMatrix.Columns.Item("U_DocLine").Editable = False
                '    objaddon.objapplication.StatusBar.SetText("Delivery Loaded to A/R Invoice Successfully!!! ", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                '    m_oProgBar.Text = "Delivery Loaded to A/R Invoice Successfully!!!"
                '    objRS1 = Nothing
                'End If


            Catch ex As Exception
                objaddon.objapplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Finally
                'm_oProgBar.Stop()
                'System.Runtime.InteropServices.Marshal.ReleaseComObject(m_oProgBar)
                'm_oProgBar = Nothing
                GC.Collect()
                GC.WaitForPendingFinalizers()
            End Try
        End Sub

        Private Sub Load_ARInvoice_GRPO(ByVal DocEntry As String, ByVal DocNum As String, ByVal CardCode As String)
            'Dim FormID, series As String
            Dim objRS1 As SAPbobsCOM.Recordset
            Dim objGRform As SAPbouiCOM.Form
            Dim objGRMatrix As SAPbouiCOM.Matrix
            Dim objCombo As SAPbouiCOM.ComboBox
            Dim Row As Integer = 0
            Dim FormType As Integer
            'Dim m_oProgBar As SAPbouiCOM.ProgressBar
            Try
                If DocEntry = "" Then
                    Exit Sub
                End If
                'GRPODocEntry = DocEntry
                'objaddon.objapplication.SetStatusBarMessage("Delivery Loading to A/R Invoice Please Wait...", SAPbouiCOM.BoMessageTime.bmt_Short, False)
                FormType = objaddon.objapplication.Forms.ActiveForm.TypeCount
                objGRform = objaddon.objapplication.Forms.GetForm("143", FormType)
                objGRMatrix = objGRform.Items.Item("38").Specific
                objRS1 = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                'objRS = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                'objaddon.objapplication.SetStatusBarMessage("A/R Invoice Loading to GRPO Please wait... DocumentNumber-> " & DocNum, SAPbouiCOM.BoMessageTime.bmt_Medium, False)

                StrQuery = "Select T0.""CardCode"",T0.""BPLId"",(Select ""BPLid"" from OWHS Where ""WhsCode""=T0.""U_TOWHS"") as ""To_Branch"",T0.""U_TRANSTYPE"",T0.""U_TOWHS"","
                StrQuery += vbCrLf + "T0.""Comments"",T0.""DocNum"",T1.""DocEntry"",T1.""LineNum"",T1.""ItemCode"",T1.""Project"",T1.""Quantity"",T1.""U_MRP"",T1.""TaxCode"",T1.""WhsCode"","
                StrQuery += vbCrLf + "(Select ""State""||'IGST'||(Select case when ""U_SALESTAX"" is null then 0 else ""U_SALESTAX"" End From OITM where ""ItemCode""=T1.""ItemCode"") from OWHS where ""WhsCode""=T0.""U_TOWHS"")as ""ATaxCode"","
                StrQuery += vbCrLf + "T1.""AcctCode"",T1.""StockPrice"",T1.""LocCode"",T0.""Address"",T0.""Address2"""
                StrQuery += vbCrLf + "from OINV T0 join INV1 T1 on T0.""DocEntry""=T1.""DocEntry"" where "
                StrQuery += vbCrLf + "T0.""DocType""='I' and T0.""DocEntry"" in (" & DocEntry & ") and T0.""CardCode""=(Select distinct ""U_CUSTOMER"" from ""@MIPL_STBP"""
                StrQuery += vbCrLf + "where ""U_FRBRANCH""=(Select ""BPLid"" from OWHS Where ""WhsCode""=T0.""U_TOWHS"") and ""U_VENDOR""='" & CardCode & "')"
                StrQuery += vbCrLf + "order by T0.""DocNum"",T1.""LineNum""  "

                objRS1.DoQuery(StrQuery)
                objGRMatrix.Clear()
                objGRMatrix.AddRow()

                objGRform.Items.Item("trefno").Specific.String = DocEntry
                objaddon.objapplication.SetStatusBarMessage("A/R Invoice DocumentNumber-> " & DocNum, SAPbouiCOM.BoMessageTime.bmt_Medium, False)

                'objCombo = objGRform.Items.Item("2001").Specific
                'objCombo.Select(Trim(objRS1.Fields.Item("To_Branch").Value.ToString), SAPbouiCOM.BoSearchKey.psk_ByValue)
                'If objRS1.RecordCount > 0 Then
                '    m_oProgBar = objaddon.objapplication.StatusBar.CreateProgressBar("My Progress", objRS1.RecordCount, True)
                '    m_oProgBar.Text = "A/R Invoice Loading to GRPO Please wait... DocumentNumber-> " & DocNum
                '    m_oProgBar.Value = 0
                '    Dim oUDFForm, tempform As SAPbouiCOM.Form
                '    oUDFForm = objaddon.objapplication.Forms.Item(objGRform.UDFFormUID)
                '    tempform = objaddon.objapplication.Forms.GetForm("0", 0) 'objaddon.objapplication.Forms.Item("0")
                '    oUDFForm.Items.Item("U_TOWHS").Specific.String = objRS1.Fields.Item("U_TOWHS").Value.ToString
                '    oUDFForm.Items.Item("U_RefNo").Enabled = False
                '    objCombo = oUDFForm.Items.Item("U_TRANSTYPE").Specific
                '    objCombo.Select("STOCK TRANSFER", SAPbouiCOM.BoSearchKey.psk_ByDescription)
                '    If tempform.Visible = True Then
                '        tempform.Items.Item("1").Click()
                '    End If
                '    If objGRMatrix.Columns.Item("U_MRP").Editable = False Or objGRMatrix.Columns.Item("U_DocLine").Editable = False Then
                '        objGRMatrix.Columns.Item("U_MRP").Editable = True
                '        objGRMatrix.Columns.Item("U_DocLine").Editable = True
                '    End If
                '    For i As Integer = 0 To objRS1.RecordCount - 1
                '        Row += 1
                '        objGRMatrix.Columns.Item("1").Cells.Item(Row).Specific.String = Trim(objRS1.Fields.Item("ItemCode").Value.ToString)
                '        objGRMatrix.Columns.Item("11").Cells.Item(Row).Specific.String = Trim(CDbl(objRS1.Fields.Item("Quantity").Value.ToString))
                '        objGRMatrix.Columns.Item("160").Cells.Item(Row).Specific.String = Trim(objRS1.Fields.Item("ATaxCode").Value.ToString) ' Trim(objRS1.Fields.Item("TaxCode").Value.ToString)
                '        objGRMatrix.Columns.Item("24").Cells.Item(Row).Specific.String = Trim(objRS1.Fields.Item("U_TOWHS").Value.ToString)
                '        objGRMatrix.Columns.Item("31").Cells.Item(Row).Specific.String = Trim(objRS1.Fields.Item("Project").Value.ToString)
                '        objGRMatrix.Columns.Item("U_DocLine").Cells.Item(Row).Specific.String = Trim(objRS1.Fields.Item("LineNum").Value.ToString)
                '        objGRMatrix.Columns.Item("U_MRP").Cells.Item(Row).Specific.String = Trim(objRS1.Fields.Item("U_MRP").Value.ToString)
                '        objRS1.MoveNext()
                '        m_oProgBar.Value = i
                '    Next
                '    objGRMatrix.AutoResizeColumns()
                '    objGRMatrix.Columns.Item("11").Cells.Item(1).Click()
                '    'objGRMatrix.Columns.Item("U_MRP").Editable = False
                '    objGRMatrix.Columns.Item("U_DocLine").Editable = False
                '    objaddon.objapplication.StatusBar.SetText("A/R Invoice Loaded to GRPO Successfully!!! ", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                '    m_oProgBar.Text = "A/R Invoice Loaded to GRPO Successfully!!!"
                '    objRS1 = Nothing
                'End If
            Catch ex As Exception
                objaddon.objapplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Finally
                'm_oProgBar.Stop()
                'System.Runtime.InteropServices.Marshal.ReleaseComObject(m_oProgBar)
                'm_oProgBar = Nothing
                GC.Collect()
                GC.WaitForPendingFinalizers()
            End Try
        End Sub

    End Class
End Namespace
