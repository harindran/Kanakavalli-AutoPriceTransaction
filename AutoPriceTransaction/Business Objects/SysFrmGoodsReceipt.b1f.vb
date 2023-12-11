Option Strict Off
Option Explicit On

Imports SAPbouiCOM.Framework

Namespace AutoPriceTransaction
    <FormAttribute("721", "Business Objects/SysFrmGoodsReceipt.b1f")>
    Friend Class SysFrmGoodsReceipt
        Inherits SystemFormBase
        Public WithEvents objform, objformUDF As SAPbouiCOM.Form
        Private Shared FormCount As Integer = 0
        'Dim FormCount As Integer = 0
        Dim StrQuery As String
        Dim objRs, objRs1 As SAPbobsCOM.Recordset
        Private WithEvents objCombo As SAPbouiCOM.ComboBox
        Public Sub New()
        End Sub

        Public Overrides Sub OnInitializeComponent()
            Me.Button0 = CType(Me.GetItem("btnsave").Specific, SAPbouiCOM.Button)
            Me.Matrix0 = CType(Me.GetItem("13").Specific, SAPbouiCOM.Matrix)
            Me.Button1 = CType(Me.GetItem("1").Specific, SAPbouiCOM.Button)
            Me.ComboBox0 = CType(Me.GetItem("2310000079").Specific, SAPbouiCOM.ComboBox)
            Me.StaticText0 = CType(Me.GetItem("lgirefno").Specific, SAPbouiCOM.StaticText)
            Me.EditText0 = CType(Me.GetItem("trefno").Specific, SAPbouiCOM.EditText)
            Me.OnCustomInitialize()

        End Sub

        Public Overrides Sub OnInitializeFormEvents()
            AddHandler CloseAfter, AddressOf Me.Form_CloseAfter
            AddHandler DataLoadAfter, AddressOf Me.Form_DataLoadAfter
            AddHandler ClickBefore, AddressOf Me.Form_ClickBefore

        End Sub

        Private WithEvents Button0 As SAPbouiCOM.Button
        Private WithEvents Matrix0 As SAPbouiCOM.Matrix
        Private WithEvents Button1 As SAPbouiCOM.Button

        Private Sub OnCustomInitialize()
            Try
                FormCount += 1
                'Dim FormType As Integer
                'FormType = objaddon.objapplication.Forms.ActiveForm.TypeCount
                objform = objaddon.objapplication.Forms.GetForm("721", FormCount)
                'objform = objaddon.objapplication.Forms.ActiveForm
                objform.Items.Item("1").Enabled = False
                objform.Items.Item("btnsave").Left = objform.Items.Item("2").Left + objform.Items.Item("2").Width + 5
                objform.Items.Item("btnsave").Top = objform.Items.Item("2").Top
                objform.Items.Item("btnsave").Height = objform.Items.Item("2").Height
                objform.Items.Item("btnsave").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Ok, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                objform.Items.Item("btnsave").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Find, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                'Try
                '    If Not objaddon.objapplication.Menus.Item("6913").Checked = True Then
                '        objaddon.objapplication.Menus.Item("6913").Activate()
                '        objformUDF = objaddon.objapplication.Forms.GetForm("-721", 1)
                '        objformUDF.Items.Item("U_GIRefNo").Enabled = False
                '    Else
                '        objformUDF = objaddon.objapplication.Forms.GetForm("-721", 1) '(objform.UDFFormUID, 0)
                '        objformUDF.Items.Item("U_GIRefNo").Enabled = False
                '    End If
                'Catch ex As Exception
                'End Try

                ''Dim oDrafts, objGoodsR As SAPbobsCOM.Documents
                ''oDrafts = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDrafts)
                ''oDrafts.SaveDraftToDocument()
                'If oDrafts.SaveDraftToDocument() = 0 Then
                '    If objaddon.objcompany.InTransaction Then objaddon.objcompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
                'Else
                '    If objaddon.objcompany.InTransaction Then objaddon.objcompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                '    objaddon.objapplication.StatusBar.SetText("Goods Issue : " & objaddon.objcompany.GetLastErrorDescription, SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                'End If

            Catch ex As Exception
                'MsgBox("Error : " + ex.Message + vbCrLf + "Position : " + ex.StackTrace, MsgBoxStyle.Critical)
            End Try
        End Sub

        Private Sub Button0_ClickBefore(sboObject As Object, pVal As SAPbouiCOM.SBOItemEventArg, ByRef BubbleEvent As Boolean) Handles Button0.ClickBefore
            Try
                'Dim ErrorFlag As Boolean
                'Dim ItemCode As String = ""
                'Dim Qty As Double

                If objform.Mode <> SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                    BubbleEvent = False : Exit Sub
                End If
                If objform.Items.Item("trefno").Specific.String = "" Then 'TransactionEntry = ""
                    objaddon.objapplication.StatusBar.SetText("Please select a Goods Issue Entries...", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    Exit Sub
                End If
                If objform.Items.Item("9").Specific.String = "" Then
                    objaddon.objapplication.StatusBar.SetText("Posting Date is Missing", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    BubbleEvent = False : Exit Sub
                End If
                If objform.Items.Item("38").Specific.String = "" Then
                    objaddon.objapplication.StatusBar.SetText("Document Date is Missing", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    BubbleEvent = False : Exit Sub
                End If
                TransactionEntry = objform.Items.Item("trefno").Specific.String
                'objaddon.objapplication.StatusBar.SetText("Validating Please wait...", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                'objRs = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                'For j As Integer = 1 To Matrix0.VisualRowCount
                '    ItemCode = Matrix0.Columns.Item("1").Cells.Item(j).Specific.string
                '    Qty = CDbl(Matrix0.Columns.Item("9").Cells.Item(j).Specific.string)
                '    If ItemCode = "" Then Continue For
                '    StrQuery = "Select T0.""BPLId"",T0.""U_TRANSTYPE"",T0.""U_TOWHS"",T0.""Comments"",T1.""DocEntry"",T1.""LineNum"",T1.""ItemCode"",T1.""Quantity"",T1.""WhsCode"",T1.""AcctCode"",T1.""StockPrice"",T1.""LocCode"""
                '    StrQuery += vbCrLf + " from OIGE T0 join IGE1 T1 on T0.""DocEntry""=T1.""DocEntry"" where ifnull(T1.""LineStatus"",'O')='O' and T0.""DocEntry"" in (" & TransactionEntry & ") and T1.""ItemCode""='" & ItemCode & "' order by T0.""DocNum"",T1.""LineNum""  "
                '    objRs.DoQuery(StrQuery)
                '    If objRs.RecordCount > 0 Then
                '        If ItemCode <> "" And Qty > CDbl(objRs.Fields.Item("Quantity").Value) Then
                '            ErrorFlag = True
                '            Exit Try
                '        End If
                '    Else
                '        objaddon.objapplication.StatusBar.SetText("No Data Found...", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                '        BubbleEvent = False
                '    End If
                'Next

                'If ErrorFlag Then
                '    objaddon.objapplication.StatusBar.SetText("Validate: Quantity cannot exceed the quantity in the base document, '" & ItemCode & "'", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                '    BubbleEvent = False
                'End If
                'objaddon.objapplication.StatusBar.SetText("Validating Completed...", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

            Catch ex As Exception

            End Try

        End Sub

        Private Function Create_GoodsReceipt() As Boolean
            Try
                Dim BranchEnabled, series, StrQuery As String
                Dim Retval As Integer
                Dim objGoodsReceipt As SAPbobsCOM.Documents
                Dim objRs As SAPbobsCOM.Recordset
                objRs = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                objGoodsReceipt = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInventoryGenEntry)
                If Not objaddon.objcompany.InTransaction Then objaddon.objcompany.StartTransaction()
                objaddon.objapplication.StatusBar.SetText("Creating Goods Receipt Please wait...", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                objGoodsReceipt.DocDate = Now.Date
                objGoodsReceipt.TaxDate = Now.Date
                If objaddon.HANA Then
                    BranchEnabled = objaddon.objglobalmethods.getSingleValue("select ""MltpBrnchs"" from OADM")
                Else
                    BranchEnabled = objaddon.objglobalmethods.getSingleValue("select MltpBrnchs from OADM")
                End If
                'objGoodsReceipt.Reference1 = objform.Items.Item("txtvrefno").Specific.string
                StrQuery = "Select T0.""BPLId"",T0.""U_TRANSTYPE"",T0.""U_TOWHS"",T0.""Comments"",T1.""DocEntry"",T1.""LineNum"",T1.""ItemCode"",T1.""Quantity"",T1.""WhsCode"",T1.""AcctCode"",T1.""StockPrice"",T1.""LocCode"""
                StrQuery += vbCrLf + " from OIGE T0 join IGE1 T1 on T0.""DocEntry""=T1.""DocEntry"" where ifnull(T1.""LineStatus"",'O')='O' and T0.""DocEntry"" in (" & TransactionEntry & ") order by T0.""DocNum"",T1.""LineNum""  "
                objRs.DoQuery(StrQuery)
                If objRs.RecordCount > 0 Then
                    If BranchEnabled = "Y" Then
                        If objaddon.HANA Then
                            series = objaddon.objglobalmethods.getSingleValue("select ""Series"" From NNM1 where ""ObjectCode""='59' and ""Indicator""=(select Top 1 ""Indicator""  from OFPR where '" & Now.Date.ToString("yyyyMMdd") & "' between ""F_RefDate"" and ""T_RefDate"") " &
                                                                                " and ""BPLId""='" & objRs.Fields.Item("BPLId").Value.ToString & "' ")
                        Else
                            series = objaddon.objglobalmethods.getSingleValue("select Series From NNM1 where ObjectCode='59' and Indicator=(select Top 1 Indicator  from OFPR where '" & Now.Date.ToString("yyyyMMdd") & "' between F_RefDate and T_RefDate) " &
                                                                                " and BPLId='" & objRs.Fields.Item("BPLId").Value.ToString & "' ")
                        End If
                        If series <> "" Then
                            objGoodsReceipt.Series = series
                        End If
                        objGoodsReceipt.BPL_IDAssignedToInvoice = objRs.Fields.Item("BPLId").Value.ToString
                    End If
                    objGoodsReceipt.Comments = objRs.Fields.Item("Comments").Value.ToString & " Auto Posted ->" & Now.ToString
                    objGoodsReceipt.UserFields.Fields.Item("U_TRANSTYPE").Value = objRs.Fields.Item("U_TRANSTYPE").Value.ToString
                    objGoodsReceipt.UserFields.Fields.Item("U_TOWHS").Value = objRs.Fields.Item("U_TOWHS").Value.ToString

                    For i As Integer = 0 To objRs.RecordCount - 1
                        objGoodsReceipt.Lines.ItemCode = objRs.Fields.Item("ItemCode").Value.ToString
                        objGoodsReceipt.Lines.Quantity = CDbl(objRs.Fields.Item("Quantity").Value.ToString)
                        objGoodsReceipt.Lines.UnitPrice = CDbl(objRs.Fields.Item("StockPrice").Value.ToString)
                        objGoodsReceipt.Lines.WarehouseCode = objRs.Fields.Item("U_TOWHS").Value.ToString 'objRs.Fields.Item("WhsCode").Value.ToString
                        objGoodsReceipt.Lines.BaseType = 60
                        objGoodsReceipt.Lines.BaseEntry = CInt(objRs.Fields.Item("DocEntry").Value.ToString)
                        objGoodsReceipt.Lines.BaseLine = CInt(objRs.Fields.Item("LineNum").Value.ToString)
                        objGoodsReceipt.Lines.AccountCode = objRs.Fields.Item("AcctCode").Value.ToString
                        objGoodsReceipt.Lines.LocationCode = objRs.Fields.Item("LocCode").Value.ToString
                        'objGoodsReceipt.Lines.SetCurrentLine(i)
                        objGoodsReceipt.Lines.Add()
                        objRs.MoveNext()
                    Next
                End If
                Retval = objGoodsReceipt.Add()
                If Retval <> 0 Then
                    If objaddon.objcompany.InTransaction Then objaddon.objcompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                    objaddon.objapplication.StatusBar.SetText("Goods Receipt : " & objaddon.objcompany.GetLastErrorDescription, SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(objGoodsReceipt)
                    Return False
                Else
                    If objaddon.objcompany.InTransaction Then objaddon.objcompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
                    'Docentry = objaddon.objcompany.GetNewObjectKey()
                    TransactionEntry = ""
                    objaddon.objapplication.StatusBar.SetText("Goods Receipt Created Successfully...", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(objGoodsReceipt)
                    GC.Collect()
                    Return True
                End If

            Catch ex As Exception
                objaddon.objapplication.StatusBar.SetText("Goods Receipt " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            End Try
        End Function

        Private Function Create_GoodsReceipt_Updated(ByVal DocEntry As String) As Boolean
            Dim m_oProgBar As SAPbouiCOM.ProgressBar
            Try
                Dim BranchEnabled, series, Batchs, Serial, TranEntry As String
                Dim Retval As Integer
                Dim ItemCode As String = ""
                Dim Qty As Double
                Dim objGoodsReceipt As SAPbobsCOM.Documents
                Dim objRecset As SAPbobsCOM.Recordset
                Dim objEdit As SAPbouiCOM.EditText
                objRs = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                objRs1 = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                objRecset = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                objGoodsReceipt = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInventoryGenEntry)
                objaddon.objapplication.StatusBar.SetText("Creating Goods Receipt Please wait...", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                objEdit = objform.Items.Item("9").Specific
                Dim DocDate As Date = Date.ParseExact(objEdit.Value, "yyyyMMdd", System.Globalization.DateTimeFormatInfo.InvariantInfo)
                objEdit = objform.Items.Item("38").Specific
                Dim TaxDate As Date = Date.ParseExact(objEdit.Value, "yyyyMMdd", System.Globalization.DateTimeFormatInfo.InvariantInfo)
                objGoodsReceipt.DocDate = DocDate ' Now.Date
                objGoodsReceipt.TaxDate = TaxDate ' Now.Date
                If objaddon.HANA Then
                    BranchEnabled = objaddon.objglobalmethods.getSingleValue("select ""MltpBrnchs"" from OADM")
                Else
                    BranchEnabled = objaddon.objglobalmethods.getSingleValue("select MltpBrnchs from OADM")
                End If
                'objGoodsReceipt.Reference1 = objform.Items.Item("txtvrefno").Specific.string
                StrQuery = "Select T0.""BPLId"",(Select ""BPLid"" from OWHS Where ""WhsCode""=T0.""U_TOWHS"") as ""To_Branch"",T0.""U_TRANSTYPE"",T0.""U_TOWHS"",T0.""U_FROMWHS"",T0.""Comments"",T0.""DocNum"",T1.""DocEntry"",T1.""LineNum"",T1.""ItemCode"",T1.""Project"",T1.""Quantity"",T1.""WhsCode"",T1.""AcctCode"",T1.""StockPrice"",T1.""LocCode"""
                StrQuery += vbCrLf + ",(SELECT distinct ""Price"" FROM ITM1 Where ""PriceList""=2 and ""ItemCode""=T1.""ItemCode"") as ""MRP"""
                StrQuery += vbCrLf + " from OIGE T0 join IGE1 T1 on T0.""DocEntry""=T1.""DocEntry"" where ifnull(T1.""LineStatus"",'O')='O' and T0.""DocEntry"" in (" & DocEntry & ") order by T0.""DocNum"",T1.""LineNum""  "
                objRs.DoQuery(StrQuery)
                'm_oProgBar = objaddon.objapplication.StatusBar.CreateProgressBar("My Progress", Matrix0.VisualRowCount, True)

                If objRs.RecordCount > 0 Then
                    m_oProgBar = objaddon.objapplication.StatusBar.CreateProgressBar("My Progress", objRs.RecordCount, True)
                    If Not objaddon.objcompany.InTransaction Then objaddon.objcompany.StartTransaction()
                    m_oProgBar.Text = "Creating Goods Receipt Please wait..."
                    m_oProgBar.Value = 0
                    If BranchEnabled = "Y" Then
                        'If objaddon.HANA Then
                        '    series = objaddon.objglobalmethods.getSingleValue("select ""Series"" From NNM1 where ""ObjectCode""='59' and ""Indicator""=(select Top 1 ""Indicator""  from OFPR where '" & Now.Date.ToString("yyyyMMdd") & "' between ""F_RefDate"" and ""T_RefDate"") " &
                        '                                                        " and ""BPLId""='" & objRs.Fields.Item("BPLId").Value.ToString & "' ")
                        'Else
                        '    series = objaddon.objglobalmethods.getSingleValue("select Series From NNM1 where ObjectCode='59' and Indicator=(select Top 1 Indicator  from OFPR where '" & Now.Date.ToString("yyyyMMdd") & "' between F_RefDate and T_RefDate) " &
                        '                                                        " and BPLId='" & objRs.Fields.Item("BPLId").Value.ToString & "' ")
                        'End If

                        If objaddon.HANA Then
                            series = objaddon.objglobalmethods.getSingleValue("select ""Series"" From NNM1 where ""ObjectCode""='59' and ""Indicator""=(select Top 1 ""Indicator""  from OFPR where '" & Now.Date.ToString("yyyyMMdd") & "' between ""F_RefDate"" and ""T_RefDate"") " &
                                                                                " and ""BPLId""=(Select ""BPLid"" from OWHS Where ""WhsCode""='" & objRs.Fields.Item("U_TOWHS").Value.ToString & "')")
                        Else
                            series = objaddon.objglobalmethods.getSingleValue("select Series From NNM1 where ObjectCode='59' and Indicator=(select Top 1 Indicator  from OFPR where '" & Now.Date.ToString("yyyyMMdd") & "' between F_RefDate and T_RefDate) " &
                                                                                " and BPLId=(Select BPLid from OWHS Where WhsCode='" & objRs.Fields.Item("U_TOWHS").Value.ToString & "')")
                        End If
                        If series <> "" Then
                            objGoodsReceipt.Series = series
                        End If
                        objGoodsReceipt.BPL_IDAssignedToInvoice = objRs.Fields.Item("To_Branch").Value.ToString
                    End If
                    objGoodsReceipt.Comments = "Based On Goods Issue " & objRs.Fields.Item("DocNum").Value.ToString & "-" & objRs.Fields.Item("Comments").Value.ToString & " Auto Posted ->" & Now.ToString("dd/MMM/yyyy HH:mm:ss")
                    objGoodsReceipt.UserFields.Fields.Item("U_TRANSTYPE").Value = objRs.Fields.Item("U_TRANSTYPE").Value.ToString
                    objGoodsReceipt.UserFields.Fields.Item("U_TOWHS").Value = objRs.Fields.Item("U_TOWHS").Value.ToString
                    objGoodsReceipt.UserFields.Fields.Item("U_FROMWHS").Value = objRs.Fields.Item("U_FROMWHS").Value.ToString
                    objGoodsReceipt.UserFields.Fields.Item("U_RefNo").Value = DocEntry
                    Dim iRow As Integer = 0
                    Dim Loc As String
                    If objaddon.HANA Then
                        Loc = objaddon.objglobalmethods.getSingleValue("select T0.""Code"" from OLCT T0 join OWHS T1 on T0.""Code""=T1.""Location"" where T1.""WhsCode""='" & Trim(objRs.Fields.Item("U_TOWHS").Value.ToString) & "'")
                    Else
                        Loc = objaddon.objglobalmethods.getSingleValue("select T0.Code from OLCT T0 join OWHS T1 on T0.Code=T1.Location where T1.WhsCode='" & Trim(objRs.Fields.Item("U_TOWHS").Value.ToString) & "'")
                    End If
                    While Not objRs.EoF
                        ItemCode = Trim(objRs.Fields.Item("ItemCode").Value.ToString)
                        Qty = CDbl(objRs.Fields.Item("Quantity").Value.ToString)
                        objGoodsReceipt.Lines.ItemCode = ItemCode
                        objGoodsReceipt.Lines.Quantity = Qty 'CDbl(objRs.Fields.Item("Quantity").Value.ToString)
                        objGoodsReceipt.Lines.UnitPrice = CDbl(objRs.Fields.Item("StockPrice").Value.ToString)
                        objGoodsReceipt.Lines.WarehouseCode = objRs.Fields.Item("U_TOWHS").Value.ToString 'objRs.Fields.Item("WhsCode").Value.ToString
                        objGoodsReceipt.Lines.BaseType = 60
                        objGoodsReceipt.Lines.BaseEntry = CInt(objRs.Fields.Item("DocEntry").Value.ToString)
                        objGoodsReceipt.Lines.BaseLine = CInt(objRs.Fields.Item("LineNum").Value.ToString)
                        objGoodsReceipt.Lines.AccountCode = objRs.Fields.Item("AcctCode").Value.ToString
                        'select T0."Code" from OLCT T0 join OWHS T1 on T0."Code"=T1."Location" where T1."WhsCode"='E20/0001'

                        If Loc <> "" Then objGoodsReceipt.Lines.LocationCode = Loc 'objRs.Fields.Item("LocCode").Value.ToString
                        objGoodsReceipt.Lines.ProjectCode = objRs.Fields.Item("Project").Value.ToString
                        objGoodsReceipt.Lines.UserFields.Fields.Item("U_MRP").Value = objRs.Fields.Item("MRP").Value.ToString
                        objGoodsReceipt.Lines.UserFields.Fields.Item("U_DocLine").Value = objRs.Fields.Item("LineNum").Value.ToString
                        If objaddon.HANA Then
                            Serial = objaddon.objglobalmethods.getSingleValue("select ""ManSerNum"" from OITM WHERE ""ItemCode""='" & ItemCode & "'")
                            Batchs = objaddon.objglobalmethods.getSingleValue("select ""ManBtchNum"" from OITM WHERE ""ItemCode""='" & ItemCode & "'")
                        Else
                            Serial = objaddon.objglobalmethods.getSingleValue("select ManSerNum from OITM WHERE ItemCode='" & ItemCode & "'")
                            Batchs = objaddon.objglobalmethods.getSingleValue("select ManBtchNum from OITM WHERE ItemCode='" & ItemCode & "'")
                        End If
                        If Batchs = "Y" And Serial = "N" Then
                            StrQuery = "Select A.""BatchNum"" As ""BatchSerial"",  SUM(A.""Quantity"") As ""Qty"" FROM ("
                            StrQuery += vbCrLf + "Select T.""BatchNum"" , T.""Quantity"" from ibt1 T left join oibt T1 on T.""ItemCode""=T1.""ItemCode"" and T.""BatchNum""=T1.""BatchNum"" and T.""WhsCode""=T1.""WhsCode"""
                            StrQuery += vbCrLf + "left outer join IGE1 T2 on T2.""DocEntry""=T.""BaseEntry"" and T2.""ItemCode""=T.""ItemCode"" and T2.""LineNum""=T.""BaseLinNum"""
                            StrQuery += vbCrLf + "left outer join OIGE T3 on T2.""DocEntry""=T3.""DocEntry"" where T.""BaseType""='60' and T.""Direction""=1 and "
                            StrQuery += vbCrLf + "T.""ItemCode""='" & ItemCode & "' and T.""BaseEntry""in (" & DocEntry & ") )A GROUP BY A.""BatchNum"" having SUM(A.""Quantity"") >0"
                            objRecset.DoQuery(StrQuery)
                            Dim BQty As Double = Qty
                            If objRecset.RecordCount > 0 Then
                                For j As Integer = 0 To objRecset.RecordCount - 1
                                    objGoodsReceipt.Lines.BatchNumbers.BatchNumber = CStr(objRecset.Fields.Item("BatchSerial").Value)
                                    objGoodsReceipt.Lines.BatchNumbers.Quantity = BQty
                                    objGoodsReceipt.Lines.BatchNumbers.Add()
                                Next
                            End If
                        ElseIf Batchs = "N" And Serial = "Y" Then
                            StrQuery = "Select * from (SELECT distinct T4.""IntrSerial"" ""BatchSerial"",T4.""WhsCode"",T1.""DocEntry"",T1.""ItemCode"", T4.""Quantity"",T4.""Status"" from OIGE T0 left join IGE1 T1 on T0.""DocEntry""=T1.""DocEntry"""
                            StrQuery += vbCrLf + "left outer join SRI1 I1 on T1.""ItemCode""=I1.""ItemCode""   and (T1.""DocEntry""=I1.""BaseEntry"" and T1.""ObjType""=I1.""BaseType"") and T1.""LineNum""=I1.""BaseLinNum"""
                            StrQuery += vbCrLf + "left outer join OSRI T4 on T4.""ItemCode""=I1.""ItemCode"" and I1.""SysSerial""=T4.""SysSerial"" and I1.""WhsCode"" = T4.""WhsCode"" ) A "
                            StrQuery += vbCrLf + "Where A.""DocEntry"" in (" & DocEntry & ")  and A.""ItemCode""='" & ItemCode & "' and A.""BatchSerial"" <>'' and A.""Status""=1"
                            objRecset.DoQuery(StrQuery)
                            Dim SQty As Double = 0, TotSerialQty As Double = 0
                            SQty = Qty
                            If objRecset.RecordCount > 0 Then
                                For j As Integer = 0 To objRecset.RecordCount - 1
                                    objGoodsReceipt.Lines.SerialNumbers.InternalSerialNumber = CStr(objRecset.Fields.Item("BatchSerial").Value)
                                    objGoodsReceipt.Lines.SerialNumbers.Quantity = CDbl(1)
                                    objGoodsReceipt.Lines.SerialNumbers.Add()
                                    TotSerialQty += CDbl(1)
                                    If SQty - TotSerialQty > 0 Then
                                        objRecset.MoveNext()
                                    Else
                                        Exit For
                                    End If
                                Next
                            End If
                        End If
                        objGoodsReceipt.Lines.Add()
                        iRow += 1
                        m_oProgBar.Value = iRow
                        objRs.MoveNext()
                    End While

                    'For i As Integer = 1 To Matrix0.VisualRowCount
                    '    ItemCode = Trim(Matrix0.Columns.Item("1").Cells.Item(i).Specific.string)
                    '    Qty = CDbl(Matrix0.Columns.Item("9").Cells.Item(i).Specific.string)
                    '    If ItemCode = "" Then Continue For
                    '    StrQuery = "Select T0.""BPLId"",(Select ""BPLid"" from OWHS Where ""WhsCode""=T0.""U_TOWHS"") as ""To_Branch"",T0.""U_TRANSTYPE"",T0.""U_TOWHS"",T0.""U_FROMWHS"",T0.""Comments"",T0.""DocNum"",T1.""DocEntry"",T1.""LineNum"",T1.""ItemCode"",T1.""Project"",T1.""Quantity"",T1.""WhsCode"",T1.""AcctCode"",T1.""StockPrice"",T1.""LocCode"""
                    '    StrQuery += vbCrLf + ",(SELECT distinct ""Price"" FROM ITM1 Where ""PriceList""=2 and ""ItemCode""=T1.""ItemCode"") as ""MRP"""
                    '    StrQuery += vbCrLf + " from OIGE T0 join IGE1 T1 on T0.""DocEntry""=T1.""DocEntry"" where ifnull(T1.""LineStatus"",'O')='O' and T0.""DocEntry"" in (" & DocEntry & ") and T1.""ItemCode""='" & Trim(ItemCode) & "' and T1.""LineNum""='" & IIf(Trim(Matrix0.Columns.Item("U_DocLine").Cells.Item(i).Specific.string) = "", "-1", Trim(Matrix0.Columns.Item("U_DocLine").Cells.Item(i).Specific.string)) & "' order by T0.""DocNum"",T1.""LineNum""  "
                    '    objRs1.DoQuery(StrQuery)
                    '    If objRs1.RecordCount > 0 Then
                    '        If Qty <= CDbl(objRs1.Fields.Item("Quantity").Value) Then
                    '            objGoodsReceipt.Lines.ItemCode = ItemCode ' Matrix0.Columns.Item("1").Cells.Item(i).Specific.string
                    '            objGoodsReceipt.Lines.Quantity = CDbl(objRs1.Fields.Item("Quantity").Value.ToString) ' Qty 'CDbl(Matrix0.Columns.Item("9").Cells.Item(i).Specific.string)
                    '            objGoodsReceipt.Lines.UnitPrice = CDbl(objRs1.Fields.Item("StockPrice").Value.ToString)
                    '            objGoodsReceipt.Lines.WarehouseCode = objRs1.Fields.Item("U_TOWHS").Value.ToString 'objRs.Fields.Item("WhsCode").Value.ToString
                    '            objGoodsReceipt.Lines.BaseType = 60
                    '            objGoodsReceipt.Lines.BaseEntry = CInt(objRs1.Fields.Item("DocEntry").Value.ToString)
                    '            objGoodsReceipt.Lines.BaseLine = CInt(objRs1.Fields.Item("LineNum").Value.ToString)
                    '            objGoodsReceipt.Lines.AccountCode = objRs1.Fields.Item("AcctCode").Value.ToString
                    '            'objGoodsReceipt.Lines.LocationCode = objRs.Fields.Item("LocCode").Value.ToString
                    '            objGoodsReceipt.Lines.ProjectCode = objRs1.Fields.Item("Project").Value.ToString
                    '            objGoodsReceipt.Lines.UserFields.Fields.Item("U_MRP").Value = objRs1.Fields.Item("MRP").Value.ToString
                    '            objGoodsReceipt.Lines.UserFields.Fields.Item("U_DocLine").Value = objRs1.Fields.Item("LineNum").Value.ToString
                    '            'objGoodsReceipt.Lines.SetCurrentLine(i)

                    '            If objaddon.HANA Then
                    '                Serial = objaddon.objglobalmethods.getSingleValue("select ""ManSerNum"" from OITM WHERE ""ItemCode""='" & ItemCode & "'")
                    '                Batchs = objaddon.objglobalmethods.getSingleValue("select ""ManBtchNum"" from OITM WHERE ""ItemCode""='" & ItemCode & "'")
                    '            Else
                    '                Serial = objaddon.objglobalmethods.getSingleValue("select ManSerNum from OITM WHERE ItemCode='" & ItemCode & "'")
                    '                Batchs = objaddon.objglobalmethods.getSingleValue("select ManBtchNum from OITM WHERE ItemCode='" & ItemCode & "'")
                    '            End If
                    '            If Batchs = "Y" And Serial = "N" Then
                    '                StrQuery = "Select A.""BatchNum"" As ""BatchSerial"",  SUM(A.""Quantity"") As ""Qty"" FROM ("
                    '                StrQuery += vbCrLf + "Select T.""BatchNum"" , T.""Quantity"" from ibt1 T left join oibt T1 on T.""ItemCode""=T1.""ItemCode"" and T.""BatchNum""=T1.""BatchNum"" and T.""WhsCode""=T1.""WhsCode"""
                    '                StrQuery += vbCrLf + "left outer join IGE1 T2 on T2.""DocEntry""=T.""BaseEntry"" and T2.""ItemCode""=T.""ItemCode"" and T2.""LineNum""=T.""BaseLinNum"""
                    '                StrQuery += vbCrLf + "left outer join OIGE T3 on T2.""DocEntry""=T3.""DocEntry"" where T.""BaseType""='60' and T.""Direction""=1 and "
                    '                StrQuery += vbCrLf + "T.""ItemCode""='" & ItemCode & "' and T.""BaseEntry""in (" & DocEntry & ") )A GROUP BY A.""BatchNum"" having SUM(A.""Quantity"") >0"
                    '                objRecset.DoQuery(StrQuery)
                    '                Dim BQty As Double = CDbl(Matrix0.Columns.Item("9").Cells.Item(i).Specific.string)
                    '                If objRecset.RecordCount > 0 Then
                    '                    For j As Integer = 0 To objRecset.RecordCount - 1
                    '                        objGoodsReceipt.Lines.BatchNumbers.BatchNumber = CStr(objRecset.Fields.Item("BatchSerial").Value)
                    '                        objGoodsReceipt.Lines.BatchNumbers.Quantity = BQty
                    '                        objGoodsReceipt.Lines.BatchNumbers.Add()
                    '                    Next
                    '                End If
                    '            ElseIf Batchs = "N" And Serial = "Y" Then
                    '                StrQuery = "Select * from (SELECT distinct T4.""IntrSerial"" ""BatchSerial"",T4.""WhsCode"",T1.""DocEntry"",T1.""ItemCode"", T4.""Quantity"",T4.""Status"" from OIGE T0 left join IGE1 T1 on T0.""DocEntry""=T1.""DocEntry"""
                    '                StrQuery += vbCrLf + "left outer join SRI1 I1 on T1.""ItemCode""=I1.""ItemCode""   and (T1.""DocEntry""=I1.""BaseEntry"" and T1.""ObjType""=I1.""BaseType"") and T1.""LineNum""=I1.""BaseLinNum"""
                    '                StrQuery += vbCrLf + "left outer join OSRI T4 on T4.""ItemCode""=I1.""ItemCode"" and I1.""SysSerial""=T4.""SysSerial"" and I1.""WhsCode"" = T4.""WhsCode"" ) A "
                    '                StrQuery += vbCrLf + "Where A.""DocEntry"" in (" & DocEntry & ")  and A.""ItemCode""='" & ItemCode & "' and A.""BatchSerial"" <>'' and A.""Status""=1"
                    '                objRecset.DoQuery(StrQuery)
                    '                Dim SQty As Double = 0, TotSerialQty As Double = 0
                    '                SQty = CDbl(Matrix0.Columns.Item("9").Cells.Item(i).Specific.string)
                    '                If objRecset.RecordCount > 0 Then
                    '                    For j As Integer = 0 To objRecset.RecordCount - 1
                    '                        objGoodsReceipt.Lines.SerialNumbers.InternalSerialNumber = CStr(objRecset.Fields.Item("BatchSerial").Value)
                    '                        objGoodsReceipt.Lines.SerialNumbers.Quantity = CDbl(1)
                    '                        objGoodsReceipt.Lines.SerialNumbers.Add()
                    '                        TotSerialQty += CDbl(1)
                    '                        If SQty - TotSerialQty > 0 Then
                    '                            objRecset.MoveNext()
                    '                        Else
                    '                            Exit For
                    '                        End If
                    '                    Next
                    '                End If
                    '            End If
                    '            objGoodsReceipt.Lines.Add()
                    '            m_oProgBar.Value = i
                    '        Else
                    '            If objaddon.objcompany.InTransaction Then objaddon.objcompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                    '            objaddon.objapplication.StatusBar.SetText("Quantities Exceed from Goods Issue transaction.Please check...on the " & Trim(ItemCode) & " Line: " & CStr(i), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    '            Return False
                    '        End If
                    '    Else
                    '        If objaddon.objcompany.InTransaction Then objaddon.objcompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                    '        objaddon.objapplication.StatusBar.SetText("Additional Items found apart from Goods Issue transaction.Please check...on the " & Trim(ItemCode) & " Line: " & CStr(i), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    '        Return False
                    '    End If
                    'Next

                End If
                Retval = objGoodsReceipt.Add()
                If Retval <> 0 Then
                    If objaddon.objcompany.InTransaction Then objaddon.objcompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                    objaddon.objapplication.StatusBar.SetText("Goods Receipt : " & objaddon.objcompany.GetLastErrorDescription, SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(objGoodsReceipt)
                    Return False
                Else
                    If objaddon.objcompany.InTransaction Then objaddon.objcompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
                    TranEntry = objaddon.objcompany.GetNewObjectKey()
                    TransactionEntry = ""
                    Matrix0.Clear()
                    Matrix0.AddRow()
                    objform.Items.Item("trefno").Specific.String = ""
                    If objaddon.HANA Then
                        TranEntry = objaddon.objglobalmethods.getSingleValue("Select ""DocNum"" from OIGN where ""DocEntry""=" & TranEntry & "")
                        DocEntry = objaddon.objglobalmethods.getSingleValue("Select Count(*) from OIGE T0 join IGE1 T1 on T0.""DocEntry""=T1.""DocEntry"" where T0.""DocEntry""=" & DocEntry & "")
                    Else
                        TranEntry = objaddon.objglobalmethods.getSingleValue("Select DocNum from OIGN where DocEntry=" & TranEntry & "")
                        DocEntry = objaddon.objglobalmethods.getSingleValue("Select Count(*) from OIGE T0 join IGE1 T1 on T0.DocEntry""=T1.DocEntry where T0.DocEntry=" & DocEntry & "")
                    End If
                    m_oProgBar.Text = "Goods Receipt Created Successfully... Document Number->" & TranEntry & " Items Count- " & DocEntry
                    objaddon.objapplication.StatusBar.SetText("Goods Receipt Created Successfully...Document Number->" & TranEntry & " Items Count- " & DocEntry, SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                    objaddon.objapplication.MessageBox("Goods Receipt Created Successfully...Document Number->" & TranEntry & " Items Count- " & DocEntry, , "OK")
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(objGoodsReceipt)
                    GC.Collect()
                    Return True
                End If

            Catch ex As Exception
                objaddon.objapplication.StatusBar.SetText("Goods Receipt " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            Finally
                m_oProgBar.Stop()
                System.Runtime.InteropServices.Marshal.ReleaseComObject(m_oProgBar)
                m_oProgBar = Nothing
            End Try
        End Function

        Private Sub Button1_ClickBefore(sboObject As Object, pVal As SAPbouiCOM.SBOItemEventArg, ByRef BubbleEvent As Boolean) Handles Button1.ClickBefore
            Try
                If objform.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                    BubbleEvent = False
                    objaddon.objapplication.StatusBar.SetText("Disabled for adding the Document...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

                End If

            Catch ex As Exception

            End Try

        End Sub

        Private Sub Form_CloseAfter(pVal As SAPbouiCOM.SBOItemEventArg)
            Try
                FormCount -= 1
            Catch ex As Exception

            End Try

        End Sub

        Private Sub Button0_ClickAfter(sboObject As Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles Button0.ClickAfter
            Try
                If objform.Items.Item("trefno").Specific.String = "" Then Exit Sub
                If objaddon.objapplication.MessageBox("Goods Receipt Entry cannot be reversed. Do you want to continue?", 2, "Yes", "No") <> 1 Then Exit Sub
                If Create_GoodsReceipt_Updated(objform.Items.Item("trefno").Specific.String) Then
                    'objform.Items.Item("2").Click()
                End If
            Catch ex As Exception

            End Try

        End Sub

        Private WithEvents ComboBox0 As SAPbouiCOM.ComboBox

        Private Sub ComboBox0_ComboSelectBefore(sboObject As Object, pVal As SAPbouiCOM.SBOItemEventArg, ByRef BubbleEvent As Boolean) Handles ComboBox0.ComboSelectBefore
            Try
                BubbleEvent = False
                StrQuery = "Select ROW_NUMBER() OVER (order by T0.""DocEntry"" desc) as ""#"",T0.""DocEntry"",T0.""DocNum"",T0.""DocDate"",T0.""DocDueDate"",T0.""Comments"" as ""Remarks"",T0.""JrnlMemo"" as ""Journal Remark"",T0.""DocTotal"",T0.""BPLId"",T0.""BPLName"" as ""Branch Name"","
                StrQuery += vbCrLf + "T0.""U_TOWHS"" as ""To Whse"",T0.""U_TRANSTYPE"" as ""Trans Type"""
                StrQuery += vbCrLf + " from OIGE T0 Where T0.""DocStatus""='O' and T0.""U_TOWHS""<>'' and UPPER(T0.""U_TRANSTYPE"")='STOCK TRANSFER'"

                Dim activeform As New FrmOpenLists
                activeform.Show()
                DType = "GR"
                activeform.Load_Data(StrQuery, "GR")
                activeform.objform.Left = objform.Left + 100
                activeform.objform.Top = objform.Top + 100
            Catch ex As Exception

            End Try

        End Sub

        Private Sub Form_DataLoadAfter(ByRef pVal As SAPbouiCOM.BusinessObjectInfo)
            Try
                Dim oUDFForm As SAPbouiCOM.Form
                oUDFForm = objaddon.objapplication.Forms.Item(objform.UDFFormUID)
                If oUDFForm.Items.Item("U_RefNo").Enabled = True And oUDFForm.Items.Item("U_RefNo").Specific.String <> "" Then
                    oUDFForm.Items.Item("U_RefNo").Enabled = False
                Else
                    oUDFForm.Items.Item("U_RefNo").Enabled = True
                End If
            Catch ex As Exception

            End Try

        End Sub

        Private Sub Form_ClickBefore(pVal As SAPbouiCOM.SBOItemEventArg, ByRef BubbleEvent As Boolean)
            Try
                If bModal Then
                    BubbleEvent = False
                    Try
                        objaddon.objapplication.Forms.Item("OpenList").Select()
                    Catch ex As Exception
                    End Try
                End If
            Catch ex As Exception

            End Try

        End Sub

        Private WithEvents StaticText0 As SAPbouiCOM.StaticText
        Private WithEvents EditText0 As SAPbouiCOM.EditText
    End Class
End Namespace
