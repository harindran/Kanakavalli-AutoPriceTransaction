Imports System.IO
Imports SAPbouiCOM.Framework
Namespace AutoPriceTransaction
    Public Class ClsGRPO
        Public Const Formtype = "143"
        Dim objDelform, objUDFForm As SAPbouiCOM.Form
        'Dim ObjQCForm As SAPbouiCOM.Form
        'Dim ObjGRPOForm As SAPbouiCOM.Form
        Dim objmatrix As SAPbouiCOM.Matrix
        Dim StrQuery As String
        Dim strSQL As String
        Public WithEvents objDelformUDF As SAPbouiCOM.Form
        Dim objRs As SAPbobsCOM.Recordset
        Dim objCombo As SAPbouiCOM.ComboBox

        Public Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
            Try
                objDelform = objaddon.objapplication.Forms.Item(FormUID)
                objmatrix = objDelform.Items.Item("38").Specific
                If pVal.BeforeAction Then
                    Select Case pVal.EventType
                        Case SAPbouiCOM.BoEventTypes.et_VALIDATE
                            If pVal.ItemUID = "4" Then
                                If objDelform.Items.Item("4").Specific.String = "" Then Exit Sub
                                If objDelform.Mode <> SAPbouiCOM.BoFormMode.fm_ADD_MODE Then Exit Sub
                                Field_Settings(FormUID, objDelform.Items.Item("4").Specific.String)
                                VendorCode = objDelform.Items.Item("4").Specific.String
                            End If

                        Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT

                        Case SAPbouiCOM.BoEventTypes.et_CLICK
                            Try
                                If bModal Then
                                    BubbleEvent = False
                                    Try
                                        objaddon.objapplication.Forms.Item("OpenList").Select()
                                    Catch ex As Exception
                                    End Try
                                End If
                                If pVal.ItemUID = "btnsave" Then
                                    If objDelform.Items.Item("btnsave").Enabled = False Then Exit Sub
                                    If objDelform.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE Then Exit Sub
                                    'Dim ErrorFlag As Boolean
                                    Dim ItemCode As String = ""
                                    'Dim Qty As Double

                                    'objaddon.objapplication.StatusBar.SetText("Validating Please wait...", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                    If objDelform.Mode <> SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                                        BubbleEvent = False : Exit Sub
                                    End If
                                    If objDelform.Items.Item("trefno").Specific.String = "" Then
                                        objaddon.objapplication.StatusBar.SetText("Please select a A/R Invoice Entries...", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                        Exit Sub
                                    End If
                                    If objDelform.Items.Item("10").Specific.String = "" Then
                                        objaddon.objapplication.StatusBar.SetText("Posting Date is Missing", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                        BubbleEvent = False : Exit Sub
                                    End If
                                    If objDelform.Items.Item("12").Specific.String = "" Then
                                        objaddon.objapplication.StatusBar.SetText("Due Date is Missing", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                        BubbleEvent = False : Exit Sub
                                    End If
                                    If objDelform.Items.Item("46").Specific.String = "" Then
                                        objaddon.objapplication.StatusBar.SetText("Document Date is Missing", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                        BubbleEvent = False : Exit Sub
                                    End If
                                    TransactionEntry = objDelform.Items.Item("trefno").Specific.String
                                End If

                                'objRs = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                                'For j As Integer = 1 To objmatrix.VisualRowCount
                                '    ItemCode = objmatrix.Columns.Item("1").Cells.Item(j).Specific.string
                                '    Qty = CDbl(objmatrix.Columns.Item("11").Cells.Item(j).Specific.string)
                                '    If ItemCode = "" Then Continue For
                                '    StrQuery = "Select T0.""BPLId"",T0.""U_TRANSTYPE"",T0.""U_TOWHS"",T0.""Comments"",T1.""DocEntry"",T1.""LineNum"",T1.""ItemCode"",T1.""Quantity"",T1.""WhsCode"",T1.""AcctCode"",T1.""StockPrice"",T1.""LocCode"""
                                '    StrQuery += vbCrLf + " from OINV T0 join INV1 T1 on T0.""DocEntry""=T1.""DocEntry"" where ifnull(T1.""LineStatus"",'O')='O' and T0.""DocEntry"" in (" & TransactionEntry & ")  and T1.""ItemCode""='" & ItemCode & "' order by T0.""DocNum"",T1.""LineNum""  "
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
                                'If Create_GoodsReceipt() Then
                                '    objDelform.Items.Item("2").Click()
                                'Else
                                '    BubbleEvent = False
                                'End If

                            Catch ex As Exception

                            End Try
                    End Select
                Else
                    Select Case pVal.EventType
                        Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                            If pVal.ActionSuccess Then
                                CreateButton(FormUID)
                            End If
                        Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT
                            Try
                                If pVal.ItemUID = "10000330" Then
                                    objCombo = objDelform.Items.Item("10000330").Specific
                                    If objCombo.Selected.Value = "AR Invoice" Then
                                        strSQL = "Select ROW_NUMBER() OVER (order by T0.""DocEntry"" desc) as ""#"",T0.""DocEntry"",T0.""DocNum"",T0.""DocDate"",T0.""DocDueDate"",T0.""Comments"" as ""Remarks"",T0.""JrnlMemo"" as ""Journal Remark"",T0.""DocTotal"",T0.""BPLId"",T0.""BPLName"" as ""Branch Name"","
                                        strSQL += vbCrLf + "T0.""U_TOWHS"" as ""To Whse"",T0.""U_TRANSTYPE"" as ""Trans Type"",T0.""Comments"" as ""Remarks"""
                                        strSQL += vbCrLf + " from OINV T0 left join OPDN T1 on T0.""DocEntry""=T1.""U_RefNo"" Where T0.""DocStatus""='O' and T0.""CANCELED""='N' and T0.""U_TOWHS""<>'' and T1.""U_RefNo"" is null and UPPER(T0.""U_TRANSTYPE"")='STOCK TRANSFER' " 'T0.""NumAtCard"" as ""Ref No.""
                                        strSQL += vbCrLf + "And T0.""CardCode""=(Select distinct ""U_CUSTOMER"" from ""@MIPL_STBP"""
                                        strSQL += vbCrLf + "where ""U_FRBRANCH""=(Select ""BPLid"" from OWHS Where ""WhsCode""=T0.""U_TOWHS"") and ""U_VENDOR""='" & objDelform.Items.Item("4").Specific.String & "')"

                                        Dim activeform As New FrmOpenLists
                                        activeform.Show()
                                        DType = "GRPO"
                                        activeform.Load_Data(strSQL, "GRPO")
                                        activeform.objform.Left = objDelform.Left + 100
                                        activeform.objform.Top = objDelform.Top + 100
                                    End If
                                End If
                            Catch ex As Exception

                            End Try
                        Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                            Try
                                If pVal.ActionSuccess = False Then Exit Sub

                                objCombo = objDelform.Items.Item("10000330").Specific
                                objCombo.ValidValues.Add("AR Invoice", "AR Invoice")
                                objCombo.ExpandType = SAPbouiCOM.BoExpandType.et_ValueOnly
                            Catch ex As Exception

                            End Try
                        Case SAPbouiCOM.BoEventTypes.et_CLICK
                            Try
                                If pVal.ItemUID = "btnsave" And objDelform.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                                    If objDelform.Items.Item("btnsave").Enabled = False Then Exit Sub
                                    Dim objcomb As SAPbouiCOM.ComboBox
                                    objcomb = objDelform.Items.Item("3").Specific
                                    If objcomb.Selected.Value = "S" Then
                                        objaddon.objapplication.StatusBar.SetText("Service Type found.Please Check...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                        Exit Sub
                                    End If
                                    'If Trim(objmatrix.Columns.Item("1").Cells.Item(1).Specific.string) = "" Then Exit Sub
                                    If objDelform.Items.Item("trefno").Specific.String = "" Then Exit Sub
                                    If objaddon.objapplication.MessageBox("GRPO Transaction cannot be reversed. Do you want to continue?", 2, "Yes", "No") <> 1 Then Exit Sub
                                    If Create_GRPO(FormUID, objDelform.Items.Item("trefno").Specific.String) Then
                                        'objDelform.Items.Item("2").Click()
                                    End If
                                End If

                            Catch ex As Exception
                            End Try
                    End Select
                End If
            Catch ex As Exception
            End Try

        End Sub

        Public Sub FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
            Try
                objDelform = objaddon.objapplication.Forms.Item(BusinessObjectInfo.FormUID)
                If BusinessObjectInfo.BeforeAction = True Then
                    Select Case BusinessObjectInfo.EventType
                        Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD, SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE
                            If objDelform.Title.ToUpper = "GOODS RECEIPT PO - CANCELLATION" Then
                                objUDFForm = objaddon.objapplication.Forms.Item(objDelform.UDFFormUID)
                                If objUDFForm.Items.Item("U_RefNo").Specific.String <> "" Then
                                    If objUDFForm.Items.Item("U_RefNo").Enabled = False Then objUDFForm.Items.Item("U_RefNo").Enabled = True : objUDFForm.Items.Item("U_RefNo").Specific.String = ""
                                End If
                            End If
                        Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD

                    End Select
                Else
                    Select Case BusinessObjectInfo.EventType
                        Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD, SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE
                            If objDelform.Title.ToUpper = "GOODS RECEIPT PO - CANCELLATION" Then
                                strSQL = "Update OPDN Set ""U_RefNo""=null Where ""DocEntry"" in (Select T0.""BaseEntry"" from PDN1 T0 Left Join OPDN T1 On T1.""DocEntry""=T0.""DocEntry"" Where T0.""DocEntry"" =" & objDelform.DataSources.DBDataSources.Item("OPDN").GetValue("DocEntry", 0) & " and T1.""CANCELED""='C')"
                                objRs = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                objRs.DoQuery(strSQL)
                            End If
                        Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD
                            Try
                                Field_Settings(objDelform.UniqueID, objDelform.Items.Item("4").Specific.String)
                                objDelform.Items.Item("trefno").Enabled = False
                                Dim oUDFForm As SAPbouiCOM.Form
                                oUDFForm = objaddon.objapplication.Forms.Item(objDelform.UDFFormUID)
                                If oUDFForm.Items.Item("U_RefNo").Enabled = True And oUDFForm.Items.Item("U_RefNo").Specific.String <> "" Then
                                    oUDFForm.Items.Item("U_RefNo").Enabled = False
                                Else
                                    oUDFForm.Items.Item("U_RefNo").Enabled = True
                                End If
                            Catch ex As Exception
                            End Try
                    End Select
                End If

            Catch ex As Exception
                objaddon.objapplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            End Try

        End Sub

        Public Sub CreateButton(ByVal FormUID As String)
            Try
                Dim objButton As SAPbouiCOM.Button
                Dim objItem As SAPbouiCOM.Item
                objDelform = objaddon.objapplication.Forms.Item(FormUID)
                Try
                    If objDelform.Items.Item("btnsave").UniqueID = "btnsave" Then Exit Sub
                Catch ex As Exception
                End Try
                objItem = objDelform.Items.Add("btnsave", SAPbouiCOM.BoFormItemTypes.it_BUTTON)
                objItem.Left = objDelform.Items.Item("2").Left + objDelform.Items.Item("2").Width + 10
                'objItem.Left = objDelform.Items.Item("10002056").Left + objDelform.Items.Item("10002056").Width + 60
                objItem.Width = 65
                objItem.Top = objDelform.Items.Item("2").Top
                objItem.Height = objDelform.Items.Item("2").Height
                objButton = objItem.Specific
                objButton.Caption = "Save"


                Dim objLabel As SAPbouiCOM.StaticText
                'Dim objItem As SAPbouiCOM.Item
                objDelform = objaddon.objapplication.Forms.Item(FormUID)
                objItem = objDelform.Items.Add("lrefno", SAPbouiCOM.BoFormItemTypes.it_STATIC)
                objItem.Left = objDelform.Items.Item("2002").Left
                objItem.Width = 80
                objItem.Top = objDelform.Items.Item("2002").Top + objDelform.Items.Item("2002").Height + 2
                objItem.Height = 14
                objLabel = objItem.Specific
                objLabel.Caption = "Ref No."

                Dim objedit As SAPbouiCOM.EditText
                objItem = objDelform.Items.Add("trefno", SAPbouiCOM.BoFormItemTypes.it_EDIT)
                objItem.Left = objDelform.Items.Item("2003").Left
                objItem.Width = 80 '
                objItem.Top = objDelform.Items.Item("2003").Top + objDelform.Items.Item("2003").Height + 2
                objItem.Height = 14
                objItem.LinkTo = "lrefno"
                objedit = objItem.Specific
                objedit.Item.Enabled = False
                objItem.Enabled = False
                objedit.DataBind.SetBound(True, "OPDN", "U_RefNo")
                'objAddOn.objApplication.SetStatusBarMessage("Button Created", SAPbouiCOM.BoMessageTime.bmt_Short, False)
            Catch ex As Exception
            End Try

        End Sub

        Private Sub Field_Settings(ByVal FormUID As String, ByVal CardCode As String)
            Try
                'If CardCode = "" Then Exit Sub
                objDelform = objaddon.objapplication.Forms.Item(FormUID)
                strSQL = "Select 1 as ""Status"" from OCRD T0 left join OCRG T1 on T0.""GroupCode""=T1.""GroupCode"""
                strSQL += vbCrLf + "where T1.""GroupType""='S' and T1.""GroupName"" like 'Group%' and T0.""CardCode""='" & Trim(CardCode) & "'"
                CreateButton(FormUID)
                strSQL = objaddon.objglobalmethods.getSingleValue(strSQL)
                objmatrix = objDelform.Items.Item("38").Specific
                If strSQL = "1" Then
                    If objDelform.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then objDelform.Items.Item("btnsave").Enabled = True Else objDelform.Items.Item("btnsave").Enabled = False
                    'objDelform.Items.Item("btnsave").Enabled = True
                    objDelform.Items.Item("1").Enabled = False
                    objmatrix.Columns.Item("14").Visible = False 'Unit Price
                    objmatrix.Columns.Item("21").Visible = False  'Line Total
                    objmatrix.Columns.Item("259").Visible = False  'Item Cost

                Else
                    objDelform.Items.Item("btnsave").Enabled = False
                    objDelform.Items.Item("1").Enabled = True
                    objmatrix.Columns.Item("14").Visible = True 'Unit Price
                    objmatrix.Columns.Item("21").Visible = True  'Line Total
                    objmatrix.Columns.Item("259").Visible = True  'Item Cost
                End If
            Catch ex As Exception

            End Try
        End Sub

        Private Function Create_GRPO(ByVal FormUID As String, ByVal DocEntry As String) As Boolean
            Dim BranchEnabled, Series, TranEntry As String
            Dim objGRPO As SAPbobsCOM.Documents
            Dim objEdit As SAPbouiCOM.EditText
            'Dim MBAPDocNum As Long
            Dim TFlag As Boolean = False
            Dim objRs1 As SAPbobsCOM.Recordset
            Dim m_oProgBar As SAPbouiCOM.ProgressBar
            Dim ItemCode As String = ""
            Dim Qty As Double
            Try
                objDelform = objaddon.objapplication.Forms.Item(FormUID)
                objRs = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                objRs1 = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                objmatrix = objDelform.Items.Item("38").Specific
                If objaddon.HANA Then
                    BranchEnabled = objaddon.objglobalmethods.getSingleValue("select ""MltpBrnchs"" from OADM")
                Else
                    BranchEnabled = objaddon.objglobalmethods.getSingleValue("select MltpBrnchs from OADM")
                End If
                If Not BranchEnabled = "Y" Then objaddon.objapplication.StatusBar.SetText("Branch not enabled...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning) : Exit Function
                objaddon.objapplication.StatusBar.SetText("GRPO Transaction Creating Please wait...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                objGRPO = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseDeliveryNotes)
                StrQuery = "Select T0.""CardCode"",T0.""BPLId"",(Select ""BPLid"" from OWHS Where ""WhsCode""=T0.""U_TOWHS"") as ""To_Branch"",T0.""U_TRANSTYPE"",T0.""U_TOWHS"",T0.""U_FROMWHS"",T0.""NumAtCard"","
                StrQuery += vbCrLf + "T0.""Comments"",T0.""DocNum"",T1.""DocEntry"",T1.""LineNum"",T1.""ItemCode"",T1.""Project"",T1.""Quantity"",T1.""TaxCode"",T1.""WhsCode"","
                StrQuery += vbCrLf + "(Select ""State""||'IGST'||(Select case when ""U_SALESTAX"" is null then 0 else ""U_SALESTAX"" End From OITM where ""ItemCode""=T1.""ItemCode"") from OWHS where ""WhsCode""=T0.""U_TOWHS"")as ""ATaxCode"","
                StrQuery += vbCrLf + "T1.""AcctCode"",T1.""StockPrice"",T1.""LocCode"",T0.""Address"" ""BillTo"",T0.""Address2"" ""ShipTo"",T1.""U_MRP"""
                StrQuery += vbCrLf + "from OINV T0 join INV1 T1 on T0.""DocEntry""=T1.""DocEntry"" where"
                StrQuery += vbCrLf + " T0.""DocType""='I' and T0.""DocEntry"" in (" & DocEntry & ") and T0.""CardCode""=(Select distinct ""U_CUSTOMER"" from ""@MIPL_STBP"""
                StrQuery += vbCrLf + "where ""U_FRBRANCH""=(Select ""BPLid"" from OWHS Where ""WhsCode""=T0.""U_TOWHS"") and ""U_VENDOR""='" & objDelform.Items.Item("4").Specific.String & "')"
                StrQuery += vbCrLf + "order by T0.""DocNum"",T1.""LineNum""  "
                objRs.DoQuery(StrQuery)
                If objRs.RecordCount = 0 Then Exit Function
                'm_oProgBar = objaddon.objapplication.StatusBar.CreateProgressBar("My Progress", objmatrix.VisualRowCount, True)
                m_oProgBar = objaddon.objapplication.StatusBar.CreateProgressBar("My Progress", objRs.RecordCount, True)
                m_oProgBar.Text = "GRPO Transaction Creating Please wait... "
                m_oProgBar.Value = 0
                If Not objaddon.objcompany.InTransaction Then objaddon.objcompany.StartTransaction()
                'Dim DocDate As Date = Date.ParseExact(EditText1.Value, "yyyyMMdd", System.Globalization.DateTimeFormatInfo.InvariantInfo)
                objEdit = objDelform.Items.Item("10").Specific
                Dim DocDate As Date = Date.ParseExact(objEdit.Value, "yyyyMMdd", System.Globalization.DateTimeFormatInfo.InvariantInfo)
                objEdit = objDelform.Items.Item("12").Specific
                Dim DueDate As Date = Date.ParseExact(objEdit.Value, "yyyyMMdd", System.Globalization.DateTimeFormatInfo.InvariantInfo)
                objEdit = objDelform.Items.Item("46").Specific
                Dim TaxDate As Date = Date.ParseExact(objEdit.Value, "yyyyMMdd", System.Globalization.DateTimeFormatInfo.InvariantInfo)
                'MBAPDocNum = objDelform.BusinessObject.GetNextSerialNumber(objDelform.Items.Item("Series").Specific.Selected.value)
                objGRPO.CardCode = objDelform.Items.Item("4").Specific.String
                objGRPO.DocDate = DocDate
                objGRPO.DocDueDate = DueDate
                objGRPO.TaxDate = TaxDate
                objGRPO.NumAtCard = objRs.Fields.Item("NumAtCard").Value.ToString & "-" & objDelform.Items.Item("14").Specific.String
                objGRPO.JournalMemo = "Auto-Generated->  " & Now.ToString
                objGRPO.Comments = "Based on A/R Invoice Baseref " & objRs.Fields.Item("DocNum").Value.ToString & "-Created by " & objaddon.objcompany.UserName & " on " & Now.ToString & "-" & objDelform.Items.Item("16").Specific.String
                If objaddon.HANA Then
                    Series = objaddon.objglobalmethods.getSingleValue("SELECT Top 1 ""Series"" FROM NNM1 WHERE ""ObjectCode""='20' and ""Indicator""=(Select ""Indicator"" from OFPR where CURRENT_DATE Between ""F_RefDate"" and ""T_RefDate"") " &
                                                                                   " and ""BPLId""='" & objRs.Fields.Item("To_Branch").Value.ToString & "'")
                Else
                    Series = objaddon.objglobalmethods.getSingleValue("SELECT Top 1 Series FROM NNM1 WHERE ObjectCode='20' and Indicator=(Select Indicator from OFPR where CURRENT_DATE Between F_RefDate and T_RefDate) " &
                                                                                  " and BPLId='" & objRs.Fields.Item("To_Branch").Value.ToString & "'")
                End If

                If Series = "" Then objaddon.objapplication.StatusBar.SetText("Numbering Series not found.Please check...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning) : Exit Function
                objGRPO.Series = Series
                'objPurchaseInvoice.Series = 110
                'Dim oUDFForm As SAPbouiCOM.Form
                'oUDFForm = objaddon.objapplication.Forms.Item(objDelform.UDFFormUID)
                objGRPO.UserFields.Fields.Item("U_FROMWHS").Value = objRs.Fields.Item("U_FROMWHS").Value.ToString
                objGRPO.UserFields.Fields.Item("U_TOWHS").Value = objRs.Fields.Item("U_TOWHS").Value.ToString
                objGRPO.UserFields.Fields.Item("U_TRANSTYPE").Value = objRs.Fields.Item("U_TRANSTYPE").Value.ToString
                objGRPO.UserFields.Fields.Item("U_RefNo").Value = DocEntry
                If BranchEnabled = "Y" Then
                    objGRPO.BPL_IDAssignedToInvoice = objRs.Fields.Item("To_Branch").Value.ToString
                End If
                objGRPO.DocType = SAPbobsCOM.BoDocumentTypes.dDocument_Items
                'objGRPO.Address = objRs.Fields.Item("BillTo").Value.ToString 'Bill To
                'objGRPO.Address2 = objRs.Fields.Item("ShipTo").Value.ToString 'Ship To

                objGRPO.PayToCode = objDelform.Items.Item("226").Specific.Selected.Value 'Bill To
                'objDelivery.Address = objDelform.Items.Item("6").Specific.String 'Bill To
                objRs1.DoQuery("select  ""Address"",""CardCode"",""Street"",""Block"",""ZipCode"",""City"",""County"",""Country"",""State"",""Building"",""AdresType"",""Address2"",""Address3"",""AddrType"",""StreetNo""  from CRD1 where ""CardCode""='" & objDelform.Items.Item("4").Specific.String & "' and ""Address""='" & objDelform.Items.Item("226").Specific.Selected.Value & "' and ""AdresType""='B'")

                objGRPO.AddressExtension.BillToAddress2 = Trim(objRs1.Fields.Item("Address2").Value.ToString)
                objGRPO.AddressExtension.BillToAddress3 = Trim(objRs1.Fields.Item("Address3").Value.ToString)
                objGRPO.AddressExtension.BillToAddressType = Trim(objRs1.Fields.Item("AddrType").Value.ToString)
                objGRPO.AddressExtension.BillToBlock = Trim(objRs1.Fields.Item("Block").Value.ToString)
                objGRPO.AddressExtension.BillToBuilding = Trim(objRs1.Fields.Item("Building").Value.ToString)
                objGRPO.AddressExtension.BillToCity = Trim(objRs1.Fields.Item("City").Value.ToString)
                objGRPO.AddressExtension.BillToCountry = Trim(objRs1.Fields.Item("Country").Value.ToString)
                objGRPO.AddressExtension.BillToCounty = Trim(objRs1.Fields.Item("County").Value.ToString)
                objGRPO.AddressExtension.BillToState = Trim(objRs1.Fields.Item("State").Value.ToString)
                objGRPO.AddressExtension.BillToStreet = Trim(objRs1.Fields.Item("Street").Value.ToString)
                objGRPO.AddressExtension.BillToStreetNo = Trim(objRs1.Fields.Item("StreetNo").Value.ToString)
                objGRPO.AddressExtension.BillToZipCode = Trim(objRs1.Fields.Item("ZipCode").Value.ToString)

                'objGRPO.ShipToCode = objDelform.Items.Item("40").Specific.Selected.Value 'Ship To
                'objDelivery.Address2 = objDelform.Items.Item("92").Specific.String 'Ship To
                'objRs1.DoQuery("select  ""Address"",""CardCode"",""Street"",""Block"",""ZipCode"",""City"",""County"",""Country"",""State"",""Building"",""AdresType"",""Address2"",""Address3"",""AddrType"",""StreetNo""  from CRD1 where ""CardCode""='" & objDelform.Items.Item("4").Specific.String & "' and ""Address""='" & objDelform.Items.Item("226").Specific.Selected.Value & "' and ""AdresType""='S'")
                objRs1.DoQuery("Select T0.""Street"",T0.""Block"",T0.""ZipCode"",T0.""City"",T0.""County"",T0.""Country"",T0.""State"",T0.""Building"",T0.""Address2"",T0.""Address3"",T0.""AddrType"",T0.""StreetNo"" from OWHS T0 join OLCT T1 on T0.""Location""=T1.""Code"" Where T0.""WhsCode""='" & Trim(objRs.Fields.Item("U_TOWHS").Value.ToString) & "'")
                objGRPO.AddressExtension.ShipToAddress2 = Trim(objRs1.Fields.Item("Address2").Value.ToString)
                objGRPO.AddressExtension.ShipToAddress3 = Trim(objRs1.Fields.Item("Address3").Value.ToString)
                objGRPO.AddressExtension.ShipToAddressType = Trim(objRs1.Fields.Item("AddrType").Value.ToString)
                objGRPO.AddressExtension.ShipToBlock = Trim(objRs1.Fields.Item("Block").Value.ToString)
                objGRPO.AddressExtension.ShipToBuilding = Trim(objRs1.Fields.Item("Building").Value.ToString)
                objGRPO.AddressExtension.ShipToCity = Trim(objRs1.Fields.Item("City").Value.ToString)
                objGRPO.AddressExtension.ShipToCountry = Trim(objRs1.Fields.Item("Country").Value.ToString)
                objGRPO.AddressExtension.ShipToCounty = Trim(objRs1.Fields.Item("County").Value.ToString)
                objGRPO.AddressExtension.ShipToState = Trim(objRs1.Fields.Item("State").Value.ToString)
                objGRPO.AddressExtension.ShipToStreet = Trim(objRs1.Fields.Item("Street").Value.ToString)
                objGRPO.AddressExtension.ShipToStreetNo = Trim(objRs1.Fields.Item("StreetNo").Value.ToString)
                objGRPO.AddressExtension.ShipToZipCode = Trim(objRs1.Fields.Item("ZipCode").Value.ToString)


                'For Rec As Integer = 0 To objRs.RecordCount - 1
                '    For Row As Integer = 1 To objmatrix.VisualRowCount
                '        If objmatrix.Columns.Item("1").Cells.Item(Row).Specific.string <> "" And objmatrix.Columns.Item("1").Cells.Item(Row).Specific.string = objRs.Fields.Item("ItemCode").Value.ToString Then
                '            objGRPO.Lines.ItemCode = Trim(objRs.Fields.Item("ItemCode").Value.ToString)
                '            'objGRPO.Lines.ItemDescription = Trim(objRs.Fields.Item("Quantity").Value.ToString)
                '            objGRPO.Lines.Quantity = Trim(CDbl(objRs.Fields.Item("Quantity").Value.ToString))
                '            objGRPO.Lines.AccountCode = Trim(objRs.Fields.Item("AcctCode").Value.ToString)
                '            objGRPO.Lines.TaxCode = Trim(objRs.Fields.Item("TaxCode").Value.ToString)
                '            objGRPO.Lines.UnitPrice = Trim(objRs.Fields.Item("StockPrice").Value.ToString)
                '            objGRPO.Lines.WarehouseCode = Trim(objRs.Fields.Item("U_TOWHS").Value.ToString)
                '            objGRPO.Lines.ProjectCode = Trim(objRs.Fields.Item("Project").Value.ToString)
                '            objGRPO.Lines.UserFields.Fields.Item("U_MRP").Value = Trim(objRs.Fields.Item("U_MRP").Value.ToString)
                '            'objGRPO.Lines.LineTotal = Matrix0.Columns.Item("total").Cells.Item(Row).Specific.string
                '            objGRPO.Lines.Add()
                '        End If
                '    Next
                '    objRs.MoveNext()
                'Next
                Dim Loc As String
                If objaddon.HANA Then
                    Loc = objaddon.objglobalmethods.getSingleValue("select T0.""Code"" from OLCT T0 join OWHS T1 on T0.""Code""=T1.""Location"" where T1.""WhsCode""='" & Trim(objRs.Fields.Item("U_TOWHS").Value.ToString) & "'")
                Else
                    Loc = objaddon.objglobalmethods.getSingleValue("select T0.Code from OLCT T0 join OWHS T1 on T0.Code=T1.Location where T1.WhsCode='" & Trim(objRs.Fields.Item("U_TOWHS").Value.ToString) & "'")
                End If
                Dim iRow As Integer = 0
                While Not objRs.EoF
                    ItemCode = Trim(objRs.Fields.Item("ItemCode").Value.ToString)
                    Qty = Trim(CDbl(objRs.Fields.Item("Quantity").Value.ToString))
                    objGRPO.Lines.ItemCode = ItemCode ' Trim(objRs1.Fields.Item("ItemCode").Value.ToString)
                    'objGRPO.Lines.ItemDescription = Trim(objRs.Fields.Item("Quantity").Value.ToString)
                    objGRPO.Lines.Quantity = Qty ' CDbl(objmatrix.Columns.Item("11").Cells.Item(Row).Specific.string) 'Trim(CDbl(objRs1.Fields.Item("Quantity").Value.ToString))
                    objGRPO.Lines.AccountCode = Trim(objRs.Fields.Item("AcctCode").Value.ToString)
                    objGRPO.Lines.TaxCode = Trim(objRs.Fields.Item("ATaxCode").Value.ToString) 'Trim(objRs1.Fields.Item("TaxCode").Value.ToString)
                    objGRPO.Lines.UnitPrice = Trim(objRs.Fields.Item("StockPrice").Value.ToString)
                    objGRPO.Lines.WarehouseCode = Trim(objRs.Fields.Item("U_TOWHS").Value.ToString)

                    If Loc <> "" Then objGRPO.Lines.LocationCode = Loc
                    objGRPO.Lines.ProjectCode = Trim(objRs.Fields.Item("Project").Value.ToString)
                    objGRPO.Lines.UserFields.Fields.Item("U_MRP").Value = Trim(objRs.Fields.Item("U_MRP").Value.ToString)
                    objGRPO.Lines.UserFields.Fields.Item("U_DocLine").Value = Trim(objRs.Fields.Item("LineNum").Value.ToString)
                    objGRPO.Lines.Add()
                    iRow += 1
                    m_oProgBar.Value = iRow
                    objRs.MoveNext()
                End While


                'For Row As Integer = 1 To objmatrix.VisualRowCount
                '    ItemCode = Trim(objmatrix.Columns.Item("1").Cells.Item(Row).Specific.string)
                '    Qty = CDbl(objmatrix.Columns.Item("11").Cells.Item(Row).Specific.string)
                '    If ItemCode = "" Then Continue For
                '    StrQuery = "Select T0.""CardCode"",T0.""BPLId"",(Select ""BPLid"" from OWHS Where ""WhsCode""=T0.""U_TOWHS"") as ""To_Branch"",T0.""U_TRANSTYPE"",T0.""U_TOWHS"","
                '    StrQuery += vbCrLf + "T0.""Comments"",T0.""DocNum"",T1.""DocEntry"",T1.""LineNum"",T1.""ItemCode"",T1.""Project"",T1.""Quantity"",T1.""TaxCode"",T1.""WhsCode"","
                '    StrQuery += vbCrLf + "T1.""AcctCode"",T1.""StockPrice"",T1.""LocCode"",T0.""Address"" ""BillTo"",T0.""Address2"" ""ShipTo"",T1.""U_MRP"""
                '    StrQuery += vbCrLf + ",(Select ""State""||'IGST'||(Select case when ""U_SALESTAX"" is null then 0 else ""U_SALESTAX"" End From OITM where ""ItemCode""=T1.""ItemCode"") from OWHS where ""WhsCode""=T0.""U_TOWHS"")as ""ATaxCode"""
                '    StrQuery += vbCrLf + "from OINV T0 join INV1 T1 on T0.""DocEntry""=T1.""DocEntry"" where"
                '    StrQuery += vbCrLf + " T0.""DocType""='I' and T0.""DocEntry"" in (" & DocEntry & ") and T1.""ItemCode""='" & ItemCode & "' and T1.""LineNum""='" & IIf(Trim(objmatrix.Columns.Item("U_DocLine").Cells.Item(Row).Specific.string) = "", "-1", Trim(objmatrix.Columns.Item("U_DocLine").Cells.Item(Row).Specific.string)) & "' and T0.""CardCode""=(Select distinct ""U_CUSTOMER"" from ""@MIPL_STBP"""
                '    StrQuery += vbCrLf + "where ""U_FRBRANCH""=(Select ""BPLid"" from OWHS Where ""WhsCode""=T0.""U_TOWHS"") and ""U_VENDOR""='" & objDelform.Items.Item("4").Specific.String & "')"
                '    StrQuery += vbCrLf + "order by T0.""DocNum"",T1.""LineNum""  "
                '    objRs1.DoQuery(StrQuery)
                '    If objRs1.RecordCount > 0 Then
                '        If Qty <= CDbl(objRs.Fields.Item("Quantity").Value) Then
                '            objGRPO.Lines.ItemCode = ItemCode ' Trim(objRs1.Fields.Item("ItemCode").Value.ToString)
                '            'objGRPO.Lines.ItemDescription = Trim(objRs.Fields.Item("Quantity").Value.ToString)
                '            objGRPO.Lines.Quantity = Qty ' CDbl(objmatrix.Columns.Item("11").Cells.Item(Row).Specific.string) 'Trim(CDbl(objRs1.Fields.Item("Quantity").Value.ToString))
                '            objGRPO.Lines.AccountCode = Trim(objRs1.Fields.Item("AcctCode").Value.ToString)
                '            objGRPO.Lines.TaxCode = Trim(objRs1.Fields.Item("ATaxCode").Value.ToString) 'Trim(objRs1.Fields.Item("TaxCode").Value.ToString)
                '            objGRPO.Lines.UnitPrice = Trim(objRs1.Fields.Item("StockPrice").Value.ToString)
                '            objGRPO.Lines.WarehouseCode = Trim(objRs1.Fields.Item("U_TOWHS").Value.ToString)
                '            objGRPO.Lines.ProjectCode = Trim(objRs1.Fields.Item("Project").Value.ToString)
                '            objGRPO.Lines.UserFields.Fields.Item("U_MRP").Value = Trim(objRs1.Fields.Item("U_MRP").Value.ToString)
                '            objGRPO.Lines.UserFields.Fields.Item("U_DocLine").Value = Trim(objRs1.Fields.Item("LineNum").Value.ToString)
                '            'objGRPO.Lines.LineTotal = Matrix0.Columns.Item("total").Cells.Item(Row).Specific.string
                '            objGRPO.Lines.Add()
                '            m_oProgBar.Value = Row
                '        Else
                '            If objaddon.objcompany.InTransaction Then objaddon.objcompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                '            objaddon.objapplication.StatusBar.SetText("Quantities Exceed from A/R Invoice transaction.Please check...on the " & Trim(ItemCode) & " Line: " & CStr(Row), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                '            Return False
                '        End If

                '    Else
                '        If objaddon.objcompany.InTransaction Then objaddon.objcompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                '        objaddon.objapplication.StatusBar.SetText("Additional Items found apart from A/R Invoice transaction.Please check...on the " & Trim(ItemCode) & " Line: " & CStr(Row), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                '        Return False
                '    End If
                'Next

                If objGRPO.Add() <> 0 Then
                    TFlag = True
                    If objaddon.objcompany.InTransaction Then objaddon.objcompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                    objaddon.objapplication.SetStatusBarMessage("GRPO: " & objaddon.objcompany.GetLastErrorDescription & "-" & objaddon.objcompany.GetLastErrorCode, SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                    objaddon.objapplication.MessageBox("GRPO: " & objaddon.objcompany.GetLastErrorDescription & "-" & objaddon.objcompany.GetLastErrorCode,, "OK")
                Else
                    'Dim sNewObjCode As String = ""
                    'objaddon.objcompany.GetNewObjectCode(sNewObjCode)
                    'Dim str = CLng(sNewObjCode)
                    If objaddon.objcompany.InTransaction Then objaddon.objcompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
                    TranEntry = objaddon.objcompany.GetNewObjectKey()
                    If objaddon.HANA Then
                        TranEntry = objaddon.objglobalmethods.getSingleValue("Select ""DocNum"" from OPDN where ""DocEntry""=" & TranEntry & "")
                        DocEntry = objaddon.objglobalmethods.getSingleValue("Select Count(*) from OINV T0 join INV1 T1 on T0.""DocEntry""=T1.""DocEntry"" where T0.""DocEntry""=" & DocEntry & "")
                    Else
                        TranEntry = objaddon.objglobalmethods.getSingleValue("Select DocNum from OPDN where DocEntry=" & TranEntry & "")
                        DocEntry = objaddon.objglobalmethods.getSingleValue("Select Count(*) from OINV T0 join INV1 T1 on T0.DocEntry""=T1.DocEntry where T0.DocEntry=" & DocEntry & "")

                    End If
                    objmatrix.Clear()
                    objmatrix.AddRow()
                    objDelform.Items.Item("4").Specific.String = ""
                    objDelform.Items.Item("trefno").Specific.String = ""
                    Field_Settings(objDelform.UniqueID, objDelform.Items.Item("4").Specific.String)
                    'Matrix0.Columns.Item("tentry").Cells.Item(Row).Specific.String = DocEntry
                    objaddon.objapplication.StatusBar.SetText("GRPO Transaction Created Successfully...Document Number->" & TranEntry & " Items Count- " & DocEntry, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                    m_oProgBar.Text = "GRPO Transaction Created Successfully... Document Number->" & TranEntry & " Items Count- " & DocEntry
                    objaddon.objapplication.MessageBox("GRPO Transaction Created Successfully... Document Number->" & TranEntry & " Items Count- " & DocEntry, , "OK")
                End If
                System.Runtime.InteropServices.Marshal.ReleaseComObject(objGRPO)
                GC.Collect()
                If TFlag = True Then
                    objaddon.objapplication.StatusBar.SetText("Error Occurred while creating the transaction...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Return False
                Else
                    objaddon.objapplication.StatusBar.SetText("GRPO Transaction Created Successfully...Document Number->" & DocEntry, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                    Return True
                End If

            Catch ex As Exception
                objaddon.objapplication.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                'objaddon.objapplication.MessageBox(ex.Message, , "OK")
                Return False
                'If objAddOn.objCompany.InTransaction Then objAddOn.objCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
            Finally
                m_oProgBar.Stop()
                System.Runtime.InteropServices.Marshal.ReleaseComObject(m_oProgBar)
                m_oProgBar = Nothing
            End Try

        End Function

    End Class
End Namespace
