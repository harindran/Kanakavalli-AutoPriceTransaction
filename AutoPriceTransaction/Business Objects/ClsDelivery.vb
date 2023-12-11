Imports System.IO
Imports SAPbouiCOM.Framework
Namespace AutoPriceTransaction
    Public Class ClsDelivery
        Public Const Formtype = "140"
        Dim objDelform, objUDFForm As SAPbouiCOM.Form
        'Dim ObjQCForm As SAPbouiCOM.Form
        'Dim ObjGRPOForm As SAPbouiCOM.Form
        Dim objmatrix As SAPbouiCOM.Matrix
        Dim strSQL As String
        Dim GetValidCode As Boolean

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
                                GetValidCode = Find_GrpCmpny_Customer(FormUID, objDelform.Items.Item("4").Specific.String)
                                If GetValidCode Then GetValidGP = True Else GetValidGP = False
                            End If
                            'If pVal.ItemUID = "38" And (pVal.ColUID = "1" Or pVal.ColUID = "29") Then
                            '    objUDFForm = objaddon.objapplication.Forms.Item(objDelform.UDFFormUID)
                            '    If objUDFForm.Items.Item("U_TRANSTYPE").Specific.Selected.Value = "STOCK TRANSFER" Then
                            '        If objaddon.HANA Then
                            '            strSQL = "Select ""COGM_Act"" from OGAR Where ""PeriodCat""=(Select ""Category"" from OFPR Where CURRENT_DATE between ""F_RefDate"" and ""T_RefDate"") and ""UDF2""= 'STOCK TRANSFER'"
                            '        End If
                            '        strSQL = objaddon.objglobalmethods.getSingleValue(strSQL)
                            '        Dim AcctCode As String = objmatrix.Columns.Item("29").Cells.Item(pVal.Row).Specific.string
                            '        If AcctCode <> strSQL And AcctCode <> "" Then
                            '            objaddon.objapplication.StatusBar.SetText("G/L Account is not valid for the transaction type on Line: " & pVal.Row, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            '            BubbleEvent = False : Exit Sub
                            '        End If

                            '    End If
                            'End If
                        Case SAPbouiCOM.BoEventTypes.et_CLICK
                            If pVal.ItemUID = "btnsave" Then
                                Dim objcombo As SAPbouiCOM.ComboBox
                                objcombo = objDelform.Items.Item("2001").Specific
                                If objcombo.Value = "" Then
                                    objaddon.objapplication.StatusBar.SetText("Branch is Missing...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                    BubbleEvent = False : Exit Sub
                                End If
                                objUDFForm = objaddon.objapplication.Forms.Item(objDelform.UDFFormUID)
                                If objUDFForm.Items.Item("U_TRANSTYPE").Specific.Selected.Value = "STOCK TRANSFER" Then
                                    If objaddon.HANA Then
                                        strSQL = "Select ""COGM_Act"" from OGAR Where ""PeriodCat""=(Select ""Category"" from OFPR Where CURRENT_DATE between ""F_RefDate"" and ""T_RefDate"") and ""UDF2""= 'STOCK TRANSFER'"
                                    End If
                                    strSQL = objaddon.objglobalmethods.getSingleValue(strSQL)
                                    For i = 1 To objmatrix.VisualRowCount
                                        Dim AcctCode As String = objmatrix.Columns.Item("29").Cells.Item(i).Specific.string
                                        If AcctCode <> strSQL And AcctCode <> "" Then
                                            objaddon.objapplication.StatusBar.SetText("G/L Account is not valid for the transaction type on Line: " & i, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                            BubbleEvent = False : Exit Sub
                                        End If
                                    Next
                                End If
                            End If
                        Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN
                            If pVal.ItemUID = "38" And (pVal.ColUID = "14" Or pVal.ColUID = "U_MRP") And pVal.CharPressed <> 9 Then
                                BubbleEvent = False
                            End If
                        Case SAPbouiCOM.BoEventTypes.et_FORM_CLOSE
                            GetValidGP = False
                    End Select
                Else
                    Select Case pVal.EventType
                        Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                            If pVal.ActionSuccess Then
                                CreateButton(FormUID)
                            End If

                        Case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS
                            Dim GetValue As String = ""
                            If GetValidCode = True Then Exit Sub
                            If pVal.ItemUID = "38" And pVal.ColUID = "1" Then
                                If Trim(objmatrix.Columns.Item("1").Cells.Item(pVal.Row).Specific.string) = "" Then Exit Sub
                                GetValue = objaddon.objglobalmethods.getSingleValue("SELECT distinct ""Price"" FROM ITM1 Where ""PriceList""=2 and ""ItemCode""='" & Trim(objmatrix.Columns.Item("1").Cells.Item(pVal.Row).Specific.string) & "'")
                                If GetValue <> "" Then
                                    If objmatrix.Columns.Item("U_MRP").Visible = True Then
                                        objmatrix.Columns.Item("U_MRP").Cells.Item(pVal.Row).Specific.string = GetValue
                                    End If
                                End If
                                If Trim(objmatrix.Columns.Item("160").Cells.Item(pVal.Row).Specific.string) = "" Then Exit Sub
                                GetValue = objaddon.objglobalmethods.getSingleValue("SELECT (SELECT (100/(100+""Rate"")) from OSTC where ""Code""=(select '" & Trim(objmatrix.Columns.Item("160").Cells.Item(pVal.Row).Specific.string) & "' from dummy)) * (SELECT DISTINCT ""Price"" FROM ITM1 Where ""PriceList""=2 and ""ItemCode""='" & Trim(objmatrix.Columns.Item("1").Cells.Item(pVal.Row).Specific.string) & "') as ""UnitPrice"" FROM DUMMY")
                                If GetValue <> "" Then
                                    If objmatrix.Columns.Item("14").Visible = True Then
                                        objmatrix.Columns.Item("14").Cells.Item(pVal.Row).Specific.string = GetValue
                                    End If
                                End If
                            ElseIf pVal.ItemUID = "38" And pVal.ColUID = "160" Then
                                If GetValidCode = True Then Exit Sub
                                If Trim(objmatrix.Columns.Item("160").Cells.Item(pVal.Row).Specific.string) = "" Then Exit Sub
                                'GetValue = objaddon.objglobalmethods.getSingleValue("select ""Rate"" from OSTC where ""Code""='" & Trim(objmatrix.Columns.Item("160").Cells.Item(pVal.Row).Specific.string) & "'")
                                GetValue = objaddon.objglobalmethods.getSingleValue("SELECT (SELECT (100/(100+""Rate"")) from OSTC where ""Code""=(select '" & Trim(objmatrix.Columns.Item("160").Cells.Item(pVal.Row).Specific.string) & "' from dummy)) * (SELECT DISTINCT ""Price"" FROM ITM1 Where ""PriceList""=2 and ""ItemCode""='" & Trim(objmatrix.Columns.Item("1").Cells.Item(pVal.Row).Specific.string) & "') as ""UnitPrice"" FROM DUMMY")
                                If GetValue <> "" Then
                                    If objmatrix.Columns.Item("14").Visible = True Then
                                        objmatrix.Columns.Item("14").Cells.Item(pVal.Row).Specific.string = GetValue
                                    End If
                                End If
                            End If
                        Case SAPbouiCOM.BoEventTypes.et_CLICK

                            If pVal.ItemUID = "btnsave" And objDelform.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                                If objDelform.Items.Item("btnsave").Enabled = False Then Exit Sub
                                'If objform.Items.Item("tgirefno").Specific.String = "" Then Exit Sub
                                Dim objcomb As SAPbouiCOM.ComboBox
                                objcomb = objDelform.Items.Item("3").Specific
                                If objcomb.Selected.Value = "S" Then
                                    objaddon.objapplication.StatusBar.SetText("Service Type found.Please Check...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                    Exit Sub
                                End If
                                If Trim(objmatrix.Columns.Item("1").Cells.Item(1).Specific.string) = "" Then Exit Sub
                                If objaddon.objapplication.MessageBox("Delivery Transaction cannot be reversed. Do you want to continue?", 2, "Yes", "No") <> 1 Then Exit Sub
                                If Create_Delivery(FormUID) Then
                                    'objform.Items.Item("2").Click()
                                End If
                            End If
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
                        Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD

                    End Select
                Else
                    Select Case BusinessObjectInfo.EventType
                        Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD, SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE

                        Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD
                            Try
                                Field_Settings(objDelform.UniqueID, objDelform.Items.Item("4").Specific.String)
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
                'objItem.Left = objForm.Items.Item("10002056").Left + objForm.Items.Item("10002056").Width + 60
                objItem.Width = 65
                objItem.Top = objDelform.Items.Item("2").Top
                objItem.Height = objDelform.Items.Item("2").Height
                objButton = objItem.Specific
                objButton.Caption = "Save"
                objaddon.objglobalmethods.SetAutomanagedattribute_Editable(objDelform, "btnsave", True, False, False)

                'Dim objedit As SAPbouiCOM.EditText
                'objItem = objDelform.Items.Add("txtGet", SAPbouiCOM.BoFormItemTypes.it_EDIT)
                'objItem.Left = objDelform.Items.Item("BtnVal").Left + objDelform.Items.Item("BtnVal").Width + 10
                'objItem.Width = 50
                'objItem.Top = objDelform.Items.Item("BtnVal").Top
                'objItem.Height = objDelform.Items.Item("BtnVal").Height
                'objItem.LinkTo = "BtnVal"
                'objedit = objItem.Specific
                'objedit.Item.Enabled = False
                'objedit.DataBind.SetBound(True, "OPCH", "U_GetNum")
                'objAddOn.objApplication.SetStatusBarMessage("Button Created", SAPbouiCOM.BoMessageTime.bmt_Short, False)
            Catch ex As Exception
            End Try

        End Sub

        Private Sub Field_Settings(ByVal FormUID As String, ByVal CardCode As String)
            Try
                'If CardCode = "" Then Exit Sub
                objDelform = objaddon.objapplication.Forms.Item(FormUID)
                CreateButton(FormUID)
                strSQL = "Select 1 as ""Status"" from OCRD T0 left join OCRG T1 on T0.""GroupCode""=T1.""GroupCode"""
                strSQL += vbCrLf + "where T1.""GroupType""='C' and T1.""GroupName"" like 'Group%' and T0.""CardCode""='" & Trim(CardCode) & "'"

                strSQL = objaddon.objglobalmethods.getSingleValue(strSQL)
                objmatrix = objDelform.Items.Item("38").Specific
                If strSQL = "1" Then
                    If objDelform.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then objDelform.Items.Item("btnsave").Enabled = True Else objDelform.Items.Item("btnsave").Enabled = False
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
                objaddon.objapplication.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, True)
            End Try
        End Sub

        Private Function Create_Delivery(ByVal FormUID As String) As Boolean
            Dim DocEntry, BranchEnabled, Series As String
            Dim objDelivery As SAPbobsCOM.Documents
            Dim objEdit As SAPbouiCOM.EditText
            'Dim MBAPDocNum As Long
            Dim TranLine As Integer = 0
            Dim TFlag As Boolean = False
            Dim m_oProgBar As SAPbouiCOM.ProgressBar
            Dim objrs As SAPbobsCOM.Recordset
            objrs = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Try
                objDelform = objaddon.objapplication.Forms.Item(FormUID)
                objmatrix = objDelform.Items.Item("38").Specific
                If objaddon.HANA Then
                    BranchEnabled = objaddon.objglobalmethods.getSingleValue("select ""MltpBrnchs"" from OADM")
                Else
                    BranchEnabled = objaddon.objglobalmethods.getSingleValue("select MltpBrnchs from OADM")
                End If
                If Not BranchEnabled = "Y" Then objaddon.objapplication.StatusBar.SetText("Branch not enabled...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning) : Exit Function
                objaddon.objapplication.StatusBar.SetText("Delivery Transaction Creating Please wait...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                objDelivery = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDeliveryNotes) 'oDeliveryNotes

                If Not objaddon.objcompany.InTransaction Then objaddon.objcompany.StartTransaction()
                m_oProgBar = objaddon.objapplication.StatusBar.CreateProgressBar("My Progress", objmatrix.VisualRowCount, True)
                m_oProgBar.Text = "Delivery Transaction Creating Please wait... "
                m_oProgBar.Value = 0
                'Dim DocDate As Date = Date.ParseExact(EditText1.Value, "yyyyMMdd", System.Globalization.DateTimeFormatInfo.InvariantInfo)
                objEdit = objDelform.Items.Item("10").Specific
                Dim DocDate As Date = Date.ParseExact(objEdit.Value, "yyyyMMdd", System.Globalization.DateTimeFormatInfo.InvariantInfo)
                objEdit = objDelform.Items.Item("12").Specific
                Dim DueDate As Date = Date.ParseExact(objEdit.Value, "yyyyMMdd", System.Globalization.DateTimeFormatInfo.InvariantInfo)
                objEdit = objDelform.Items.Item("46").Specific
                Dim TaxDate As Date = Date.ParseExact(objEdit.Value, "yyyyMMdd", System.Globalization.DateTimeFormatInfo.InvariantInfo)
                'MBAPDocNum = objform.BusinessObject.GetNextSerialNumber(objform.Items.Item("Series").Specific.Selected.value)
                objDelivery.CardCode = objDelform.Items.Item("4").Specific.String
                objDelivery.DocDate = DocDate
                objDelivery.DocDueDate = DueDate
                objDelivery.TaxDate = TaxDate
                objDelivery.NumAtCard = objDelform.Items.Item("14").Specific.String
                objDelivery.JournalMemo = "Auto-Generated->  " & Now.ToString
                objDelivery.Comments = "Created by " & objaddon.objcompany.UserName & " on " & Now.ToString & "-" & objDelform.Items.Item("16").Specific.String
                If objaddon.HANA Then
                    Series = objaddon.objglobalmethods.getSingleValue("SELECT Top 1 ""Series"" FROM NNM1 WHERE ""ObjectCode""='15' and ""Indicator""=(Select ""Indicator"" from OFPR where CURRENT_DATE Between ""F_RefDate"" and ""T_RefDate"") " &
                                                                                   " and ""BPLId""='" & objDelform.Items.Item("2001").Specific.Selected.Value & "'")
                Else
                    Series = objaddon.objglobalmethods.getSingleValue("SELECT Top 1 Series FROM NNM1 WHERE ObjectCode='15' and Indicator=(Select Indicator from OFPR where CURRENT_DATE Between F_RefDate and T_RefDate) " &
                                                                                  " and BPLId='" & objDelform.Items.Item("2001").Specific.Selected.Value & "'")
                End If

                If Series = "" Then objaddon.objapplication.StatusBar.SetText("Numbering Series not found.Please check...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning) : Exit Function
                objDelivery.Series = Series
                'objPurchaseInvoice.Series = 110
                Dim oUDFForm As SAPbouiCOM.Form
                oUDFForm = objaddon.objapplication.Forms.Item(objDelform.UDFFormUID)
                objDelivery.UserFields.Fields.Item("U_FROMWHS").Value = oUDFForm.Items.Item("U_FROMWHS").Specific.String
                objDelivery.UserFields.Fields.Item("U_TOWHS").Value = oUDFForm.Items.Item("U_TOWHS").Specific.String
                objDelivery.UserFields.Fields.Item("U_TRANSTYPE").Value = oUDFForm.Items.Item("U_TRANSTYPE").Specific.Selected.Value
                'objDelivery.UserFields.Fields.Item("U_RefNo").Value = DocEntry
                If BranchEnabled = "Y" Then
                    objDelivery.BPL_IDAssignedToInvoice = objDelform.Items.Item("2001").Specific.Selected.Value
                End If
                objDelivery.DocType = SAPbobsCOM.BoDocumentTypes.dDocument_Items
                objDelivery.PayToCode = objDelform.Items.Item("226").Specific.Selected.Value 'Bill To
                'objDelivery.Address = objDelform.Items.Item("6").Specific.String 'Bill To
                objrs.DoQuery("select  ""Address"",""CardCode"",""Street"",""Block"",""ZipCode"",""City"",""County"",""Country"",""State"",""Building"",""AdresType"",""Address2"",""Address3"",""AddrType"",""StreetNo""  from CRD1 where ""CardCode""='" & objDelform.Items.Item("4").Specific.String & "' and ""Address""='" & objDelform.Items.Item("226").Specific.Selected.Value & "' and ""AdresType""='B'")

                objDelivery.AddressExtension.BillToAddress2 = Trim(objrs.Fields.Item("Address2").Value.ToString)
                objDelivery.AddressExtension.BillToAddress3 = Trim(objrs.Fields.Item("Address3").Value.ToString)
                objDelivery.AddressExtension.BillToAddressType = Trim(objrs.Fields.Item("AddrType").Value.ToString)
                objDelivery.AddressExtension.BillToBlock = Trim(objrs.Fields.Item("Block").Value.ToString)
                objDelivery.AddressExtension.BillToBuilding = Trim(objrs.Fields.Item("Building").Value.ToString)
                objDelivery.AddressExtension.BillToCity = Trim(objrs.Fields.Item("City").Value.ToString)
                objDelivery.AddressExtension.BillToCountry = Trim(objrs.Fields.Item("Country").Value.ToString)
                objDelivery.AddressExtension.BillToCounty = Trim(objrs.Fields.Item("County").Value.ToString)
                objDelivery.AddressExtension.BillToState = Trim(objrs.Fields.Item("State").Value.ToString)
                objDelivery.AddressExtension.BillToStreet = Trim(objrs.Fields.Item("Street").Value.ToString)
                objDelivery.AddressExtension.BillToStreetNo = Trim(objrs.Fields.Item("StreetNo").Value.ToString)
                objDelivery.AddressExtension.BillToZipCode = Trim(objrs.Fields.Item("ZipCode").Value.ToString)

                objDelivery.ShipToCode = objDelform.Items.Item("40").Specific.Selected.Value 'Ship To
                'objDelivery.Address2 = objDelform.Items.Item("92").Specific.String 'Ship To
                objrs.DoQuery("select  ""Address"",""CardCode"",""Street"",""GSTRegnNo"",""Block"",""ZipCode"",""City"",""County"",""Country"",""State"",""Building"",""AdresType"",""Address2"",""Address3"",""AddrType"",""StreetNo""  from CRD1 where ""CardCode""='" & objDelform.Items.Item("4").Specific.String & "' and ""Address""='" & objDelform.Items.Item("226").Specific.Selected.Value & "' and ""AdresType""='S'")

                objDelivery.AddressExtension.ShipToAddress2 = Trim(objrs.Fields.Item("Address2").Value.ToString)
                objDelivery.AddressExtension.ShipToAddress3 = Trim(objrs.Fields.Item("Address3").Value.ToString)
                objDelivery.AddressExtension.ShipToAddressType = Trim(objrs.Fields.Item("AddrType").Value.ToString)
                objDelivery.AddressExtension.ShipToBlock = Trim(objrs.Fields.Item("Block").Value.ToString)
                objDelivery.AddressExtension.ShipToBuilding = Trim(objrs.Fields.Item("Building").Value.ToString)
                objDelivery.AddressExtension.ShipToCity = Trim(objrs.Fields.Item("City").Value.ToString)
                objDelivery.AddressExtension.ShipToCountry = Trim(objrs.Fields.Item("Country").Value.ToString)
                objDelivery.AddressExtension.ShipToCounty = Trim(objrs.Fields.Item("County").Value.ToString)
                objDelivery.AddressExtension.ShipToState = Trim(objrs.Fields.Item("State").Value.ToString)
                objDelivery.AddressExtension.ShipToStreet = Trim(objrs.Fields.Item("Street").Value.ToString)
                objDelivery.AddressExtension.ShipToStreetNo = Trim(objrs.Fields.Item("StreetNo").Value.ToString)
                objDelivery.AddressExtension.ShipToZipCode = Trim(objrs.Fields.Item("ZipCode").Value.ToString)


                For Row As Integer = 1 To objmatrix.VisualRowCount
                    If objmatrix.Columns.Item("1").Cells.Item(Row).Specific.string <> "" Then
                        objDelivery.Lines.ItemCode = Trim(objmatrix.Columns.Item("1").Cells.Item(Row).Specific.string)
                        'objDelivery.Lines.ItemDescription = Trim(Matrix0.Columns.Item("3").Cells.Item(Row).Specific.string)
                        objDelivery.Lines.Quantity = Trim(CDbl(objmatrix.Columns.Item("11").Cells.Item(Row).Specific.string))
                        objDelivery.Lines.AccountCode = Trim(objmatrix.Columns.Item("29").Cells.Item(Row).Specific.string)
                        'If Matrix0.Columns.Item("cc1").Cells.Item(Row).Specific.string <> "" Then objDelivery.Lines.CostingCode = Matrix0.Columns.Item("cc1").Cells.Item(Row).Specific.string
                        'If Matrix0.Columns.Item("cc2").Cells.Item(Row).Specific.string <> "" Then objDelivery.Lines.CostingCode2 = Matrix0.Columns.Item("cc2").Cells.Item(Row).Specific.string
                        'If Matrix0.Columns.Item("cc3").Cells.Item(Row).Specific.string <> "" Then objDelivery.Lines.CostingCode3 = Matrix0.Columns.Item("cc3").Cells.Item(Row).Specific.string
                        'If Matrix0.Columns.Item("cc4").Cells.Item(Row).Specific.string <> "" Then objDelivery.Lines.CostingCode4 = Matrix0.Columns.Item("cc4").Cells.Item(Row).Specific.string
                        'If Matrix0.Columns.Item("cc5").Cells.Item(Row).Specific.string <> "" Then objDelivery.Lines.CostingCode5 = Matrix0.Columns.Item("cc5").Cells.Item(Row).Specific.string
                        'If Matrix0.Columns.Item("lrefno").Cells.Item(Row).Specific.string <> "" Then objDelivery.Lines.UserFields.Fields.Item("U_ymhbpref").Value = Matrix0.Columns.Item("lrefno").Cells.Item(Row).Specific.string
                        objDelivery.Lines.TaxCode = Trim(objmatrix.Columns.Item("160").Cells.Item(Row).Specific.string)
                        If objaddon.HANA Then
                            objDelivery.Lines.UnitPrice = objaddon.objglobalmethods.getSingleValue("Select ""LastPurPrc"" from OITM where ""ItemCode""='" & Trim(objmatrix.Columns.Item("1").Cells.Item(Row).Specific.string) & "'")
                        Else
                            objDelivery.Lines.UnitPrice = objaddon.objglobalmethods.getSingleValue("Select LastPurPrc from OITM where ItemCode='" & Trim(objmatrix.Columns.Item("1").Cells.Item(Row).Specific.string) & "'")
                        End If
                        'objDelivery.Lines.UnitPrice = Matrix0.Columns.Item("total").Cells.Item(Row).Specific.string
                        objDelivery.Lines.WarehouseCode = objmatrix.Columns.Item("24").Cells.Item(Row).Specific.string
                        objDelivery.Lines.ProjectCode = objmatrix.Columns.Item("31").Cells.Item(Row).Specific.string
                        'If objaddon.HANA Then
                        '    objDelivery.Lines.LocationCode = objaddon.objglobalmethods.getSingleValue("Select T1.""Location"" from OBPL T0 join OWHS T1 On T1.""BPLid""=T0.""BPLId"" And T0.""DflWhs""=T1.""WhsCode"" where T0.""BPLId""='" & Matrix0.Columns.Item("branch").Cells.Item(Row).Specific.Selected.Value & "'")
                        'Else
                        '    objDelivery.Lines.LocationCode = objaddon.objglobalmethods.getSingleValue("Select T1.Location from OBPL T0 join OWHS T1 On T1.BPLid=T0.BPLId And T0.DflWhs=T1.WhsCode where T0.BPLId='" & Matrix0.Columns.Item("branch").Cells.Item(Row).Specific.Selected.Value & "'")
                        'End If
                        Dim GetValue As String = objaddon.objglobalmethods.getSingleValue("SELECT distinct ""Price"" FROM ITM1 Where ""PriceList""=2 and ""ItemCode""='" & Trim(objmatrix.Columns.Item("1").Cells.Item(Row).Specific.string) & "'")
                        If GetValue <> "" Then
                            objDelivery.Lines.UserFields.Fields.Item("U_MRP").Value = GetValue 'objmatrix.Columns.Item("U_MRP").Cells.Item(Row).Specific.string
                        End If
                        'objDelivery.Lines.UserFields.Fields.Item("U_MRP").Value = objmatrix.Columns.Item("U_MRP").Cells.Item(Row).Specific.string
                        objDelivery.Lines.UserFields.Fields.Item("U_DocLine").Value = CStr(TranLine)
                        'objDelivery.Lines.LineTotal = Matrix0.Columns.Item("total").Cells.Item(Row).Specific.string
                        objDelivery.Lines.Add()
                        TranLine += 1
                        m_oProgBar.Value = Row
                    End If
                Next
                If objDelivery.Add() <> 0 Then
                    TFlag = True
                    If objaddon.objcompany.InTransaction Then objaddon.objcompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                    objaddon.objapplication.SetStatusBarMessage("Delivery: " & objaddon.objcompany.GetLastErrorDescription & "-" & objaddon.objcompany.GetLastErrorCode, SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                    objaddon.objapplication.MessageBox("Delivery: " & objaddon.objcompany.GetLastErrorDescription & "-" & objaddon.objcompany.GetLastErrorCode,, "OK")
                Else
                    'Dim sNewObjCode As String = ""
                    'objaddon.objcompany.GetNewObjectCode(sNewObjCode)
                    'Dim str = CLng(sNewObjCode)
                    If objaddon.objcompany.InTransaction Then objaddon.objcompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
                    DocEntry = objaddon.objcompany.GetNewObjectKey()
                    If objaddon.HANA Then
                        DocEntry = objaddon.objglobalmethods.getSingleValue("Select ""DocNum"" from ODLN where ""DocEntry""=" & DocEntry & "")
                    Else
                        DocEntry = objaddon.objglobalmethods.getSingleValue("Select DocNum from ODLN where DocEntry=" & DocEntry & "")
                    End If
                    objmatrix.Clear()
                    objmatrix.AddRow()
                    objDelform.Items.Item("4").Specific.String = ""
                    Field_Settings(objDelform.UniqueID, objDelform.Items.Item("4").Specific.String)
                    'Matrix0.Columns.Item("tentry").Cells.Item(Row).Specific.String = DocEntry
                    objaddon.objapplication.StatusBar.SetText("Delivery Transaction Created Successfully...Document Number->" & DocEntry, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                    m_oProgBar.Text = "Delivery Transaction Created Successfully... Document Number->" & DocEntry
                    objaddon.objapplication.MessageBox("Delivery Transaction Created Successfully... Document Number->" & DocEntry, , "OK")
                End If
                System.Runtime.InteropServices.Marshal.ReleaseComObject(objDelivery)
                GC.Collect()
                If TFlag = True Then
                    objaddon.objapplication.StatusBar.SetText("Error Occurred while creating the transaction...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Return False
                Else
                    objaddon.objapplication.StatusBar.SetText("Delivery Transaction Created Successfully...Document Number->" & DocEntry, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
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

        Private Function Find_GrpCmpny_Customer(ByVal FormUID As String, ByVal CardCode As String) As Boolean
            Try
                'If CardCode = "" Then Exit Sub
                objDelform = objaddon.objapplication.Forms.Item(FormUID)
                strSQL = "Select 1 as ""Status"" from OCRD T0 left join OCRG T1 on T0.""GroupCode""=T1.""GroupCode"""
                strSQL += vbCrLf + "where T1.""GroupType""='C' and T1.""GroupName"" like 'Group%' and T0.""CardCode""='" & Trim(CardCode) & "'"

                strSQL = objaddon.objglobalmethods.getSingleValue(strSQL)
                If strSQL = "1" Then
                    Return True
                Else
                    Return False
                End If
            Catch ex As Exception
            End Try
        End Function

    End Class
End Namespace
