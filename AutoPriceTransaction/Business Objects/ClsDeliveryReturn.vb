Namespace AutoPriceTransaction
    Public Class ClsDeliveryReturn
        Public Const Formtype = "180"
        Dim objDelform As SAPbouiCOM.Form
        Dim objmatrix As SAPbouiCOM.Matrix
        Dim strSQL As String

        Public Sub ItemEvent(FormUID As String, pVal As SAPbouiCOM.ItemEvent, BubbleEvent As Boolean)
            objDelform = objaddon.objapplication.Forms.Item(FormUID)
            If pVal.BeforeAction Then
                Select Case pVal.EventType
                    Case SAPbouiCOM.BoEventTypes.et_VALIDATE
                        If pVal.ItemUID = "4" Then
                            If objDelform.Items.Item("4").Specific.String = "" Then Exit Sub
                            If objDelform.Mode <> SAPbouiCOM.BoFormMode.fm_ADD_MODE Then Exit Sub
                            Field_Settings(FormUID, objDelform.Items.Item("4").Specific.String)
                        End If

                End Select
            Else
                Try
                    Select Case pVal.EventType
                        Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD

                        Case SAPbouiCOM.BoEventTypes.et_CLICK

                    End Select
                Catch ex As Exception

                End Try
            End If

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

        Private Sub Field_Settings(ByVal FormUID As String, ByVal CardCode As String)
            Try
                'If CardCode = "" Then Exit Sub
                objDelform = objaddon.objapplication.Forms.Item(FormUID)
                strSQL = "Select 1 as ""Status"" from OCRD T0 left join OCRG T1 on T0.""GroupCode""=T1.""GroupCode"""
                strSQL += vbCrLf + "where T1.""GroupType""='C' and T1.""GroupName"" like 'Group%' and T0.""CardCode""='" & Trim(CardCode) & "'"

                strSQL = objaddon.objglobalmethods.getSingleValue(strSQL)
                objmatrix = objDelform.Items.Item("38").Specific
                If strSQL = "1" Then
                    objDelform.Items.Item("1").Enabled = False
                    objmatrix.Columns.Item("14").Visible = False 'Unit Price
                    objmatrix.Columns.Item("21").Visible = False  'Line Total
                    objmatrix.Columns.Item("259").Visible = False  'Item Cost
                Else
                    objDelform.Items.Item("1").Enabled = True
                    objmatrix.Columns.Item("14").Visible = True 'Unit Price
                    objmatrix.Columns.Item("21").Visible = True  'Line Total
                    objmatrix.Columns.Item("259").Visible = True  'Item Cost
                End If
            Catch ex As Exception

            End Try
        End Sub

    End Class
End Namespace

