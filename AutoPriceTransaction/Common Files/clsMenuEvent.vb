Imports SAPbouiCOM
Namespace AutoPriceTransaction

    Public Class clsMenuEvent
        Dim objform As SAPbouiCOM.Form
        Dim objglobalmethods As New clsGlobalMethods

        Public Sub MenuEvent_For_StandardMenu(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
            Try
                Select Case objaddon.objapplication.Forms.ActiveForm.TypeEx
                    Case "940", "721", "720", "133", "143", "179", "180", "182"
                        Default_Sample_MenuEvent(pVal, BubbleEvent)
                End Select
            Catch ex As Exception

            End Try
        End Sub

        Private Sub Default_Sample_MenuEvent(ByVal pval As SAPbouiCOM.MenuEvent, ByVal BubbleEvent As Boolean)
            Try
                objform = objaddon.objapplication.Forms.ActiveForm
                Dim oUDFForm As SAPbouiCOM.Form
                If pval.BeforeAction = True Then
                Else
                    Select Case pval.MenuUID
                        Case "1287", "1281"
                            oUDFForm = objaddon.objapplication.Forms.Item(objform.UDFFormUID)
                            If oUDFForm.Items.Item("U_RefNo").Enabled = False Then
                                oUDFForm.Items.Item("U_RefNo").Enabled = True
                            End If
                            If objform.Items.Item("1").Enabled = False Then
                                objform.Items.Item("1").Enabled = True
                            End If
                            oUDFForm.Items.Item("U_RefNo").Specific.String = ""
                        Case Else
                            oUDFForm = objaddon.objapplication.Forms.Item(objform.UDFFormUID)
                            If oUDFForm.Items.Item("U_RefNo").Enabled = True Then
                                oUDFForm.Items.Item("U_RefNo").Enabled = False
                            End If
                    End Select
                End If
            Catch ex As Exception
                objaddon.objapplication.SetStatusBarMessage("Error in Standart Menu Event" + ex.ToString, SAPbouiCOM.BoMessageTime.bmt_Medium, True)
            End Try
        End Sub

        Private Sub ProductionOrder_MenuEvent(ByRef pval As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
            Try
                objform = objaddon.objapplication.Forms.ActiveForm

                If pval.BeforeAction = True Then
                    Select Case pval.MenuUID
                        Case "1283", "1284" 'Remove & Cancel
                            'objaddon.objapplication.SetStatusBarMessage("Remove or Cancel is not allowed ", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                            'BubbleEvent = False
                    End Select
                Else
                    Select Case pval.MenuUID
                        Case "1281" 'Find Mode 

                        Case "1282" ' Add Mode
                        Case "1288", "1289", "1290", "1291", "1304"

                        Case "1293"
                        Case "1292"
                    End Select
                End If
            Catch ex As Exception
                objform.Freeze(False)
                objaddon.objapplication.SetStatusBarMessage("Error in Menu Event" + ex.ToString, SAPbouiCOM.BoMessageTime.bmt_Medium, True)
            End Try
        End Sub

    End Class
End Namespace