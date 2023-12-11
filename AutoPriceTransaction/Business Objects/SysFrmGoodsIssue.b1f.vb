Option Strict Off
Option Explicit On

Imports SAPbouiCOM.Framework

Namespace AutoPriceTransaction
    <FormAttribute("720", "Business Objects/SysFrmGoodsIssue.b1f")>
    Friend Class SysFrmGoodsIssue
        Inherits SystemFormBase
        Public WithEvents objform As SAPbouiCOM.Form
        Private WithEvents objCombo As SAPbouiCOM.ComboBox
        Public Sub New()
        End Sub

        Public Overrides Sub OnInitializeComponent()
            Me.ComboBox0 = CType(Me.GetItem("2310000078").Specific, SAPbouiCOM.ComboBox)
            Me.OnCustomInitialize()

        End Sub

        Public Overrides Sub OnInitializeFormEvents()

        End Sub

        Private WithEvents ComboBox0 As SAPbouiCOM.ComboBox

        Private Sub OnCustomInitialize()
            Try
                objform = objaddon.objapplication.Forms.GetForm("720", 0)
            Catch ex As Exception

            End Try
        End Sub

        Private Sub ComboBox0_ComboSelectAfter(sboObject As Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles ComboBox0.ComboSelectAfter
            'Try
            '    Dim DocEntry As String = ""
            '    TransactionEntry = ""
            '    DocEntry = objform.DataSources.DBDataSources.Item("OIGE").GetValue("DocEntry", 0)
            '    TransactionEntry = DocEntry
            'Catch ex As Exception

            'End Try

        End Sub

        Private Sub ComboBox0_ComboSelectBefore(sboObject As Object, pVal As SAPbouiCOM.SBOItemEventArg, ByRef BubbleEvent As Boolean) Handles ComboBox0.ComboSelectBefore
            Try
                objaddon.objapplication.StatusBar.SetText("Feature disabled in add-on...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                BubbleEvent = False
            Catch ex As Exception

            End Try

        End Sub
    End Class
End Namespace
