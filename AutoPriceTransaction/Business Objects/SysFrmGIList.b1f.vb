Option Strict Off
Option Explicit On

Imports SAPbouiCOM.Framework

Namespace AutoPriceTransaction
    <FormAttribute("10059", "Business Objects/SysFrmGIList.b1f")>
    Friend Class SysFrmGIList
        Inherits SystemFormBase
        Public WithEvents objform As SAPbouiCOM.Form
        Public Sub New()
        End Sub

        Public Overrides Sub OnInitializeComponent()
            Me.Button0 = CType(Me.GetItem("1").Specific, SAPbouiCOM.Button)
            Me.Matrix0 = CType(Me.GetItem("7").Specific, SAPbouiCOM.Matrix)
            Me.OnCustomInitialize()

        End Sub

        Public Overrides Sub OnInitializeFormEvents()

        End Sub

        Private Sub OnCustomInitialize()
            Try
                objform = objaddon.objapplication.Forms.GetForm("10059", 0)
                TransactionEntry = ""
            Catch ex As Exception

            End Try
        End Sub

        Private WithEvents Button0 As SAPbouiCOM.Button

        Private Sub Button0_ClickAfter(sboObject As Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles Button0.ClickAfter
            Try
                Dim GetDocEntry As String = ""
                Dim DocDate As Date
                Dim GetDate As SAPbouiCOM.EditText
                Dim GetValues As New List(Of String)
                For i As Integer = 1 To Matrix0.VisualRowCount
                    GetDate = Matrix0.Columns.Item("DocDate").Cells.Item(i).Specific
                    If Matrix0.IsRowSelected(i) Then
                        DocDate = Date.ParseExact(GetDate.Value, "yyyyMMdd", System.Globalization.CultureInfo.InvariantCulture)
                        GetDocEntry = objaddon.objglobalmethods.getSingleValue("Select ""DocEntry"" from OIGE where ""DocNum""= " & Matrix0.Columns.Item("DocNum").Cells.Item(i).Specific.String & " and ""DocDate""='" & DocDate.ToString("yyyyMMdd") & "' ")
                        'TransactionEntry = GetDocEntry
                        GetValues.Add(GetDocEntry)
                    End If
                Next
                If GetValues.Count > 0 Then
                    Dim DocEntryList = (From gv In GetValues Select New String(gv)).ToList()
                    TransactionEntry = String.Join(",", DocEntryList)
                End If
            Catch ex As Exception
            End Try

        End Sub

        Private WithEvents Matrix0 As SAPbouiCOM.Matrix

        Private Sub Matrix0_ClickAfter(sboObject As Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles Matrix0.ClickAfter
            Try
                Dim GetDocEntry As String = ""
                Dim DocDate As Date
                'Dim GetValues As New List(Of String)
                Dim GetDate As SAPbouiCOM.EditText
                If pVal.Modifiers = SAPbouiCOM.BoModifiersEnum.mt_None Then
                    GetDate = Matrix0.Columns.Item("DocDate").Cells.Item(Matrix0.GetNextSelectedRow).Specific
                    DocDate = Date.ParseExact(GetDate.Value, "yyyyMMdd", System.Globalization.CultureInfo.InvariantCulture)
                    GetDocEntry = objaddon.objglobalmethods.getSingleValue("Select ""DocEntry"" from OIGE where ""DocNum""= " & Matrix0.Columns.Item("DocNum").Cells.Item(Matrix0.GetNextSelectedRow).Specific.String & " and ""DocDate""='" & DocDate.ToString("yyyyMMdd") & "' ")
                    TransactionEntry = GetDocEntry
                    'Else
                    '    For i As Integer = 1 To Matrix0.VisualRowCount
                    '        GetDate = Matrix0.Columns.Item("DocDate").Cells.Item(i).Specific
                    '        If Matrix0.IsRowSelected(i) Then
                    '            DocDate = Date.ParseExact(GetDate.Value, "yyyyMMdd", System.Globalization.CultureInfo.InvariantCulture)
                    '            GetDocEntry = objaddon.objglobalmethods.getSingleValue("Select ""DocEntry"" from OIGE where ""DocNum""= " & Matrix0.Columns.Item("DocNum").Cells.Item(i).Specific.String & " and ""DocDate""='" & DocDate.ToString("yyyyMMdd") & "' ")
                    '            ' TransactionEntry = GetDocEntry
                    '            GetValues.Add(GetDocEntry)
                    '        End If
                    '    Next
                    '    If GetValues.Count > 0 Then
                    '        Dim DocEntryList = (From gv In GetValues Select New String(gv)).ToList()
                    '        TransactionEntry = String.Join(",", DocEntryList)
                    '    End If
                End If

            Catch ex As Exception
            End Try
        End Sub


    End Class
End Namespace
