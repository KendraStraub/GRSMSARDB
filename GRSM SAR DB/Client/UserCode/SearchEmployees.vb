
Namespace LightSwitchApplication

    Public Class SearchEmployees


        Private Sub ExportToExcel_Execute()
            OfficeIntegration.Excel.Export(Me.IRT)
        End Sub
    End Class

End Namespace
