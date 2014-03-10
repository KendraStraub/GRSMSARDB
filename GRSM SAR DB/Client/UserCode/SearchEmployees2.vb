
Namespace LightSwitchApplication

    Public Class SearchEmployees2

        Private Sub ExportToExcel_Execute()
            OfficeIntegration.Excel.Export(Me.Employees)
        End Sub
    End Class

End Namespace
