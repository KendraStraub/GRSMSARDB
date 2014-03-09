
Namespace LightSwitchApplication

    Public Class SearchEmployees1

        Private Sub ImportFromExcel_Execute()
            OfficeIntegration.Excel.Import(Me.Employees)
        End Sub
    End Class

End Namespace
