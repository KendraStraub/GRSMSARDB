
Namespace LightSwitchApplication

    Public Class EmpoyeeDetail

        Private Sub ImportFromExcel_Execute()
            OfficeIntegration.Excel.Import(Me.Employees)
        End Sub
    End Class

End Namespace
