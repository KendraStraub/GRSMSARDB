
Namespace LightSwitchApplication

    Public Class ManageDivisions


        Private Sub ImportFromExcel_Execute()

            OfficeIntegration.Excel.Import(Me.ParkDivisions)

        End Sub
    End Class

End Namespace