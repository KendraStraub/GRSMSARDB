﻿
Namespace LightSwitchApplication

    Public Class ManageDivisions


        Private Sub ExportToExcel_Execute()

            OfficeIntegration.Excel.Export(Me.ParkDivisions)

        End Sub
    End Class

End Namespace