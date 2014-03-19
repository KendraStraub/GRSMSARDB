
Namespace LightSwitchApplication

    Public Class ManageSARCertifications

        Private Sub ImportFromExcel_Execute()
            OfficeIntegration.Excel.Import(Me.CertificationsSet)

        End Sub
    End Class

End Namespace
