
Namespace LightSwitchApplication

    Public Class xref_EmployeeWaterCert

        Private Sub Summary_Compute(ByRef result As String)
            If (WaterRescueCerts IsNot Nothing) Then
                result = WaterRescueCerts.SWRescueCert
            End If
        End Sub
    End Class

End Namespace
