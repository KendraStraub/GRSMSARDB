
Namespace LightSwitchApplication

    Public Class xref_EmployeeTechCert

        Private Sub Summary_Compute(ByRef result As String)
            If (TechRescueCerts IsNot Nothing) Then
                result = TechRescueCerts.Certification
            End If
        End Sub
    End Class

End Namespace
