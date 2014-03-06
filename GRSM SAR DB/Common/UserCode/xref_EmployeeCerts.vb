
Namespace LightSwitchApplication

    Public Class xref_EmployeeCerts

        Private Sub Summary_Compute(ByRef result As String)
            If (CertificationItem IsNot Nothing) Then
                result = CertificationItem.Certification
            End If
        End Sub
    End Class

End Namespace
