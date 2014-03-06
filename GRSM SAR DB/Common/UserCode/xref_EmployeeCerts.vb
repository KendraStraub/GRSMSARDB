
Namespace LightSwitchApplication

    Public Class xref_EmployeeCerts

        Private Sub Summary_Compute(ByRef result As String)
            If (Certifications IsNot Nothing) Then
                result = Certifications.Certification
            End If
        End Sub
    End Class

End Namespace
