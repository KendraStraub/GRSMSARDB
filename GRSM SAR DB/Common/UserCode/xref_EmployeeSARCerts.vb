
Namespace LightSwitchApplication

    Public Class xref_EmployeeSARCerts

        Private Sub Summary_Compute(ByRef result As String)
            If (SARCertificationsSetItem IsNot Nothing) Then
                result = SARCertificationsSetItem.Certification
            End If
        End Sub
    End Class

End Namespace
