
Namespace LightSwitchApplication

    Public Class xref_EmployeeZones

        Private Sub Summary_Compute(ByRef result As String)
            If (ResponseZones IsNot Nothing) Then
                result = ResponseZones.ResponseZone
            End If
        End Sub
    End Class

End Namespace
