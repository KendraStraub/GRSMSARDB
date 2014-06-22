
Namespace LightSwitchApplication

    Public Class Tracker_IRT

        Private Sub Employee_Loaded(succeeded As Boolean)
            ' Write your code here.
            Me.SetDisplayNameFromEntity(Me.Employee)
        End Sub

        Private Sub Employee_Changed()
            ' Write your code here.
            Me.SetDisplayNameFromEntity(Me.Employee)
        End Sub

        Private Sub Tracker_IRT_Saved()
            ' Write your code here.
            Me.SetDisplayNameFromEntity(Me.Employee)
        End Sub

    End Class

End Namespace