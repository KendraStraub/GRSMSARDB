
Namespace LightSwitchApplication

    Public Class Employee
        Private Sub EmployeeItem_Created()
            CreatedTime = DateTime.Now
            UpdatedTime = DateTime.Now
            UpdatedBy = Application.User.Name
            CreatedBy = Application.User.Name
        End Sub
        Private Sub DaysSinceLastFit_Compute(ByRef result As String)
            'Set result to the desired field value
            Dim daysSince = Date.Now.Subtract(Me.DateFit).Days
            If Me.DateFit Is Nothing Then
                result = "No Fitness Test Date"
            Else
                result = daysSince & " Days Since Last Fitness Test"
            End If
        End Sub

        Private Sub PersonalSMS_Compute(ByRef result As String)
            'concatonate cell phone number with carrier sms gateway from pickkist
            result = PersonalMobile + PersonalSMSCarriers.SMSGateway
        End Sub

        Private Sub WorksSMS_Compute(ByRef result As String)
            'concatonate cell phone number with carrier sms gateway from pickkist
            result = WorkMobile + WorkSMSCarriers.SMSGateway
        End Sub

        Private Sub Summary_Compute(ByRef result As String)
            'concatonate first and last name for summary field, used for lists
            result = LastName & ", " & FirstName
        End Sub
    End Class

End Namespace
