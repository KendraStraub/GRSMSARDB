
Namespace LightSwitchApplication

    Public Class Employee

        Private Sub DaysSinceLastFit_Compute(ByRef result As String)
            'Set result to the desired field value
            Dim daysSince = Date.Now.Subtract(Me.DateFit).Days
            result = daysSince & " Days Since Last Fitness Test"
        End Sub

        Private Sub PersonalSMS_Compute(ByRef result As String)
            'concatonate cell phone number with carrier sms gateway from pickkist
            result = PersonalMobile + PersonalSMSCarriers.SMSGateway
        End Sub

        Private Sub WorksSMS_Compute(ByRef result As String)
            'concatonate cell phone number with carrier sms gateway from pickkist
            result = WorkMobile + WorkSMSCarriers.SMSGateway
        End Sub
    End Class

End Namespace
