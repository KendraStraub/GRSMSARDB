
Namespace LightSwitchApplication

    Public Class Employee
        Private Sub EmployeeItem_Created()
            CreatedTime = DateTime.Now
            UpdatedTime = DateTime.Now
            UpdatedBy = Application.User.Name
            CreatedBy = Application.User.Name
        End Sub
        Private Sub DaysSinceLastFit_Compute(ByRef result As Integer)
            'If DateFit is null at form initialize, need to handle the null
            'else exception is thrown
            If Me.DateFit Is Nothing Then
                result = Nothing
            Else
                'If a DateFit is entered then calculate days since fit test
                Dim daysSince = Date.Now.Subtract(Me.DateFit).Days
                result = daysSince
            End If
        End Sub

        Private Sub PersonalSMS_Compute(ByRef result As String)
            'If PersonalSMSCarriers is null at form initialize, need to handle the null
            'else exception is thrown
            If Me.PersonalSMSCarriers Is Nothing Then
            Else
                'If a PersonalSMSCarrier is entered then concatonate cell phone number with carrier sms gateway from picklist
                result = PersonalMobile + PersonalSMSCarriers.SMSGateway
            End If
        End Sub

        Private Sub WorksSMS_Compute(ByRef result As String)
            'If PersonalSMSCarriers is null at form initialize, need to handle the null
            'else exception is thrown
            If Me.WorkSMSCarriers Is Nothing Then
            Else
                'If a WorkSMScarrier is entered then concatonate cell phone number with carrier sms gateway from picklist
                result = WorkMobile + WorkSMSCarriers.SMSGateway
            End If
        End Sub

        Private Sub Summary_Compute(ByRef result As String)
            'concatonate first and last name for summary field, used for lists
            result = LastName & ", " & FirstName
        End Sub

        Private Sub Employee_Created()
            Me.CLEO = "No"
            Me.MEDIC = "No"
            Me.Tracker = "No"
            Me.TechRescue = "No"

  
        End Sub


    End Class

End Namespace
