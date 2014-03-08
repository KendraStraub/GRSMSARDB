
Namespace LightSwitchApplication

    Public Class SendAnEmail

        Private Sub SendMyEmailOnCommand_Execute()
            Dim newEmail = DataWorkspace.ApplicationData.ProxyEmails.AddNew()
            With newEmail
                .RecipientEmailAddress = "thomas_colson@nps.gov"
                .RecipientName = "Thomas Colson"
                .SenderEmailAddress = "thomas_colson@nps.gov"
                .SenderName = "Thomas Colson"
            End With

            DataWorkspace.ApplicationData.SaveChanges()
            newEmail.Delete()
            DataWorkspace.ApplicationData.SaveChanges()
        End Sub
    End Class

End Namespace
