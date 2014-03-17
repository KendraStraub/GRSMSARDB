
Namespace LightSwitchApplication

    Public Class SearchEmployees


        Private Sub ExportToExcel_Execute()
            OfficeIntegration.Excel.Export(Me.IRT)
        End Sub

        Private Sub EmailBlast_Execute()
            'loops through the results of the query
            'using unique employee ID as a limiter
            'then loops through the array 4 times
            'to send an email to up to 4 addresses
            'an employee can have
            Dim sendarray = From detail In IRT
                            Where detail.Id = detail.Id
                            Select detail
            'send email to each employee work cell phone
            For Each d In sendarray
                Dim newEmail = DataWorkspace.ApplicationData.ProxyEmails.AddNew()
                With newEmail
                    .RecipientEmailAddress = d.WorksSMS
                    .RecipientName = d.Summary
                    .SenderEmailAddress = "big.foot@sasquatch.com"
                    .SenderName = "Big Foot"
                    .Message = Type + " " + Location + " " + "Priority: " + Priority + " Staging at " + StagingArea + ". Call (865)436-1230."
                End With
                DataWorkspace.ApplicationData.SaveChanges()
                newEmail.Delete()
                DataWorkspace.ApplicationData.SaveChanges()
            Next

            'send email to each employee personal cell phone
            For Each e In sendarray
                Dim newEmail = DataWorkspace.ApplicationData.ProxyEmails.AddNew()
                With newEmail
                    .RecipientEmailAddress = e.PersonalSMS
                    .RecipientName = e.Summary
                    .SenderEmailAddress = "big.foot@sasquatch.com"
                    .SenderName = "Big Foot"
                    .Message = Type + " " + Location + " " + "Priority: " + Priority + " Staging at " + StagingArea + ". Call (865)436-1230."
                End With
                DataWorkspace.ApplicationData.SaveChanges()
                newEmail.Delete()
                DataWorkspace.ApplicationData.SaveChanges()
            Next
            'send email to each employee work email
            For Each f In sendarray
                Dim newEmail = DataWorkspace.ApplicationData.ProxyEmails.AddNew()
                With newEmail
                    .RecipientEmailAddress = f.WorkEmail
                    .RecipientName = f.Summary
                    .SenderEmailAddress = "big.foot@sasquatch.com"
                    .SenderName = "Big Foot"
                    .Message = Type + " " + Location + " " + "Priority: " + Priority + " Staging at " + StagingArea + ". Call (865)436-1230."
                End With
                DataWorkspace.ApplicationData.SaveChanges()
                newEmail.Delete()
                DataWorkspace.ApplicationData.SaveChanges()
            Next
            'send email to each employee personal email
            For Each g In sendarray
                Dim newEmail = DataWorkspace.ApplicationData.ProxyEmails.AddNew()
                With newEmail
                    .RecipientEmailAddress = g.HomeEmail
                    .RecipientName = g.Summary
                    .SenderEmailAddress = "big.foot@sasquatch.com"
                    .SenderName = "Big Foot"
                    .Message = Type + " " + Location + " " + "Priority: " + Priority + " Staging at " + StagingArea + ". Call (865)436-1230."
                End With
                DataWorkspace.ApplicationData.SaveChanges()
                newEmail.Delete()
                DataWorkspace.ApplicationData.SaveChanges()
            Next

        End Sub

 

        Private Sub OrderIsFilled_Execute()
            'loops through the results of the query
            'using unique employee ID as a limiter
            Dim sendarray = From detail In IRT
                            Where detail.Id = detail.Id
                            Select detail
            'send email to each employee work cell phone
            For Each h In sendarray
                Dim newEmail = DataWorkspace.ApplicationData.ProxyEmails.AddNew()
                With newEmail
                    .RecipientEmailAddress = h.WorksSMS
                    .RecipientName = h.Summary
                    .SenderEmailAddress = "big.foot@sasquatch.com"
                    .SenderName = "Big Foot"
                    .Message = "SAR resource order has been FILLED!!!" + Type + " " + Location + " " + "Priority: " + Priority + " Staging at " + StagingArea
                End With
                DataWorkspace.ApplicationData.SaveChanges()
                newEmail.Delete()
                DataWorkspace.ApplicationData.SaveChanges()
            Next
            'send email to each employee personal cell phone
            For Each i In sendarray
                Dim newEmail = DataWorkspace.ApplicationData.ProxyEmails.AddNew()
                With newEmail
                    .RecipientEmailAddress = i.PersonalSMS
                    .RecipientName = i.Summary
                    .SenderEmailAddress = "big.foot@sasquatch.com"
                    .SenderName = "Big Foot"
                    .Message = "SAR resource order has been FILLED!!!" + Type + " " + Location + " " + "Priority: " + Priority + " Staging at " + StagingArea
                End With
                DataWorkspace.ApplicationData.SaveChanges()
                newEmail.Delete()
                DataWorkspace.ApplicationData.SaveChanges()
            Next
            'send email to each employee work email 
            For Each j In sendarray
                Dim newEmail = DataWorkspace.ApplicationData.ProxyEmails.AddNew()
                With newEmail
                    .RecipientEmailAddress = j.WorkEmail
                    .RecipientName = j.Summary
                    .SenderEmailAddress = "big.foot@sasquatch.com"
                    .SenderName = "Big Foot"
                    .Message = "SAR resource order has been FILLED!!!" + Type + " " + Location + " " + "Priority: " + Priority + " Staging at " + StagingArea
                End With
                DataWorkspace.ApplicationData.SaveChanges()
                newEmail.Delete()
                DataWorkspace.ApplicationData.SaveChanges()
            Next
            'send email to each employee home email
            For Each k In sendarray
                Dim newEmail = DataWorkspace.ApplicationData.ProxyEmails.AddNew()
                With newEmail
                    .RecipientEmailAddress = k.HomeEmail
                    .RecipientName = k.Summary
                    .SenderEmailAddress = "big.foot@sasquatch.com"
                    .SenderName = "Big Foot"
                    .Message = "SAR resource order has been FILLED!!!" + Type + " " + Location + " " + "Priority: " + Priority + " Staging at " + StagingArea
                End With
                DataWorkspace.ApplicationData.SaveChanges()
                newEmail.Delete()
                DataWorkspace.ApplicationData.SaveChanges()
            Next
        End Sub

        Private Sub OrderIsCancelled_Execute()
            'loops through the results of the query
            'using unique employee ID as a limiter
            Dim sendarray = From detail In IRT
                            Where detail.Id = detail.Id
                            Select detail
            'send email to each employee work cell phone
            For Each l In sendarray
                Dim newEmail = DataWorkspace.ApplicationData.ProxyEmails.AddNew()
                With newEmail
                    .RecipientEmailAddress = l.WorksSMS
                    .RecipientName = l.Summary
                    .SenderEmailAddress = "big.foot@sasquatch.com"
                    .SenderName = "Big Foot"
                    .Message = "SAR Resource order CANCELLED!!!" + Type + " " + Location + " " + "Priority: " + Priority + " Staging at " + StagingArea
                End With
                DataWorkspace.ApplicationData.SaveChanges()
                newEmail.Delete()
                DataWorkspace.ApplicationData.SaveChanges()
            Next
            'send email to each employee personal cell phone
            For Each m In sendarray
                Dim newEmail = DataWorkspace.ApplicationData.ProxyEmails.AddNew()
                With newEmail
                    .RecipientEmailAddress = m.PersonalSMS
                    .RecipientName = m.Summary
                    .SenderEmailAddress = "big.foot@sasquatch.com"
                    .SenderName = "Big Foot"
                    .Message = "SAR Resource order CANCELLED!!!" + Type + " " + Location + " " + "Priority: " + Priority + " Staging at " + StagingArea
                End With
                DataWorkspace.ApplicationData.SaveChanges()
                newEmail.Delete()
                DataWorkspace.ApplicationData.SaveChanges()
            Next
            'send email to each employee work email 
            For Each n In sendarray
                Dim newEmail = DataWorkspace.ApplicationData.ProxyEmails.AddNew()
                With newEmail
                    .RecipientEmailAddress = n.WorkEmail
                    .RecipientName = n.Summary
                    .SenderEmailAddress = "big.foot@sasquatch.com"
                    .SenderName = "Big Foot"
                    .Message = "SAR Resource order CANCELLED!!!" + Type + " " + Location + " " + "Priority: " + Priority + " Staging at " + StagingArea
                End With
                DataWorkspace.ApplicationData.SaveChanges()
                newEmail.Delete()
                DataWorkspace.ApplicationData.SaveChanges()
            Next
            'send email to each employee home email
            For Each o In sendarray
                Dim newEmail = DataWorkspace.ApplicationData.ProxyEmails.AddNew()
                With newEmail
                    .RecipientEmailAddress = o.HomeEmail
                    .RecipientName = o.Summary
                    .SenderEmailAddress = "big.foot@sasquatch.com"
                    .SenderName = "Big Foot"
                    .Message = "SAR Resource order CANCELLED!!!" + Type + " " + Location + " " + "Priority: " + Priority + " Staging at " + StagingArea
                End With
                DataWorkspace.ApplicationData.SaveChanges()
                newEmail.Delete()
                DataWorkspace.ApplicationData.SaveChanges()
            Next
        End Sub
    End Class

End Namespace
