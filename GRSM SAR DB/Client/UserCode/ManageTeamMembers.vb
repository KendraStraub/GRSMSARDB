
Namespace LightSwitchApplication

    Public Class ManageTeamMembers


        Private Sub AddZone_Execute()
            If (Employees.SelectedItem IsNot Nothing And AllResponseZones.SelectedItem IsNot Nothing) Then
                ' Loop through the category list and see if we already have this category
                ' If so, don't allow it to be added (again)
                For Each c In xref_EmployeeZonesCollection
                    If (c.ResponseZones Is AllResponseZones.SelectedItem) Then
                        Exit Sub
                    End If

                Next
                ' Add the new category to the category list
                Dim cc As xref_EmployeeZones = xref_EmployeeZonesCollection.AddNew()
                cc.Employee = Employees.SelectedItem
                cc.ResponseZones = AllResponseZones.SelectedItem
            End If
        End Sub

        Private Sub DeleteZone_Execute()
            If (xref_EmployeeZonesCollection.SelectedItem IsNot Nothing) Then
                xref_EmployeeZonesCollection.SelectedItem.Delete()
            End If
        End Sub




        Private Sub TestEmail_Execute()
            If Employees.SelectedItem.WorkEmail Is Nothing Then
            Else
                Dim newEmail = DataWorkspace.ApplicationData.ProxyEmails.AddNew()
                With newEmail
                    .RecipientEmailAddress = Employees.SelectedItem.WorkEmail
                    .RecipientName = Employees.SelectedItem.Summary
                    .SenderEmailAddress = "GRSM_EMERGENCY_CALLOUT@NPS.GOV"
                    .SenderName = "Dispatch"
                    .Message = "This is a test of the GRSM SAR Notification System."
                End With
                DataWorkspace.ApplicationData.SaveChanges()
                newEmail.Delete()
                DataWorkspace.ApplicationData.SaveChanges()
            End If
            If Employees.SelectedItem.HomeEmail Is Nothing Then
            Else
                Dim newEmail = DataWorkspace.ApplicationData.ProxyEmails.AddNew()
                With newEmail
                    .RecipientEmailAddress = Employees.SelectedItem.HomeEmail
                    .RecipientName = Employees.SelectedItem.Summary
                    .SenderEmailAddress = "GRSM_EMERGENCY_CALLOUT@NPS.GOV"
                    .SenderName = "Dispatch"
                    .Message = "This is a test of the GRSM SAR Notification System."
                End With
                DataWorkspace.ApplicationData.SaveChanges()
                newEmail.Delete()
                DataWorkspace.ApplicationData.SaveChanges()
            End If
            If Employees.SelectedItem.WorksSMS Is Nothing Then
            Else
                Dim newEmail = DataWorkspace.ApplicationData.ProxyEmails.AddNew()
                With newEmail
                    .RecipientEmailAddress = Employees.SelectedItem.WorksSMS
                    .RecipientName = Employees.SelectedItem.Summary
                    .SenderEmailAddress = "GRSM_EMERGENCY_CALLOUT@NPS.GOV"
                    .SenderName = "Dispatch"
                    .Message = "This is a test of the GRSM SAR Notification System."
                End With
                DataWorkspace.ApplicationData.SaveChanges()
                newEmail.Delete()
                DataWorkspace.ApplicationData.SaveChanges()
            End If
            If Employees.SelectedItem.PersonalSMS Is Nothing Then
            Else
                Dim newEmail = DataWorkspace.ApplicationData.ProxyEmails.AddNew()
                With newEmail
                    .RecipientEmailAddress = Employees.SelectedItem.PersonalSMS
                    .RecipientName = Employees.SelectedItem.Summary
                    .SenderEmailAddress = "GRSM_EMERGENCY_CALLOUT@NPS.GOV"
                    .SenderName = "Dispatch"
                    .Message = "This is a test of the GRSM SAR Notification System."
                End With
                DataWorkspace.ApplicationData.SaveChanges()
                newEmail.Delete()
                DataWorkspace.ApplicationData.SaveChanges()
            End If

        End Sub

        Private Sub NotifyEmployee_Execute()
            If Employees.SelectedItem.WorkEmail Is Nothing Then
            Else
                Dim newEmail = DataWorkspace.ApplicationData.ProxyEmails.AddNew()
                With newEmail
                    .RecipientEmailAddress = Employees.SelectedItem.WorkEmail
                    .RecipientName = Employees.SelectedItem.Summary
                    .SenderEmailAddress = "GRSM_EMERGENCY_CALLOUT@NPS.GOV"
                    .SenderName = "Dispatch"
                    .Message = "Please call dispatch on radio or (865)436-1230."
                End With
                DataWorkspace.ApplicationData.SaveChanges()
                newEmail.Delete()
                DataWorkspace.ApplicationData.SaveChanges()
            End If
            If Employees.SelectedItem.HomeEmail Is Nothing Then
            Else
                Dim newEmail = DataWorkspace.ApplicationData.ProxyEmails.AddNew()
                With newEmail
                    .RecipientEmailAddress = Employees.SelectedItem.HomeEmail
                    .RecipientName = Employees.SelectedItem.Summary
                    .SenderEmailAddress = "GRSM_EMERGENCY_CALLOUT@NPS.GOV"
                    .SenderName = "Dispatch"
                    .Message = "Please call dispatch on radio or (865)436-1230."
                End With
                DataWorkspace.ApplicationData.SaveChanges()
                newEmail.Delete()
                DataWorkspace.ApplicationData.SaveChanges()
            End If
            If Employees.SelectedItem.WorksSMS Is Nothing Then
            Else
                Dim newEmail = DataWorkspace.ApplicationData.ProxyEmails.AddNew()
                With newEmail
                    .RecipientEmailAddress = Employees.SelectedItem.WorksSMS
                    .RecipientName = Employees.SelectedItem.Summary
                    .SenderEmailAddress = "GRSM_EMERGENCY_CALLOUT@NPS.GOV"
                    .SenderName = "Dispatch"
                    .Message = "Please call dispatch on radio or (865)436-1230."
                End With
                DataWorkspace.ApplicationData.SaveChanges()
                newEmail.Delete()
                DataWorkspace.ApplicationData.SaveChanges()
            End If
            If Employees.SelectedItem.PersonalSMS Is Nothing Then
            Else
                Dim newEmail = DataWorkspace.ApplicationData.ProxyEmails.AddNew()
                With newEmail
                    .RecipientEmailAddress = Employees.SelectedItem.PersonalSMS
                    .RecipientName = Employees.SelectedItem.Summary
                    .SenderEmailAddress = "GRSM_EMERGENCY_CALLOUT@NPS.GOV"
                    .SenderName = "Dispatch"
                    .Message = "Please call dispatch on radio or (865)436-1230."
                End With
                DataWorkspace.ApplicationData.SaveChanges()
                newEmail.Delete()
                DataWorkspace.ApplicationData.SaveChanges()
            End If
        End Sub
    End Class

End Namespace
