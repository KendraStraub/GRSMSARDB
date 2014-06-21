
Namespace LightSwitchApplication

    Public Class Home

        Private Sub SARCerts_Execute()
            Application.ShowManageSARCertifications()
        End Sub

        Private Sub Divisions_Execute()
            Application.ShowManageDivisions()
        End Sub

        Private Sub Zones_Execute()
            Application.ShowManageResponseZones()
        End Sub

        Private Sub View_Employees_Execute()
            Application.ShowEmpoyeeDetail()
        End Sub

        Private Sub AddEmployee_Execute()
            Application.ShowCreateNewTeamMember()
        End Sub

        Private Sub EditEmployee_Execute()
            Application.ShowManageTeamMembers()
        End Sub

        Private Sub SMS_Execute()
            Application.ShowManageSMSCarriers()

        End Sub

        Private Sub IRT_Execute()
            Application.ShowSearchEmployees()

        End Sub



        Private Sub TestTheSystem_Execute()
            If EmailAddress Is Nothing Then
            Else
                Dim newEmail = DataWorkspace.ApplicationData.ProxyEmails.AddNew()
                With newEmail
                    .RecipientEmailAddress = EmailAddress
                    .RecipientName = Me.Application.User.FullName
                    .SenderEmailAddress = "GRSM_EMERGENCY_CALLOUT@NPS.GOV"
                    .SenderName = "Dispatch"
                    .Message = "Warp speed Mr. Zulu! The engines appear to the working!"
                End With
                DataWorkspace.ApplicationData.SaveChanges()
                newEmail.Delete()
                DataWorkspace.ApplicationData.SaveChanges()
            End If

        End Sub

        Private Sub MIRT_Execute()
            Application.ShowMedicIRT()

        End Sub
    End Class

End Namespace
