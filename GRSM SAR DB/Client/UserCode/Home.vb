
Namespace LightSwitchApplication

    Public Class Home

        Private Sub SARCerts_Execute()
            Application.ShowManageSARCertifications()
        End Sub

        Private Sub Divisions_Execute()
            Application.ShowManageDivisions()
        End Sub

        Private Sub Zones_Execute(ByRef result As Boolean)
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
    End Class

End Namespace
