
Namespace LightSwitchApplication

    Public Class ManageTeamMembers

        Private Sub AddCertification_Execute()
            If (Employees.SelectedItem IsNot Nothing And AllCertifications.SelectedItem IsNot Nothing) Then
                ' Loop through the category list and see if we already have this category
                ' If so, don't allow it to be added (again)
                For Each c In xref_EmployeeCertsCollection
                    If (c.SARCertificationsSetItem Is AllCertifications.SelectedItem) Then
                        Exit Sub
                    End If
                Next
                ' Add the new category to the category list
                Dim cc As xref_EmployeeSARCerts = xref_EmployeeCertsCollection.AddNew()
                cc.Employee = Employees.SelectedItem
                cc.SARCertificationsSetItem = AllCertifications.SelectedItem
            End If
        End Sub

        Private Sub RemoveCertification_Execute()
            If (xref_EmployeeCertsCollection.SelectedItem IsNot Nothing) Then
                xref_EmployeeCertsCollection.SelectedItem.Delete()
            End If
        End Sub

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

        Private Sub AddTechCert_Execute()
            If (Employees.SelectedItem IsNot Nothing And AllTechCertifications.SelectedItem IsNot Nothing) Then
                ' Loop through the category list and see if we already have this category
                ' If so, don't allow it to be added (again)
                For Each c In xref_EmployeeTechCerts
                    If (c.TechRescueCerts Is AllTechCertifications.SelectedItem) Then
                        Exit Sub
                    End If
                Next
                ' Add the new category to the category list
                Dim cc As xref_EmployeeTechCert = xref_EmployeeTechCerts.AddNew()
                cc.Employee = Employees.SelectedItem
                cc.TechRescueCerts = AllTechCertifications.SelectedItem
            End If
        End Sub

        Private Sub DeleteTechCert_Execute()
            If (xref_EmployeeTechCerts.SelectedItem IsNot Nothing) Then
                xref_EmployeeTechCerts.SelectedItem.Delete()
            End If
        End Sub
    End Class

End Namespace
