
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

        Private Sub AddWaterCertification_Execute()
            If (Employees.SelectedItem IsNot Nothing And AllWaterRescueCerts.SelectedItem IsNot Nothing) Then
                ' Loop through the category list and see if we already have this category
                ' If so, don't allow it to be added (again)
                For Each c In xref_EmployeeWaterCerts
                    If (c.WaterRescueCerts Is AllWaterRescueCerts.SelectedItem) Then
                        Exit Sub
                    End If
                Next
                ' Add the new category to the category list
                Dim cc As xref_EmployeeWaterCert = xref_EmployeeWaterCerts.AddNew()
                cc.Employee = Employees.SelectedItem
                cc.WaterRescueCerts = AllWaterRescueCerts.SelectedItem
            End If
        End Sub

        Private Sub DeleteWaterCertification_Execute()
            If (xref_EmployeeWaterCerts.SelectedItem IsNot Nothing) Then
                xref_EmployeeWaterCerts.SelectedItem.Delete()
            End If
        End Sub
    End Class

End Namespace
