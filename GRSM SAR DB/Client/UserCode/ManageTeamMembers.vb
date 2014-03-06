
Namespace LightSwitchApplication

    Public Class ManageTeamMembers

        Private Sub AddCertification_Execute()
            If (Employees.SelectedItem IsNot Nothing And AllCertifications.SelectedItem IsNot Nothing) Then
                ' Loop through the category list and see if we already have this category
                ' If so, don't allow it to be added (again)
                For Each c In xref_EmployeeCertsCollection
                    If (c.Certifications Is AllCertifications.SelectedItem) Then
                        Exit Sub
                    End If
                Next
                ' Add the new category to the category list
                Dim cc As xref_EmployeeCerts = xref_EmployeeCertsCollection.AddNew()
                cc.Employee = Employees.SelectedItem
                cc.Certifications = AllCertifications.SelectedItem
            End If
        End Sub

        Private Sub RemoveCertification_Execute()
            If (xref_EmployeeCertsCollection.SelectedItem IsNot Nothing) Then
                xref_EmployeeCertsCollection.SelectedItem.Delete()
            End If
        End Sub

        Private Sub AddCertification_CanExecute(ByRef result As Boolean)
            ' Write your code here.

        End Sub
    End Class

End Namespace
