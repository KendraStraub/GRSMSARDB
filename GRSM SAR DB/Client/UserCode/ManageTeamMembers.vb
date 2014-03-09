
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




        Private Sub ImportFromExcel_Execute()
            'Use the Office Integration Pack Extension to do a variety of things with MS Office
            'Download here: http://visualstudiogallery.msdn.microsoft.com/35c4cf2a-5148-4716-afcf-0ccf8899cabf 

            OfficeIntegration.Excel.Import(Me.Employees)
        End Sub
    End Class

End Namespace
