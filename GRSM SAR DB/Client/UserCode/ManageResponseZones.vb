
Namespace LightSwitchApplication

    Public Class ManageResponseZones

        Private Sub ManageResponseZones_InitializeDataWorkspace(ByVal saveChangesTo As Global.System.Collections.Generic.List(Of Global.Microsoft.LightSwitch.IDataService))
            ' Write your code here.
            Me.ResponseZonesProperty = New ResponseZones()
        End Sub

        Private Sub ManageResponseZones_Saved()
            ' Write your code here.
            Me.Close(False)
            Application.Current.ShowDefaultScreen(Me.ResponseZonesProperty)
        End Sub

    End Class

End Namespace