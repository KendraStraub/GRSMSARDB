
Namespace LightSwitchApplication

    Public Class ManageDivisions

        Private Sub ManageDivisions_InitializeDataWorkspace(ByVal saveChangesTo As Global.System.Collections.Generic.List(Of Global.Microsoft.LightSwitch.IDataService))
            ' Write your code here.
            Me.ParkDivisionProperty = New ParkDivision()
        End Sub

        Private Sub ManageDivisions_Saved()
            ' Write your code here.
            Me.Close(False)
            Application.Current.ShowDefaultScreen(Me.ParkDivisionProperty)
        End Sub

    End Class

End Namespace