﻿
Namespace LightSwitchApplication

    Public Class CreateNewTeamMember

        Private Sub CreateNewTeamMember_InitializeDataWorkspace(ByVal saveChangesTo As Global.System.Collections.Generic.List(Of Global.Microsoft.LightSwitch.IDataService))
            ' Write your code here.
            Me.EmployeeProperty = New Employee()
        End Sub

        Private Sub CreateNewTeamMember_Saved()
            ' Write your code here.
            Me.Close(False)
            Application.Current.ShowDefaultScreen(Me.EmployeeProperty)
        End Sub

    End Class

End Namespace