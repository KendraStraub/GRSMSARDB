
Namespace LightSwitchApplication

    Public Class ApplicationDataService

 
        Private Sub QRYFitYear_PreprocessQuery(FitYear As System.Nullable(Of Integer), ByRef query As System.Linq.IQueryable(Of LightSwitchApplication.Employee))
            query = query.Where(Function(a) (a.DaysSinceLastFit))

        End Sub


        Private Sub ProxyEmails_Inserted(entity As ProxyEmail)

            Dim mailHelper As New EMailHelper(
                entity.SenderEmailAddress, _
                entity.SenderName, _
                entity.RecipientEmailAddress, _
                entity.RecipientName, _
                entity.Message)
            mailHelper.SendMail()
        End Sub

        'based on user selection of SAR Certification, will return results where SAR CERT rank is equal to or less than picked CERT
        Private Sub QRYParameters_PreprocessQuery(Year As System.Nullable(Of Integer), MaxRank As System.Nullable(Of Integer), MaxFit As System.Nullable(Of Integer), ByRef query As System.Linq.IQueryable(Of LightSwitchApplication.Employee))
            If MaxRank.HasValue Then
                query = From q In query
                        Where q.SARCertifications.Rank >= MaxRank
                                            Select q
            Else
                query = From q In query
                        Select q
            End If
            If Year.HasValue Then
                Dim lastYear = DateAndTime.Now.AddYears(-Year.ToString)
                query = From q In query
                        Where q.DateFit > lastYear
                 Select q
            Else
                Dim lastYear = DateAndTime.Now.AddYears(-1)
                query = From q In query
                        Where q.DateFit > lastYear
                 Select q
            End If
            If MaxFit.HasValue Then
                query = From q In query
                        Where q.FitnessLevel >= MaxFit
                                            Select q
            Else
                query = From q In query
                        Select q
            End If


        End Sub
        'Private Sub QRYZone_PreprocessQuery(ZoneID As System.Nullable(Of Integer), ByRef query As System.Linq.IQueryable(Of LightSwitchApplication.ResponseZones))
        ' query = From q In query
        ' Where q.xref_EmployeeZonesCollection.Where(Function(s) s.ResponseZones.Id = ZoneID).Count() > 0
        'End Sub


        'allow global viewing to all tables
        Private Sub Employees_CanRead(ByRef result As Boolean)
            result = Me.Application.User.HasPermission(Permissions.Viewer)
        End Sub
        Private Sub ParkDivisions_CanRead(ByRef result As Boolean)
            result = Me.Application.User.HasPermission(Permissions.Viewer)
        End Sub
        Private Sub ResponseZonesSet_CanRead(ByRef result As Boolean)
            result = Me.Application.User.HasPermission(Permissions.Viewer)
        End Sub
        Private Sub SARCertificationsSet_CanRead(ByRef result As Boolean)
            result = Me.Application.User.HasPermission(Permissions.Viewer)
        End Sub
        Private Sub SMSCarriers_CanRead(ByRef result As Boolean)
            result = Me.Application.User.HasPermission(Permissions.Viewer)
        End Sub
        Private Sub xref_EmployeeZonesSet_CanRead(ByRef result As Boolean)
            result = Me.Application.User.HasPermission(Permissions.Viewer)
        End Sub
        Private Sub YearRoundResSet_CanRead(ByRef result As Boolean)
            result = Me.Application.User.HasPermission(Permissions.Viewer)
        End Sub
        Private Sub ProxyEmails_CanRead(ByRef result As Boolean)
            result = Me.Application.User.HasPermission(Permissions.Viewer)
        End Sub


        'allow editors to add new employees and execute email sending
        Private Sub Employees_CanUpdate(ByRef result As Boolean)
            result = Me.Application.User.HasPermission(Permissions.Editing)
        End Sub
        Private Sub Employees_CanDelete(ByRef result As Boolean)
            result = Me.Application.User.HasPermission(Permissions.Editing)
        End Sub
        Private Sub Employees_CanInsert(ByRef result As Boolean)
            result = Me.Application.User.HasPermission(Permissions.Editing)
        End Sub
        Private Sub ProxyEmails_CanUpdate(ByRef result As Boolean)
            result = Me.Application.User.HasPermission(Permissions.Editing)
        End Sub
        Private Sub ProxyEmails_CanInsert(ByRef result As Boolean)
            result = Me.Application.User.HasPermission(Permissions.Editing)
        End Sub
        Private Sub ProxyEmails_CanDelete(ByRef result As Boolean)
            result = Me.Application.User.HasPermission(Permissions.Editing)
        End Sub


        'allow admins to edit pick-list tables
        Private Sub ParkDivisions_CanDelete(ByRef result As Boolean)
            result = Me.Application.User.HasPermission(Permissions.Administration)
        End Sub
        Private Sub ParkDivisions_CanUpdate(ByRef result As Boolean)
            result = Me.Application.User.HasPermission(Permissions.Administration)
        End Sub
        Private Sub ParkDivisions_CanInsert(ByRef result As Boolean)
            result = Me.Application.User.HasPermission(Permissions.Administration)
        End Sub
        Private Sub ResponseZonesSet_CanDelete(ByRef result As Boolean)
            result = Me.Application.User.HasPermission(Permissions.Administration)
        End Sub
        Private Sub ResponseZonesSet_CanUpdate(ByRef result As Boolean)
            result = Me.Application.User.HasPermission(Permissions.Administration)
        End Sub
        Private Sub ResponseZonesSet_CanInsert(ByRef result As Boolean)
            result = Me.Application.User.HasPermission(Permissions.Administration)
        End Sub
        Private Sub SARCertificationsSet_CanDelete(ByRef result As Boolean)
            result = Me.Application.User.HasPermission(Permissions.Administration)
        End Sub
        Private Sub SARCertificationsSet_CanUpdate(ByRef result As Boolean)
            result = Me.Application.User.HasPermission(Permissions.Administration)
        End Sub
        Private Sub SARCertificationsSet_CanInsert(ByRef result As Boolean)
            result = Me.Application.User.HasPermission(Permissions.Administration)
        End Sub
        Private Sub SMSCarriers_CanDelete(ByRef result As Boolean)
            result = Me.Application.User.HasPermission(Permissions.Administration)
        End Sub
        Private Sub SMSCarriers_CanUpdate(ByRef result As Boolean)
            result = Me.Application.User.HasPermission(Permissions.Administration)
        End Sub
        Private Sub SMSCarriers_CanInsert(ByRef result As Boolean)
            result = Me.Application.User.HasPermission(Permissions.Administration)
        End Sub
        Private Sub xref_EmployeeZonesSet_CanDelete(ByRef result As Boolean)
            result = Me.Application.User.HasPermission(Permissions.Administration)
        End Sub
        Private Sub xref_EmployeeZonesSet_CanUpdate(ByRef result As Boolean)
            result = Me.Application.User.HasPermission(Permissions.Administration)
        End Sub
        Private Sub xref_EmployeeZonesSet_CanInsert(ByRef result As Boolean)
            result = Me.Application.User.HasPermission(Permissions.Administration)
        End Sub
        Private Sub YearRoundResSet_CanDelete(ByRef result As Boolean)
            result = Me.Application.User.HasPermission(Permissions.Administration)
        End Sub
        Private Sub YearRoundResSet_CanUpdate(ByRef result As Boolean)
            result = Me.Application.User.HasPermission(Permissions.Administration)
        End Sub
        Private Sub YearRoundResSet_CanInsert(ByRef result As Boolean)
            result = Me.Application.User.HasPermission(Permissions.Administration)
        End Sub
    End Class




End Namespace
