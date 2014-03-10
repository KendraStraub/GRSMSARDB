
Namespace LightSwitchApplication

    Public Class ApplicationDataService


        Private Sub QRYFitYear_PreprocessQuery(FitYear As System.Nullable(Of Integer), ByRef query As System.Linq.IQueryable(Of LightSwitchApplication.Employee))
            query = query.Where(Function(a) (a.DaysSinceLastFit))

        End Sub


        Private Sub ProxyEmails_Inserted(entity As ProxyEmail)
            Dim sSubject = "Test Email."
            Dim carRtn = Environment.NewLine & Environment.NewLine

            Dim sMessage = "The following email has come from a button on LightSwitch..." & carRtn
            sMessage += "Testing 1, 2, 3!!"
            Dim mailHelper As New EMailHelper(entity.SenderEmailAddress, _
                                             entity.RecipientName, _
                                             entity.RecipientEmailAddress, _
                                             entity.RecipientName, _
                                             sSubject, _
                                             sMessage)
            mailHelper.SendMail()
        End Sub

        'Private Sub QRYDateFit_PreprocessQuery(Year As System.Nullable(Of Integer), ByRef query As System.Linq.IQueryable(Of LightSwitchApplication.Employee))
        ' If Year.HasValue Then
        '  Dim lastYear = DateAndTime.Now.AddYears(-Year.ToString)
        '  query = From q In query
        '        Where q.DateFit > lastYear
        '  Select q
        ' Else
        '  Dim lastYear = DateAndTime.Now.AddYears(-1)
        '  query = From q In query
        '          Where q.DateFit > lastYear
        '      Select q
        '  End If
        'End Sub
        ' Private Sub QRYMaxRank_PreprocessQuery(MaxRank As System.Nullable(Of Integer), ByRef query As System.Linq.IQueryable(Of LightSwitchApplication.Employee))
        '   If MaxRank.HasValue Then
        '      query = From q In query
        '             Where q.SARCertifications.Rank <= MaxRank
        '                       Select q
        '  Else
        '     query = From q In query
        '            Select q
        ' End If
        ' End Sub
        Private Sub QRYParameters_PreprocessQuery(Year As System.Nullable(Of Integer), MaxRank As System.Nullable(Of Integer), ByRef query As System.Linq.IQueryable(Of LightSwitchApplication.Employee))
            If MaxRank.HasValue Then
                query = From q In query
                        Where q.SARCertifications.Rank <= MaxRank
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
        End Sub
        'Private Sub QRYZone_PreprocessQuery(ZoneID As System.Nullable(Of Integer), ByRef query As System.Linq.IQueryable(Of LightSwitchApplication.ResponseZones))
        ' query = From q In query
        ' Where q.xref_EmployeeZonesCollection.Where(Function(s) s.ResponseZones.Id = ZoneID).Count() > 0
        'End Sub



    End Class




End Namespace
