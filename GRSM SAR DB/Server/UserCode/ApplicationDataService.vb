
Namespace LightSwitchApplication

    Public Class ApplicationDataService


        Private Sub QRYFitYear_PreprocessQuery(FitYear As System.Nullable(Of Integer), ByRef query As System.Linq.IQueryable(Of LightSwitchApplication.Employee))
            query = query.Where(Function(a) (a.DaysSinceLastFit))

        End Sub

        'Private Sub QRYSarCert_PreprocessQuery(MaxRank As System.Nullable(Of Integer), ByRef query As System.Linq.IQueryable(Of LightSwitchApplication.Employee))
        ' query = From q In query
        '  From c In q.SARCertifications.Rank
        '  Where c.Rank <= MaxRank
        '  Group q By New q.Id Into g()
        ' Select g.FirstOrDefault()
        ' End Sub

        Private Sub ProxyEmails_Inserted(entity As ProxyEmail)
            Dim sSubject = "Test Email."
            Dim carRtn = Environment.NewLine & Environment.NewLine

            Dim sMessage = "The following email has come from a button on LightSwitch..." & carRtn
            sMessage += "Testing 1, 2, 3!!"

            ' Create the MailHelper class created in the Server project.
            Dim mailHelper As New EMailHelper(entity.SenderEmailAddress, _
                                             entity.RecipientName, _
                                             entity.RecipientEmailAddress, _
                                             entity.RecipientName, _
                                             sSubject, _
                                             sMessage)

            mailHelper.SendMail()
        End Sub

        Private Sub QRYDateFit_PreprocessQuery(Year As System.Nullable(Of Integer), ByRef query As System.Linq.IQueryable(Of LightSwitchApplication.Employee))
            'if a user-supplied year range is selected, use that value in the query. 
            'values are "1","2", and "10" (integers)

            If Year.HasValue Then

                Dim lastYear = DateAndTime.Now.AddYears(-Year.ToString)
                query = From q In query
                        Where q.DateFit > lastYear
                 Select q
                'if the user does nothing, the search 
                'defaults to certifications valid within the last year
            Else
                Dim lastYear = DateAndTime.Now.AddYears(-1)
                query = From q In query
                        Where q.DateFit > lastYear
                 Select q
            End If
        End Sub


        Private Sub QRYMaxRank_PreprocessQuery(MaxRank As System.Nullable(Of Integer), ByRef query As System.Linq.IQueryable(Of LightSwitchApplication.Employee))
            'if a user supplied certification is selected from the list of 
            'certification in the certification table
            'then select ALL employees that have that certification
            'AND ALL employees that have a lesser certification
            'allows IC to generate a list of responders that 
            'contains anyone that is certified at level [q]
            'or below
            If MaxRank.HasValue Then
                query = From q In query
                        Where q.SARCertifications.Rank <= MaxRank
                                            Select q
                'at runtime the search results include all responders
                'regardless of their certification, which an IC may want
                'this is the default, if IC wants filter, then use picklist
            Else
                query = From q In query
                        Select q
            End If
        End Sub

        Private Sub QRYParameters_PreprocessQuery(Year As System.Nullable(Of Integer), MaxRank As System.Nullable(Of Integer), ByRef query As System.Linq.IQueryable(Of LightSwitchApplication.Employee))
            'if a user supplied certification is selected from the list of 
            'certification in the certification table
            'then select ALL employees that have that certification
            'AND ALL employees that have a lesser certification
            'allows IC to generate a list of responders that 
            'contains anyone that is certified at level [q]
            'or below
            If MaxRank.HasValue Then
                query = From q In query
                        Where q.SARCertifications.Rank <= MaxRank
                                            Select q
                'at runtime the search results include all responders
                'regardless of their certification, which an IC may want
                'this is the default, if IC wants filter, then use picklist
            Else
                query = From q In query
                        Select q
            End If
            'if a user-supplied year range is selected, use that value in the query. 
            'values are "1","2", and "10" (integers)
            If Year.HasValue Then
                Dim lastYear = DateAndTime.Now.AddYears(-Year.ToString)
                query = From q In query
                        Where q.DateFit > lastYear
                 Select q
                'if the user does nothing, the search 
                'defaults to certifications valid within the last year
            Else
                Dim lastYear = DateAndTime.Now.AddYears(-1)
                query = From q In query
                        Where q.DateFit > lastYear
                 Select q
            End If
        End Sub
    End Class

End Namespace
