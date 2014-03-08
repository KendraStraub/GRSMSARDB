
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

        Private Sub QRYDateFit_PreprocessQuery(ByRef query As System.Linq.IQueryable(Of LightSwitchApplication.Employee))
            Dim lastYear = DateAndTime.Now.AddYears(-1)
            query = From q In query
                    Where q.DateFit > lastYear
             Select q
        End Sub
    End Class

End Namespace
