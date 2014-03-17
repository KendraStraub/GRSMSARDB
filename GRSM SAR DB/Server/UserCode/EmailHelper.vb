Imports System.Net
Imports System.Net.Mail

Namespace LightSwitchApplication
    Public Class EMailHelper

        Private Property SMTPServer As String = My.Settings.SMTPServer
        Private Property SMTPPort As Integer = My.Settings.SMTPPort

        Private Property MailFrom As String
        Private Property MailFromName As String
        Private Property MailTo As String
        Private Property MailToName As String
        Private Property MailBody As String

        Sub New(ByVal SendFrom As String,
                ByVal SendFromName As String, _
                ByVal SendTo As String, _
                ByVal SendToName As String, _
                ByVal Body As String)
            _MailFrom = SendFrom
            _MailFromName = SendFromName
            _MailTo = SendTo
            _MailToName = SendToName
            _MailBody = Body
        End Sub

        Public Sub SendMail()


            Dim mail As New MailMessage
            Dim mailFrom As New Mail.MailAddress(_MailFrom, _MailFromName)
            Dim mailTo As New Mail.MailAddress(_MailTo, _MailToName)

            With mail
                .From = mailFrom
                .To.Add(mailTo)
                .Body = _MailBody
            End With

            Dim smtp As New SmtpClient(_SMTPServer, _SMTPPort)
            smtp.EnableSsl = False


            smtp.Send(mail)
        End Sub

    End Class
End Namespace
