
Namespace LightSwitchApplication

    Public Class SearchEmployees


   

        Private Sub EmailBlast_Execute()


            'loops through the results of the query
            'using unique employee ID as a limiter
            'then loops through the array 4 times
            'to send an email to up to 4 addresses
            'an employee can have
            Dim sendarray = From detail In IRT
                            Where detail.Id = detail.Id
                            Select detail
            'send email to each employee work cell phone
      
            For Each d In sendarray
                If d.WorksSMS Is Nothing Then
                Else
                    Dim newEmail = DataWorkspace.ApplicationData.ProxyEmails.AddNew()
                    With newEmail
                        .RecipientEmailAddress = d.WorksSMS
                        .RecipientName = d.Summary
                        .SenderEmailAddress = "GRSM_EMERGENCY_CALLOUT@NPS.GOV"
                        .SenderName = "Dispatch"
                        .Message = Type + " " + Location + " " + "Priority: " + Priority + " Staging at " + StagingArea + ". Call (865)436-1230."
                    End With
                    DataWorkspace.ApplicationData.SaveChanges()
                    newEmail.Delete()
                    DataWorkspace.ApplicationData.SaveChanges()
                End If
            Next



                'send email to each employee personal cell phone
            For Each e In sendarray
                If e.PersonalSMS Is Nothing Then
                Else
                    Dim newEmail = DataWorkspace.ApplicationData.ProxyEmails.AddNew()
                    With newEmail
                        .RecipientEmailAddress = e.PersonalSMS
                        .RecipientName = e.Summary
                        .SenderEmailAddress = "GRSM_EMERGENCY_CALLOUT@NPS.GOV"
                        .SenderName = "Dispatch"
                        .Message = Type + " " + Location + " " + "Priority: " + Priority + " Staging at " + StagingArea + ". Call (865)436-1230."
                    End With
                    DataWorkspace.ApplicationData.SaveChanges()
                    newEmail.Delete()
                    DataWorkspace.ApplicationData.SaveChanges()
                End If
            Next



                'send email to each employee work email
            For Each f In sendarray
                If f.WorkEmail Is Nothing Then
                Else
                    Dim newEmail = DataWorkspace.ApplicationData.ProxyEmails.AddNew()
                    With newEmail
                        .RecipientEmailAddress = f.WorkEmail
                        .RecipientName = f.Summary
                        .SenderEmailAddress = "GRSM_EMERGENCY_CALLOUT@NPS.GOV"
                        .SenderName = "Dispatch"
                        .Message = Type + " " + Location + " " + "Priority: " + Priority + " Staging at " + StagingArea + ". Call (865)436-1230."
                    End With
                    DataWorkspace.ApplicationData.SaveChanges()
                    newEmail.Delete()
                    DataWorkspace.ApplicationData.SaveChanges()
                End If
            Next



                'send email to each employee personal email
            For Each g In sendarray
                If g.HomeEmail Is Nothing Then
                Else
                    Dim newEmail = DataWorkspace.ApplicationData.ProxyEmails.AddNew()
                    With newEmail
                        .RecipientEmailAddress = g.HomeEmail
                        .RecipientName = g.Summary
                        .SenderEmailAddress = "GRSM_EMERGENCY_CALLOUT@NPS.GOV"
                        .SenderName = "Dispatch"
                        .Message = Type + " " + Location + " " + "Priority: " + Priority + " Staging at " + StagingArea + ". Call (865)436-1230."
                    End With
                    DataWorkspace.ApplicationData.SaveChanges()
                    newEmail.Delete()
                    DataWorkspace.ApplicationData.SaveChanges()
                End If
            Next



        End Sub

 

        Private Sub OrderIsFilled_Execute()
            'loops through the results of the query
            'using unique employee ID as a limiter
            Dim sendarray = From detail In IRT
                            Where detail.Id = detail.Id
                            Select detail
            'send email to each employee work cell phone
            For Each h In sendarray
                If h.WorksSMS Is Nothing Then
                Else

                    Dim newEmail = DataWorkspace.ApplicationData.ProxyEmails.AddNew()
                    With newEmail
                        .RecipientEmailAddress = h.WorksSMS
                        .RecipientName = h.Summary
                        .SenderEmailAddress = "GRSM_EMERGENCY_CALLOUT@NPS.GOV"
                        .SenderName = "Dispatch"
                        .Message = "SAR resource order has been FILLED!!!" + Type + " " + Location + " " + "Priority: " + Priority + " Staging at " + StagingArea
                    End With
                    DataWorkspace.ApplicationData.SaveChanges()
                    newEmail.Delete()
                    DataWorkspace.ApplicationData.SaveChanges()
                End If
            Next


            'send email to each employee personal cell phone
            For Each i In sendarray
                If i.PersonalSMS Is Nothing Then
                Else

                    Dim newEmail = DataWorkspace.ApplicationData.ProxyEmails.AddNew()
                    With newEmail
                        .RecipientEmailAddress = i.PersonalSMS
                        .RecipientName = i.Summary
                        .SenderEmailAddress = "GRSM_EMERGENCY_CALLOUT@NPS.GOV"
                        .SenderName = "Dispatch"
                        .Message = "SAR resource order has been FILLED!!!" + Type + " " + Location + " " + "Priority: " + Priority + " Staging at " + StagingArea
                    End With
                    DataWorkspace.ApplicationData.SaveChanges()
                    newEmail.Delete()
                    DataWorkspace.ApplicationData.SaveChanges()
                End If
            Next


            'send email to each employee work email 
            For Each j In sendarray
                If j.WorkEmail Is Nothing Then
                Else

                    Dim newEmail = DataWorkspace.ApplicationData.ProxyEmails.AddNew()
                    With newEmail
                        .RecipientEmailAddress = j.WorkEmail
                        .RecipientName = j.Summary
                        .SenderEmailAddress = "GRSM_EMERGENCY_CALLOUT@NPS.GOV"
                        .SenderName = "Dispatch"
                        .Message = "SAR resource order has been FILLED!!!" + Type + " " + Location + " " + "Priority: " + Priority + " Staging at " + StagingArea
                    End With
                    DataWorkspace.ApplicationData.SaveChanges()
                    newEmail.Delete()
                    DataWorkspace.ApplicationData.SaveChanges()
                End If
            Next


            'send email to each employee home email
            For Each k In sendarray
                If k.HomeEmail Is Nothing Then
                Else

                    Dim newEmail = DataWorkspace.ApplicationData.ProxyEmails.AddNew()
                    With newEmail
                        .RecipientEmailAddress = k.HomeEmail
                        .RecipientName = k.Summary
                        .SenderEmailAddress = "GRSM_EMERGENCY_CALLOUT@NPS.GOV"
                        .SenderName = "Dispatch"
                        .Message = "SAR resource order has been FILLED!!!" + Type + " " + Location + " " + "Priority: " + Priority + " Staging at " + StagingArea
                    End With
                    DataWorkspace.ApplicationData.SaveChanges()
                    newEmail.Delete()
                    DataWorkspace.ApplicationData.SaveChanges()
                End If
            Next


        End Sub

        Private Sub OrderIsCancelled_Execute()
            'loops through the results of the query
            'using unique employee ID as a limiter
            Dim sendarray = From detail In IRT
                            Where detail.Id = detail.Id
                            Select detail
            'send email to each employee work cell phone
            For Each l In sendarray
                If l.WorksSMS Is Nothing Then
                Else

                    Dim newEmail = DataWorkspace.ApplicationData.ProxyEmails.AddNew()
                    With newEmail
                        .RecipientEmailAddress = l.WorksSMS
                        .RecipientName = l.Summary
                        .SenderEmailAddress = "GRSM_EMERGENCY_CALLOUT@NPS.GOV"
                        .SenderName = "Dispatch"
                        .Message = "SAR Resource order CANCELLED!!!" + Type + " " + Location + " " + "Priority: " + Priority + " Staging at " + StagingArea
                    End With
                    DataWorkspace.ApplicationData.SaveChanges()
                    newEmail.Delete()
                    DataWorkspace.ApplicationData.SaveChanges()
                End If
            Next


            'send email to each employee personal cell phone
            For Each m In sendarray
                If m.PersonalSMS Is Nothing Then
                Else

                    Dim newEmail = DataWorkspace.ApplicationData.ProxyEmails.AddNew()
                    With newEmail
                        .RecipientEmailAddress = m.PersonalSMS
                        .RecipientName = m.Summary
                        .SenderEmailAddress = "GRSM_EMERGENCY_CALLOUT@NPS.GOV"
                        .SenderName = "Dispatch"
                        .Message = "SAR Resource order CANCELLED!!!" + Type + " " + Location + " " + "Priority: " + Priority + " Staging at " + StagingArea
                    End With
                    DataWorkspace.ApplicationData.SaveChanges()
                    newEmail.Delete()
                    DataWorkspace.ApplicationData.SaveChanges()
                End If
            Next


            'send email to each employee work email 
            For Each n In sendarray
                If n.WorkEmail Is Nothing Then
                Else

                    Dim newEmail = DataWorkspace.ApplicationData.ProxyEmails.AddNew()
                    With newEmail
                        .RecipientEmailAddress = n.WorkEmail
                        .RecipientName = n.Summary
                        .SenderEmailAddress = "GRSM_EMERGENCY_CALLOUT@NPS.GOV"
                        .SenderName = "Dispatch"
                        .Message = "SAR Resource order CANCELLED!!!" + Type + " " + Location + " " + "Priority: " + Priority + " Staging at " + StagingArea
                    End With
                    DataWorkspace.ApplicationData.SaveChanges()
                    newEmail.Delete()
                    DataWorkspace.ApplicationData.SaveChanges()
                End If
            Next


            'send email to each employee home email
            For Each o In sendarray
                If o.HomeEmail Is Nothing Then
                Else

                    Dim newEmail = DataWorkspace.ApplicationData.ProxyEmails.AddNew()
                    With newEmail
                        .RecipientEmailAddress = o.HomeEmail
                        .RecipientName = o.Summary
                        .SenderEmailAddress = "GRSM_EMERGENCY_CALLOUT@NPS.GOV"
                        .SenderName = "Dispatch"
                        .Message = "SAR Resource order CANCELLED!!!" + Type + " " + Location + " " + "Priority: " + Priority + " Staging at " + StagingArea
                    End With
                    DataWorkspace.ApplicationData.SaveChanges()
                    newEmail.Delete()
                    DataWorkspace.ApplicationData.SaveChanges()
                End If
            Next
        End Sub

        Private Sub EmailBlast_CanExecute(ByRef result As Boolean)
            result = Me.Application.User.HasPermission(Permissions.Editing)

        End Sub

        Private Sub OrderIsCancelled_CanExecute(ByRef result As Boolean)
            result = Me.Application.User.HasPermission(Permissions.Editing)

        End Sub

        Private Sub OrderIsFilled_CanExecute(ByRef result As Boolean)
            result = Me.Application.User.HasPermission(Permissions.Editing)

        End Sub

        Private Sub ExportToExcel_Execute()
            OfficeIntegration.Excel.Export(IRT)
        End Sub
        Public Function GetIRTCSV() As String
            Dim csv As New System.Text.StringBuilder()
            Dim i As Integer = 0

            For Each a In IRT
                If i = 0 Then
                    csv.AppendFormat("Summary" & "," & "Summary" & "," & "Summary" & System.Environment.NewLine, a)
                End If
                csv.AppendFormat(a.Summary & "," & a.Summary & "," & a.Summary & System.Environment.NewLine, a)
                i = i + 1
            Next

            If csv.Length > 0 Then
                Return csv.ToString(0, csv.Length - 1)
            Else
                Return ""
            End If
        End Function
        Private Sub Test_Execute()
            Dim csv As String = GetIRTCSV()
            AddHandler Me.FindControl("Test").ControlAvailable, (Function(object1, eventargs1)
                                                                     Dim btnExport As Button = DirectCast(eventargs1.Control, Button)
                                                                     AddHandler btnExport.Click,
                                                                     (Function(object2, eventargs2)
                                                                          Dim dialog = New SaveFileDialog()
                                                                          dialog.Filter = "CSV (*.csv)|*.csv"
                                                                          If dialog.ShowDialog() = True Then
                                                                              Using stream As New StreamWriter(dialog.OpenFile())
                                                                                  stream.Write(csv)
                                                                                  stream.Close()
                                                                              End Using
                                                                          End If
                                                                      End Function)
                                                                 End Function)
        End Sub
    End Class

End Namespace
