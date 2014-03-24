Imports System.Windows.Controls
Imports System.Text
Namespace LightSwitchApplication

    Public Class SearchEmployees

        'initialize the export button
        Private Sub SearchEmployees_Created()
            Dim CSVButton = Me.FindControl("ExportCSV")
            AddHandler CSVButton.ControlAvailable, AddressOf exportAvailable
        End Sub
        'call the export button
        Private Sub exportAvailable(sender As Object, e As ControlAvailableEventArgs)
            RemoveHandler Me.FindControl("ExportCSV").ControlAvailable, AddressOf exportAvailable
            Dim Button = DirectCast(e.Control, Button)
            AddHandler Button.Click, AddressOf exportClicked
        End Sub
        'once export is clicked open dialogue to save csv file
        Private Sub exportClicked(sender As Object, e As System.Windows.RoutedEventArgs)
            Dim collectionProperty As Microsoft.LightSwitch.Details.Client.IScreenCollectionProperty = Me.Details.Properties.IRT
            Dim intPageSize = collectionProperty.PageSize
            'Get the Current PageSize and store to variable
            collectionProperty.PageSize = 0

            Dim dialog = New SaveFileDialog()
            dialog.Filter = "Excel (*.xls)|*.xls"
            If dialog.ShowDialog() = True Then

                Using stream As New StreamWriter(dialog.OpenFile())
                    Dim csv As String = GetCSV()
                    stream.Write(csv)
                    stream.Close()
                    Me.ShowMessageBox("Excel File Created Successfully. " & vbCrLf & "NOTE: When you open excel file and if you receive prompt about invalid format then just click yes to continue.", "Excel Export", MessageBoxOption.Ok)
                End Using
            End If
            collectionProperty.PageSize = intPageSize
            'Reset the Current PageSize
        End Sub



        Private Function GetCSV() As String
            Dim csv As New StringBuilder()

            Dim i As Integer = 0
            Dim csvarray = From detail In IRT
            For Each r In csvarray
                '//HEADER
                If i = 0 Then
                    Dim c As Integer = 0
                    For Each prop In r.Details.Properties.All.OfType(Of Microsoft.LightSwitch.Details.IEntityStorageProperty)()
                        If c > 0 Then
                            csv.Append(vbTab)
                        End If
                        c = c + 1
                        csv.Append(prop.DisplayName)
                    Next
                End If
                csv.AppendLine("")

                '//DATA ROWS

                Dim c1 As Integer = 0
                For Each prop In r.Details.Properties.All().OfType(Of Microsoft.LightSwitch.Details.IEntityStorageProperty)()
                    If c1 > 0 Then
                        csv.Append(vbTab)
                    End If
                    c1 = c1 + 1
                    csv.Append(prop.Value)
                Next
                i = i + 1
            Next

            If csv.Length > 0 Then
                Return csv.ToString(0, csv.Length - 1)
            Else
                Return ""
            End If
        End Function

       












        'build a csv file by looping through query results
        Private Function GetTextCSV() As String
            Dim csv As New StringBuilder()
            Dim i As Integer = 0

            For Each u In IRT
                If i = 0 Then
                    csv.AppendFormat("Summary" & "," & _
                                        "ParkDivision" & "," & _
                                        "YearRoundRes" & "," & _
                                        "DateFit" & "," & _
                                        "SARCertifications" & "," & _
                                        "CLEO" & "," & _
                                        "Medic" & "," & _
                                        "Tracker" & "," & _
                                        "TechRescue" & "," & _
                                        "WorkEmail" & "," & _
                                        "HomeEmail" & "," & _
                                        "WorkMobile" & "," & _
                                        "PersonalMobile" & "," & _
                                        "WorkSMS" & "," & _
                                        "PersonalSMS" & _
                                        System.Environment.NewLine, u)
                End If
                csv.AppendFormat(u.Summary & "'," & _
                                    u.ParkDivision.Division & "," & _
                                    u.YearRoundRes.Residency & "," & _
                                    u.DateFit & "," & _
                                    u.SARCertifications.Certification & "," & _
                                    u.CLEO & "," & _
                                    u.MEDIC & "," & _
                                    u.Tracker & "," & _
                                    u.TechRescue & "," & _
                                    u.WorkEmail & "," & _
                                    u.HomeEmail & "," & _
                                    u.WorkMobile & "," & _
                                    u.PersonalMobile & "," & _
                                    u.WorksSMS & "," & _
                                    u.PersonalSMS & "," & _
                                    System.Environment.NewLine, u)
                i = i + 1
            Next

            If csv.Length > 0 Then
                Return csv.ToString(0, csv.Length - 1)
            Else
                Return ""
            End If
        End Function

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

        

    End Class

End Namespace
