Imports System.IO
Imports System.Net
Imports System.Net.Mail

Public Class DrawSheets

    Private Sub CategoryDetermine(ByVal tournamentevent As String)
        'Defines the categories for each event
        Dim eventcategories() As String = {"\Children(8-10)", "\Children(11-14)", "\Junior(14-16)", "\Junior(17-19)", "\Adult(20-40)", "\Veteran(40+)"}

        Dim categorycounter As Integer = 0
        'Loops through each category in each event
        Do Until categorycounter = eventcategories.Count
            'If the event is sparring, checks either the male and female categories 
            If tournamentevent.Contains("Sparring") Then
                Dim sparringcounter As Integer = 0
                For sparringcounter = 0 To 1
                    If sparringcounter = 0 Then
                        'Calls the draw sheet generator for the males
                        DrawSheetGenerator(tournamentevent + eventcategories(categorycounter).Replace(")", " M)"))
                    Else
                        'Calls the draw sheet generator for the females
                        DrawSheetGenerator(tournamentevent + eventcategories(categorycounter).Replace(")", " F)"))
                    End If
                Next sparringcounter

            Else
                'Calls the draw sheet generator for non gendered events
                DrawSheetGenerator(tournamentevent + eventcategories(categorycounter))
            End If

            categorycounter += 1
        Loop
    End Sub

    Private Sub DrawSheetGenerator(ByVal tournamentEvent As String)
        Try
            'Tries to open the event file
            FileOpen(1, tournamentEvent + ".csv", OpenMode.Input)
        Catch
            'Removes the gender if looking for a gendered file in a non gendered event
            FileOpen(1, tournamentEvent.Replace(" M)", ")").Replace(" F)", ")") + ".csv", OpenMode.Input)
        End Try
        'Adds each participant to a list
        Dim participantslist As New List(Of String)
        Do Until EOF(1)
            Dim fullline As String = LineInput(1)
            participantslist.Add(fullline)
        Loop

        FileClose(1)
        'Saves the number of participants
        Dim participants As New Integer
        'Stops the program crashing when there are zero elements in the list
        Try
            participants = participantslist.Count
        Catch ex As Exception
            participants = 0
        End Try
        'Creates a list to store entrant information and array index location

        Dim entrantlocations(participants, 1) As String

        'Runs when there are more than 3 participants
        If participants > 3 Then
            'Clears all elements from the form so that a completely new draw sheet can be generated
            Me.Controls.Clear()
            'Defines the category size to the nearest highest power of 2
            Dim categorysize As Integer = 2 ^ Math.Ceiling(Math.Log(participants, 2))
            'Defines the x axis location of the column
            Dim column As Integer = 0
            'Works out the column index
            Dim columnnumber As Integer = Math.Ceiling(Math.Log(categorysize, 2))
            'Saves the amount of textboxes in each column
            Dim columnHeight As New List(Of Integer)
            'Defines the entryboxes in a 2d array
            Dim entryboxes(columnnumber, categorysize) As TextBox
            Dim x As Integer = 0
            Dim highestpowerof2 As New Integer
            'Calculates the nearest power of two below the number of participtants
            Do While x <= columnnumber + 2
                If participants - (x ^ 2) < 0 Then
                    highestpowerof2 = 2 ^ (x - 3)
                End If
                x += 1
            Loop
            'Defines how many competitors should be in the first round
            Dim firstround As Integer = (participants - highestpowerof2) * 2
            'Runs until there is two boxes in the final column
            Do While categorysize >= 2
                x = 0
                'Used as a second counter for the first round
                Dim y As Integer = 0
                'Defines the widening gap between boxes
                Dim offset As Integer = 20 * (2 ^ (column))
                'Runs until the counter reaches the entrant number
                Do While x < categorysize
                    'Runs if currently on the first competitor column
                    If column = 0 Then
                        'Runs whilst the counter is less than the calculated amount of rounds in the first round
                        Do While y < firstround
                            'Creates a new textbox
                            entryboxes(column, y) = New TextBox
                            'Defines the entrybox parameters

                            With entryboxes(column, y)
                                'Defines its size as a rectangle
                                .Size = New Size(200, 100)
                                'Determines its position based on column number and x number
                                .Location = New Point(20 + (230 * (column)), offset + (y * 40) * (2 ^ (column)))
                            End With
                            'Adds the entry box to the form
                            Me.Controls.Add(entryboxes(column, y))

                            y += 1
                            If y = firstround - 1 Then
                                columnHeight.Add(firstround)
                            End If
                        Loop

                    Else
                        'Creates a new textbox
                        entryboxes(column, x) = New TextBox
                        With entryboxes(column, x)
                            'Defines its size as a rectangle
                            .Size = New Size(200, 100)
                            'Determines its position based on column number and x number
                            .Location = New Point(20 + (230 * (column)), offset + (x * 40) * (2 ^ (column)))
                        End With
                        'Adds the entry box to the form
                        Me.Controls.Add(entryboxes(column, x))
                    End If
                    If x = categorysize - 1 And column <> 0 Then
                        columnHeight.Add(x + 1)
                    End If
                    x += 1

                Loop
                'Divides the columns into even factors of two
                categorysize = categorysize / 2
                column += 1
            Loop


            'Adds users to the draw sheet
            Dim counter As Integer = 0
            'Fills in the first column from top to bottom
            Do Until counter = columnHeight(0)
                Dim participant() As String = participantslist(counter).Split(",")

                entryboxes(0, counter).Text = participant(0)
                'Adds all participant information and index locations to list
                entrantlocations(counter, 0) = participantslist(counter)
                entrantlocations(counter, 1) = "(0," + counter.ToString + ")"

                counter += 1
            Loop
            'Fills in the second column from bottom to top as bi number is not known
            Dim counter2 As Integer = columnHeight(1)
            Do Until counter = participants And counter2 > 0
                Dim participant() As String = participantslist(counter).Split(",")
                'Increments backwards
                counter2 -= 1
                entryboxes(1, counter2).Text = participant(0)
                'Adds all participant information and index locations to the list
                entrantlocations(counter, 0) = participantslist(counter)
                entrantlocations(counter, 1) = "(1," + counter2.ToString + ")"
                counter += 1
            Loop

            'Allows the sorting process to run multiple times
            Dim loopcount As Integer = 0
            'Iterates through the draw sheet
            Do Until loopcount = participants
                Dim item As Integer = 0
                Do Until item = participants - 1
                    'Sort users by club, alphabetically
                    If String.Compare(entrantlocations(item, 0).Split(",")(6), entrantlocations(item + 1, 0).Split(",")(6)) > 0 Then
                        'Creates variable to store value of overwritten user
                        Dim temp As String = entrantlocations(item + 1, 0)
                        'Retrieves the location of the next entrant
                        Dim nextlocation As New Point(entrantlocations(item + 1, 1).Split(",")(0).Replace("(", ""), entrantlocations(item + 1, 1).Split(",")(1).Replace(")", ""))
                        'Retrieves the location of the original entrant
                        Dim originallocation As New Point(entrantlocations(item, 1).Split(",")(0).Replace("(", ""), entrantlocations(item, 1).Split(",")(1).Replace(")", ""))
                        'Swaps entrant usernames on draw sheet
                        entryboxes(nextlocation.X, nextlocation.Y).Text = entrantlocations(item, 0).Split(",")(0)
                        entryboxes(originallocation.X, originallocation.Y).Text = temp.Split(",")(0)
                        'Swaps entrant data in list
                        entrantlocations(item + 1, 0) = entrantlocations(item, 0)
                        entrantlocations(item + 1, 1) = "(" + nextlocation.X.ToString + "," + nextlocation.Y.ToString + ")"
                        entrantlocations(item, 0) = temp
                        entrantlocations(item, 1) = "(" + originallocation.X.ToString + "," + originallocation.Y.ToString + ")"
                        'matches += 1
                    End If
                    item += 1
                Loop
                loopcount += 1
            Loop

            'Resets loopcount, rather than creating a new variable
            loopcount = 0
            'Runs the loop for different amounts for odds or evens
            'Leaves middle intact for evens
            Dim loopend As New Integer
            If participants Mod 2 = 0 Then
                loopend = (participants \ 2) - 2
            Else
                loopend = (participants \ 2) - 1
            End If
            Do Until loopcount = loopend
                'Creates variable to store value of overwritten user
                Dim temp As String = entrantlocations(participants - (loopcount * 2) - 1, 0)
                'Retrieves the location of the next entrant
                Dim nextlocation As New Point(entrantlocations(participants - (loopcount * 2) - 1, 1).Split(",")(0).Replace("(", ""), entrantlocations(participants - (loopcount * 2) - 1, 1).Split(",")(1).Replace(")", ""))
                'Retrieves the location of the original entrant
                Dim originallocation As New Point(entrantlocations(loopcount * 2, 1).Split(",")(0).Replace("(", ""), entrantlocations(loopcount * 2, 1).Split(",")(1).Replace(")", ""))
                'Swaps entrant usernames on draw sheet
                entryboxes(nextlocation.X, nextlocation.Y).Text = entrantlocations(loopcount * 2, 0).Split(",")(0)
                entryboxes(originallocation.X, originallocation.Y).Text = temp.Split(",")(0)
                'Swaps entrant data in list
                entrantlocations(participants - (loopcount * 2) - 1, 0) = entrantlocations(loopcount * 2, 0)
                entrantlocations(participants - (loopcount * 2) - 1, 1) = "(" + nextlocation.X.ToString + "," + nextlocation.Y.ToString + ")"
                entrantlocations(loopcount * 2, 0) = temp
                entrantlocations(loopcount * 2, 1) = "(" + originallocation.X.ToString + "," + originallocation.Y.ToString + ")"
                loopcount += 1
            Loop
            'Runs if there are an even number of participants
            If participants Mod 2 = 0 Then
                'Saves the information in the first and last boxes
                Dim temp1 As String = entrantlocations(0, 0)
                Dim temp2 As String = entrantlocations(participants - 1, 0)
                'Retrieves the location of the final entrant
                Dim finallocation As New Point(entrantlocations(participants - 1, 1).Split(",")(0).Replace("(", ""), entrantlocations(participants - 1, 1).Split(",")(1).Replace(")", ""))
                'Retrieves the sheet location of each of the middle entrants
                Dim middlelocationfirst As New Point(entrantlocations(participants / 2, 1).Split(",")(0).Replace("(", ""), entrantlocations(participants / 2, 1).Split(",")(1).Replace(")", ""))
                Dim middlelocationsecond As New Point(entrantlocations(participants / 2 + 1, 1).Split(",")(0).Replace("(", ""), entrantlocations(participants / 2 + 1, 1).Split(",")(1).Replace(")", ""))
                'Replaces entrant information in the list
                entrantlocations(0, 0) = entrantlocations(participants / 2, 0)
                entrantlocations(participants / 2, 0) = temp1
                entrantlocations(participants - 1, 0) = entrantlocations(participants / 2 + 1, 0)
                entrantlocations(participants / 2 + 1, 0) = temp2
                'Replaces entrant information in the textboxes
                entryboxes(0, 0).Text = entrantlocations(0, 0).Split(",")(0)
                entryboxes(finallocation.X, finallocation.Y).Text = entrantlocations(participants - 1, 0).Split(",")(0)
                entryboxes(middlelocationfirst.X, middlelocationfirst.Y).Text = temp1.Split(",")(0)
                entryboxes(middlelocationsecond.X, middlelocationsecond.Y).Text = temp2.Split(",")(0)
            End If
            GenerateImage(tournamentEvent)
            EmailUsers(tournamentEvent)

            'Runs when there are three participants

        ElseIf participants = 3 Then
            Me.Controls.Clear()
            'Defines the entry boxes in a 2d array of textboxes
            Dim entryboxes(2, 3) As TextBox
            'Defines the column index
            Dim column As Integer = 0
            'Allows the x count to decrease by 1 between the first and second column
            Dim z As Integer = 0
            'Defines the height of the entry boxes
            Dim x As Integer = 0
            'Runs until 2 columns have been placed in the form
            Do Until column = 2
                'Caluulates the column offset and the height offset
                Dim offset As Integer = 20 * (2 ^ column)
                'Decreases the amount of rows in a column for each run through
                Do Until x = 4 - z
                    'Initialises a usable textbox in the array
                    entryboxes(column, x) = New TextBox
                    'Defines the parameters for the textbox
                    With entryboxes(column, x)
                        .Size = New Size(200, 100)
                        If column = 0 And x > 1 Then
                            'Separates the secondary group from the primary group
                            .Location = New Point(20 + (230 * (column)), offset * 5 + (x * 40) * (2 ^ (column)))
                        Else
                            .Location = New Point(20 + (230 * (column)), offset + (x * 40) * (2 ^ (column)))

                        End If
                        If x > 1 Then
                            'Highlights secondary group
                            .BackColor = Color.LightGray
                        End If
                    End With
                    'Adds the textbox to the form
                    Me.Controls.Add(entryboxes(column, x))
                    x += 1
                Loop
                column += 1
                x = 0
                z += 1
            Loop
            'Adds a final textbox to primary group. Needed as winner determines if secondary group is used
            entryboxes(2, 0) = New TextBox With {.Size = New Size(200, 100), .Location = New Point(480, 80)}
            Me.Controls.Add(entryboxes(2, 0))
            'Adds the users to the draw sheet
            entryboxes(0, 0).Text = participantslist(0).Split(",")(0)
            entryboxes(0, 1).Text = participantslist(1).Split(",")(0)
            entryboxes(0, 2).Text = participantslist(0).Split(",")(0) + " / " + participantslist(1).Split(",")(0)
            entryboxes(0, 3).Text = participantslist(2).Split(",")(0)
            entryboxes(1, 1).Text = participantslist(2).Split(",")(0)

            GenerateImage(tournamentEvent)
            EmailUsers(tournamentEvent)
        Else
            Debug.WriteLine("Event Invalid!")
        End If


    End Sub

    'Gets a list of the emails of the entrants in the current event.
    Private Function GetUserEmails(ByVal tournamentevent As String)
        Dim UserEmails As New List(Of String)
        'Opens the file, using CSV openmode
        Dim CSVRead As New FileIO.TextFieldParser(tournamentevent + ".csv") With {
            .Delimiters = New String() {","},
            .TextFieldType = FileIO.FieldType.Delimited
        }
        'Runs until the end of the file is reached
        While CSVRead.EndOfData = False
            'Gets the email on the current line
            Dim Email As String = CSVRead.ReadFields(1)
            UserEmails.Add(Email)
        End While
        CSVRead.Close()
        Return UserEmails
    End Function

    Private Sub EmailUsers(ByVal TournamentEvent As String)
        'Gets the user emails from the CSV file
        Dim UserEmails As List(Of String) = GetUserEmails(TournamentEvent)

        'Sends the drawsheet to each entrant
        For Each UserEmail In UserEmails
            'Defines the email contents and sender and recipient information
            Dim DrawSheetEmail As New MailMessage
            DrawSheetEmail.From = New MailAddress("TournamentManagerBot@gmail.com")
            DrawSheetEmail.To.Add(UserEmail)
            'Creates the email subject message
            Dim TournamentEventSplit As List(Of String) = TournamentEvent.Split("\").ToList
            DrawSheetEmail.Subject = "Drawsheet for: " + TournamentEventSplit(3) + " " + TournamentEventSplit(4) + " " + TournamentEventSplit(5)
            'Attatches the generated image to the email
            Dim DrawSheetAttachment As New Attachment(TournamentEvent + ".jpg")
            DrawSheetEmail.Attachments.Add(DrawSheetAttachment)
            'Attempts to send the email
            Try
                'Using Google's SMTP Server
                Dim SMTP As New SmtpClient("smtp.gmail.com")
                SMTP.Port = 587
                SMTP.EnableSsl = True
                'Logs into the Gmail account
                SMTP.Credentials = New NetworkCredential("TournamentManagerBot@gmail.com", "TournamentManager1234")
                SMTP.Send(DrawSheetEmail)

            Catch ex As Exception
                Debug.WriteLine("Invalid Email")
            End Try
        Next
    End Sub

    'Takes a screenshot of the windows form
    Private Sub GenerateImage(ByVal TournamentEvent As String)
        'Creates the dimensions of the image, using the form height and width
        Dim ScreenshotBMP As Bitmap = New Bitmap(Me.Width, Me.Height, Imaging.PixelFormat.Format32bppArgb)
        'Creates a graphics object for the form
        Dim ScreenshotGFX As Graphics = Graphics.FromImage(ScreenshotBMP)

        'Allows the program to catch up and process the data entries, before taking the screenshot
        Application.DoEvents()

        'Takes the screenshot
        ScreenshotGFX.CopyFromScreen(Me.Location.X, Me.Location.Y, 0, 0, Me.Size, CopyPixelOperation.SourceCopy)
        'Saves it in the tournament folder using the tournament category as a file name
        ScreenshotBMP.Save(TournamentEvent + ".jpg", Imaging.ImageFormat.Jpeg)
    End Sub


    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Show()
        'Retrieves all directory names in the tournament folder to find available tournaments
        Dim tournamentNames() As String = Directory.GetDirectories("C:\TransferredFiles\Tournaments\")
        Debug.WriteLine(tournamentNames(0))
        Dim x As Integer = 0
        'Runs for each tournament in the folder
        Try
            Do Until x = tournamentNames.Count
                'Opens the tournament information file
                Debug.WriteLine(tournamentNames(x))
                FileOpen(3, tournamentNames(x) + "\" + tournamentNames(x).Replace("C:\TransferredFiles\Tournaments\", "") + " Information.csv", OpenMode.Input)
                'Retrieves the tournament closing date
                Dim fullline() As String = LineInput(3).Split(",")
                'Obtains the tournament closing date
                Dim tournamentdate As Date
                'Runs for all different forms of date read as 0s are ignored by system
                Try
                    tournamentdate = DateTime.ParseExact(fullline(5), "dd/MM/yyyy HH:mm:ss", Nothing)
                Catch
                    tournamentdate = DateTime.ParseExact(fullline(5), "dd/MM/yyyy HH:mm", Nothing)
                Catch
                    tournamentdate = DateTime.ParseExact(fullline(5), "dd/MM/yyyy HH", Nothing)
                End Try
                'Runs if the tournament closing date has been passed
                If tournamentdate < DateTime.Now Then
                    'Gets all events in the tournament
                    Dim eventlist() As String = IO.Directory.GetDirectories(tournamentNames(x))
                    Dim eventcounter As Integer = 0
                    'Runs for each event
                    Do Until eventcounter = eventlist.Count
                        'Calls category determine to generate the draw sheet
                        CategoryDetermine(eventlist(eventcounter))
                        eventcounter += 1
                    Loop
                End If
                FileClose(3)
                x += 1
            Loop
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub
End Class
