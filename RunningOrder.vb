Imports System.IO
Imports System.Net
Imports System.Net.Mail
Imports Microsoft.Office.Interop
Imports Spire.Xls
Public Class RunningOrder

    Private Class Graph
        Public NodeGraph = New Dictionary(Of String, Dictionary(Of String, Integer))
        'Defines the backup graph, to save an unedited copy of the graph, for the dijkstra algorithm
        Public GraphBackup = New Dictionary(Of String, Dictionary(Of String, Integer))

        'Auto properties to encapsulate getters and setters
        Public Property Rings As Integer
        Public Property OpeningTime As DateTime
        Public Property ClosingTime As DateTime
        'Stores the difference between the provided times
        Public Property TournamentLength As DateTime

        Public Sub New(Opening As DateTime, Closing As DateTime, RingNum As Integer)
            'Assigns these values to the properties
            Rings = RingNum
            OpeningTime = Opening
            ClosingTime = Closing
            'Calculates the difference between the hours
            Dim HourDifference As Integer = Closing.Hour - Opening.Hour
            'Calculates the difference between the hours, ensuring that it is always positive
            Dim MinuteDifference As Integer = Closing.Minute - Opening.Minute
            'Checks whether the opening minute is larger than the closing
            If MinuteDifference < 0 Then
                'Therefore subtracts one hour if it is
                HourDifference -= 1
                'And converts the minute difference to a positive integer
                MinuteDifference = -MinuteDifference
            End If
            'Adds these values to the TournamentLength datetime object
            TournamentLength = DateTime.Parse((HourDifference.ToString + ":" + MinuteDifference.ToString + ":" + "00").ToString()).ToLongTimeString

        End Sub

        'Adds a new dictionary for the specified node to the graph
        Public Sub AddNode(ByVal NodeName As String)
            NodeGraph.add(NodeName, New Dictionary(Of String, Integer))
            'And the backup graph
            GraphBackup.add(NodeName, New Dictionary(Of String, Integer))
        End Sub

        'Adds a new edge to the node
        Public Sub AddEdge(ByVal InitialNode As String, ByVal EndNode As String, ByVal Weight As Integer)
            'Gets the current edges
            Dim CurrentEdges As Dictionary(Of String, Integer) = GraphBackup(InitialNode)
            'Adds the new edge weight
            CurrentEdges.Add(EndNode, Weight)
            'Updates the backup graph only as the nodegraph will be updated in the edge calculation process
            GraphBackup.item(InitialNode) = CurrentEdges
        End Sub


        Public Sub CalculateEdges(ByVal RootNode As String)
            'Calculates the estimated time for the event
            Dim EventTime As Integer = EstimateEventTime(RootNode)
            'Only runs if the event has enough participants to be viable (>3)
            If EventTime <> 0 Then
                'Checks if the event will contain umpires
                If RootNode.Contains("Adult") Or RootNode.Contains("Veteran") Then
                    'Cycles through each dictionary entry
                    For Each kvp As KeyValuePair(Of String, Dictionary(Of String, Integer)) In NodeGraph
                        'Only runs if the dictionary entry is not the same as the input entry
                        If kvp.Key <> RootNode Then
                            'Only runs if the dictionary entry is not an adult event
                            If Not (kvp.Key.Contains("Adult")) And Not (kvp.Key.Contains("Veteran")) Then
                                'Adds the edge from the root node, with the estimated time
                                AddEdge(RootNode, kvp.Key, EventTime)
                            End If
                        End If
                    Next
                Else    'Runs if the event is not an adult event
                    'Cycles through each event in the dictionary
                    For Each kvp As KeyValuePair(Of String, Dictionary(Of String, Integer)) In NodeGraph
                        'Only runs if the current entry is not the provided entry
                        If kvp.Key <> RootNode Then
                            'Adds the edge from the root node, with the estimated time
                            AddEdge(RootNode, kvp.Key, EventTime)
                        End If
                    Next
                End If

            End If
        End Sub

        Public Function Dijkstra(StartNode As String, EndNode As String) As List(Of String)
            'Creates a copy of the graph, as timetabling will rely on all events being present initially
            Dim GraphCopy As New Dictionary(Of String, Dictionary(Of String, Integer))
            For Each item As KeyValuePair(Of String, Dictionary(Of String, Integer)) In NodeGraph
                GraphCopy.Add(item.Key, item.Value)
            Next
            'Holds the current distances to each node
            Dim Distances = New Dictionary(Of String, Integer)
            'Cycles through each node in the graph
            For Each Node In GraphCopy
                'Sets the distance to the start node to be 0
                If Node.Key = StartNode Then
                    Distances(StartNode) = 0
                Else
                    'Sets the distance to all other nodes to be the max value
                    Distances.Add(Node.Key, Integer.MaxValue)
                End If
            Next
            'Stores the node visited prior to the one we're currently at
            Dim PreviousNode = New Dictionary(Of String, String)
            'Loops until all nodes have been visited
            Do Until GraphCopy.Count = 0
                'Looks for the shortest edge weight
                Dim CurrentShortestEdge As String = ""
                'Cycles through each node in the graph
                For Each Node In GraphCopy
                    'If shortest edge has no value, sets first encountered weight to be the shortest edge
                    If CurrentShortestEdge = "" Then
                        CurrentShortestEdge = Node.Key
                    ElseIf Distances(Node.Key) < Distances(CurrentShortestEdge) Then
                        'Otherwise, checks that the distance and edge weight is less than the current edge weight
                        CurrentShortestEdge = Node.Key  'And sets the current shortest edge to be that node
                    End If

                Next Node

                'This works out the shortest distance to each adjacent node
                For Each ConnectedNode In GraphCopy(CurrentShortestEdge)
                    'Gets the edge weight of the adjacent node
                    Dim Weight As Integer = ConnectedNode.Value
                    'Checks if the new value is lower than the shortest distance

                    If GraphCopy.ContainsKey(ConnectedNode.Key) And Distances(CurrentShortestEdge) + Weight < Distances(ConnectedNode.Key) Then
                        'Sets the previous node to the sghortest edge weight node
                        PreviousNode(ConnectedNode.Key) = CurrentShortestEdge
                        'Updates the distances for this node
                        Distances(ConnectedNode.Key) = Distances(CurrentShortestEdge) + Weight
                    End If

                Next ConnectedNode
                'As all potential paths from this node have been visited, we can remove the node from the graph copy
                GraphCopy.Remove(CurrentShortestEdge)
            Loop
            'This starts from the End Node
            Dim CurrentNode As String = EndNode
            'Holds the list of nodes making up the shortest path
            Dim ShortestPath = New List(Of String)
            'Runs backwards through the list, until the initial node is reached
            Do Until CurrentNode = StartNode
                'Stops any invalid vertex inputs from crashing the program
                Try
                    'Inserts the node into the shortest paths list
                    ShortestPath.Insert(0, CurrentNode)
                    CurrentNode = PreviousNode(CurrentNode)
                Catch
                    Debug.WriteLine("That node is not in the graph")
                    'Returns an empty list
                    Return Nothing
                    'Exits the subroutine
                    Exit Function
                End Try
            Loop
            'The initial vertex is now added the the start of the list
            ShortestPath.Insert(0, StartNode)

            GraphCopy.Clear()
            'Resets Graph Copy back to its initial state
            For Each item As KeyValuePair(Of String, Dictionary(Of String, Integer)) In GraphBackup
                GraphCopy.Add(item.Key, item.Value)
                'MsgBox(item.Value.Count)
            Next

            Return ShortestPath ' The list of nodes is returned
        End Function


    End Class

    'Function takes the graph and tournament file as parameters
    Private Sub GenerateRunningOrder(ByVal TournamentFile As String, ByVal TournamentGraph As Graph)

        'Creates the new excel application and active workbook
        Dim ExcelApp As Excel.Application
        Dim ExcelWorkbook As Excel.Workbook
        Dim ExcelWorksheet As Excel.Worksheet
        Dim ExcelRange As Excel.Range
        ExcelApp = CreateObject("Excel.Application")
        'Set to true for testing purposes
        ExcelApp.Visible = True
        ExcelWorkbook = ExcelApp.Workbooks.Add
        ExcelWorksheet = ExcelWorkbook.ActiveSheet

        'Saves the time that the spreadsheet is working with
        Dim CurrentTime As DateTime = TournamentGraph.OpeningTime

        'Keeps track of the remaining time available for events at each ring
        Dim RingRemainingTime As New List(Of DateTime)
        'Keeps track of the ring number,for adding to the excel sheet and to the list
        For RingExcelCounter As Integer = 1 To TournamentGraph.Rings
            'Adds the tournament length as remaining time
            RingRemainingTime.Add(TournamentGraph.TournamentLength)
            'Adds the times, horizontally, starting from the third column
            Dim CurrentColumn As Integer = 3
            Do Until CurrentTime > TournamentGraph.ClosingTime
                ExcelWorksheet.Cells(1, CurrentColumn).value = CurrentTime.ToShortTimeString
                'Increments the time
                CurrentTime = CurrentTime.AddMinutes(5)
                CurrentColumn += 1
            Loop

            'Adds each ring, seperated by three spaces to the graph
            ExcelWorksheet.Cells(RingExcelCounter * 3, 1).value = "Ring " & RingExcelCounter
        Next

        'Sets the fontsize to 20 for entire sheet
        ExcelRange = ExcelWorksheet.Range("A1:AV36")
        ExcelRange.Font.Size = 20

        'Gets all the available events in a tournament
        Dim AllEvents As List(Of String) = Directory.GetDirectories("C:\TransferredFiles\Tournaments\Adidas International\").ToList
        'Creates a list to store all the category information for those events
        Dim AllEventCategories As New List(Of String)
        'Cycles through each directory and adds the category csv file information to the list
        For Each TournamentEvent In AllEvents
            'Calls GetEventCategories to get all tournament event files in the folder
            Dim EventCategories As List(Of String) = GetEventCategories(TournamentEvent)
            For Each TournamentCategory In EventCategories
                AllEventCategories.Add(TournamentEvent + "\" + TournamentCategory)
            Next
        Next

        'Counter to keep track of how many events are remaining
        Dim AllEventCounter As Integer = 0
        'Runs until there are no more events to be added
        Do Until AllEventCategories.Count = 0
            'Runs when there are multiple adult events in a tournament
            If AdultEvents(TournamentGraph.NodeGraph) > 2 Then
                'Saves each adult event to be usedby the Djkstra algorithm
                Dim FirstAdultEvent As String = ""
                Dim SecondAdultEvent As String = ""
                'Gets the string name of the first encountered viable adult event
                For Each item In AllEventCategories
                    If item.Contains("Adult") Or item.Contains("Veteran") Then
                        If EstimateEventTime(item) <> 0 Then
                            FirstAdultEvent = item
                            Exit For
                        End If
                    End If
                Next
                'Gets the string name of the second encountered viable adult event
                For Each item In AllEventCategories
                    If item.Contains("Adult") Or item.Contains("Veteran") Then
                        'Ensures this event isn't the same one as the first event
                        If item <> FirstAdultEvent Then
                            If EstimateEventTime(item) <> 0 Then
                                SecondAdultEvent = item
                                Exit For
                            End If
                        End If
                    End If
                Next
                'Calculates the shortest path between these events
                Dim NodePath As List(Of String) = TournamentGraph.Dijkstra(FirstAdultEvent, SecondAdultEvent)
                'Adds these events to the running order
                Dim EventAdded = AddRunningOrderAdultEvent(NodePath, ExcelWorksheet, TournamentGraph.Rings)
                'Removes the event from the graph and list of events
                If EventAdded = True Then
                    For Each Node In NodePath
                        TournamentGraph.NodeGraph.remove(Node)
                        AllEventCategories.Remove(Node)
                    Next
                End If
                AllEventCounter += 1
                'Resets event counter if past list bounds
                If AllEventCounter >= AllEventCategories.Count Then
                    AllEventCounter = 0
                End If

            Else ' Runs when tournaments have less than 2 adult events
                'Adds the event directly to the running order
                Dim EventAdded As Boolean = AddRunningOrderEvent(AllEventCategories(AllEventCounter), ExcelWorksheet, TournamentGraph.Rings)
                'Removes the event from the graph and list
                If EventAdded = True Then
                    TournamentGraph.NodeGraph.remove(AllEventCategories(AllEventCounter))
                    AllEventCategories.Remove(AllEventCategories(AllEventCounter))
                End If
                AllEventCounter += 1
                'Resets event counter if past list bounds
                If AllEventCounter >= AllEventCategories.Count Then
                    AllEventCounter = 0
                End If
            End If
        Loop

        ExcelRange.EntireColumn.AutoFit()
        'Saves the workbook
        ExcelWorkbook.SaveAs(TournamentFile.Replace("Information", "RunningOrder"))
        'Closes the workbook
        ExcelWorkbook.Close()
        'Quits Excel
        ExcelApp.Quit()

        GenerateExcelImage(TournamentFile)
    End Sub

    'Subroutine to add the event to the spreadsheet
    Private Function AddRunningOrderAdultEvent(ByVal NodePath As List(Of String), ByVal RunningOrder As Excel.Worksheet, ByVal RingCount As Integer)
        Dim Added As Boolean = False
        'Calculates the total available events for columns.
        Dim TotalColumnCount As Integer = GetColumnCount(RunningOrder, 1)
        'Creates a list of integers to hold the amount of used columns in each ring
        Dim RingColumnCount As New List(Of Integer)
        'Calculates the columns for each ring
        For RingCounter As Integer = 1 To RingCount
            RingColumnCount.Add(GetColumnCount(RunningOrder, RingCounter * 3))
        Next

        'Holds the event information lists for each event.
        Dim EventPartition As New List(Of List(Of String))
        'Holds the required times for the events.
        Dim TournamentEventTimes As New List(Of Integer)
        'Holds the required cells for the events.
        Dim TournamentEventCells As New List(Of Integer)
        Dim TotalAdultEventCells As Integer = 0
        'Iterates through each event, adding the split to event partition, the estimated time to the estimated time list, and the required cells
        For Each TournamentEvent In NodePath
            EventPartition.Add(TournamentEvent.Split("\").ToList)
            'Calculates the allocated time for the event
            Dim EventTime As Integer = EstimateEventTime(TournamentEvent)
            TournamentEventTimes.Add(EventTime)
            TournamentEventCells.Add(EventTime / 5)
            TotalAdultEventCells += (EventTime / 5)
        Next
        'Adds the required event name parts to the final string, for each event
        Dim StandardEventName As New List(Of String)
        For Each TournamentEvent In EventPartition
            StandardEventName.Add(TournamentEvent(4) + " " + TournamentEvent(5).Replace(".csv", ""))
        Next
        'Calculates the ring with the most free space for the events
        Dim AvailableRing As Integer = GetAvailableRing(RingColumnCount, TotalAdultEventCells, TotalColumnCount)
        'Runs if ring was found
        If AvailableRing <> -1 Then
            'Works out if the events will conflict with other events at their current place
            Dim EventCategoryConflict As New List(Of Boolean)
            For Each TournamentEvent In NodePath
                EventCategoryConflict.Add(EventConflict(TournamentEvent, RunningOrder, RingCount, TotalAdultEventCells, RingColumnCount(AvailableRing), AvailableRing))
            Next
            'Runs only if no conflicts are met
            If Not EventCategoryConflict.Contains(True) Then
                Dim TournamentEventCounter As Integer = 0
                'Runs through the shortest path
                For Each TournamentEvent In NodePath
                    'Adds the event to the specified ring
                    Added = AddEventToWorksheet(AvailableRing, TournamentEventCells(TournamentEventCounter), TotalColumnCount, RingColumnCount(AvailableRing), StandardEventName(TournamentEventCounter), RunningOrder)
                    'Updates the column count for that ring
                    RingColumnCount(AvailableRing) += TournamentEventCells(TournamentEventCounter)
                    TournamentEventCounter += 1
                Next
            Else ' Otherwise, runs for each ring, until  unconflicting spaces are found
                For Ring As Integer = 0 To RingCount - 1
                    'Calculates whether there is a conflict at the current ring
                    Dim NewEventCategoryConflict As Boolean = False
                    For Each TournamentEvent In NodePath
                        NewEventCategoryConflict = EventConflict(TournamentEvent, RunningOrder, RingCount, TotalAdultEventCells, RingColumnCount(Ring), Ring)
                    Next
                    'Runs if no conflict has been detected
                    If NewEventCategoryConflict = False Then
                        Dim TournamentEventCounter As Integer = 0
                        'Adds each event to the spreadsheet, and updates column counter and event counter
                        For Each TournamentEvent In NodePath
                            Added = AddEventToWorksheet(Ring, TournamentEventCells(TournamentEventCounter), TotalColumnCount, RingColumnCount(Ring), StandardEventName(TournamentEventCounter), RunningOrder)
                            RingColumnCount(AvailableRing) += TournamentEventCells(TournamentEventCounter)
                            TournamentEventCounter += 1
                        Next
                        Exit For
                    End If
                Next
            End If
        End If
        Return Added
    End Function

    'Subroutine to add the event to the spreadsheet
    Private Function AddRunningOrderEvent(ByVal TournamentEvent As String, ByVal RunningOrder As Excel.Worksheet, ByVal RingCount As Integer)
        Dim Added As Boolean = False
        'Calculates the total available events for columns.
        Dim TotalColumnCount As Integer = GetColumnCount(RunningOrder, 1)
        'Creates a list of integers to hold the amount of used columns in each ring
        Dim RingColumnCount As New List(Of Integer)
        'Calculates the columns for each ring
        For RingCounter As Integer = 1 To RingCount
            RingColumnCount.Add(GetColumnCount(RunningOrder, RingCounter * 3))
        Next

        'Splits each aspect of the event, by \ symbol
        Dim EventPartition As List(Of String) = TournamentEvent.Split("\").ToList
        'Adds the required event name parts to the final string
        Dim StandardEventName As String = EventPartition(4) + " " + EventPartition(5).Replace(".csv", "")
        'Calculates the allocated time for the event
        Dim TournamentEventTime As Integer = EstimateEventTime(TournamentEvent)
        'Calculates the required cells for the event.
        Dim TournamentEventCells As Integer = TournamentEventTime / 5
        'Calculates the ring with the most free space for the event
        Dim AvailableRing As Integer = GetAvailableRing(RingColumnCount, TournamentEventCells, TotalColumnCount)
        If AvailableRing <> -1 Then
            'Works out if the event will conflict with other events at its current place
            Dim EventCategoryConflict As Boolean = EventConflict(TournamentEvent, RunningOrder, RingCount, TournamentEventCells, RingColumnCount(AvailableRing), AvailableRing)
            'Runs only if no conflicts are met
            If EventCategoryConflict = False Then
                Added = AddEventToWorksheet(AvailableRing, TournamentEventCells, TotalColumnCount, RingColumnCount(AvailableRing), StandardEventName, RunningOrder)
            Else ' Otherwise, runs for each ring, until an unconflicting space is found
                For Ring As Integer = 0 To RingCount - 1
                    'Calcuates whether there is a conflict at the current ring
                    Dim NewEventCategoryConflict As Boolean = EventConflict(TournamentEvent, RunningOrder, RingCount, TournamentEventCells, RingColumnCount(Ring), Ring)
                    'Runs if no conflict has been detected
                    If NewEventCategoryConflict = False Then
                        'Adds the event to the spreadsheet and exits the loop
                        Added = AddEventToWorksheet(Ring, TournamentEventCells, TotalColumnCount, RingColumnCount(Ring), StandardEventName, RunningOrder)
                        Exit For
                    End If
                Next
            End If
        Else
            MsgBox(TournamentEvent & "   " & TournamentEventCells)
        End If
        Return Added
    End Function

    'Calculates the amount of filled in cells in a row, starting from the third cell
    Private Function GetColumnCount(ByVal RunningOrderExcel As Excel.Worksheet, ByVal Column As Integer)
        'Stores running total of column count
        Dim ColumnCount As Integer = 0
        'Starts from the third cell
        Dim CurrentColumn As Integer = 3
        'Runs until an empty cell is reached
        Do Until RunningOrderExcel.Cells()(Column, CurrentColumn).text.ToString = ""
            CurrentColumn += 1
            ColumnCount += 1
        Loop
        'Returns the calculated amount
        Return ColumnCount

    End Function

    'Calculates the ring with the most empty cells
    Private Function GetAvailableRing(ByVal RingColumnCount As List(Of Integer), ByVal TournamentEventCells As Integer, ByVal TotalColumnCount As Integer)
        'Initialises with a null ring
        Dim AvailableRing As Integer = -1
        'Initialises the lowest amount of available cells
        Dim CurrentMostAvailableCells As Integer = 0
        'Iterates through each ring
        For RingCounter As Integer = 0 To RingColumnCount.Count - 1
            'Calculates the available cells in the current ring
            Dim RingAvailableCells As Integer = TotalColumnCount - TournamentEventCells - RingColumnCount(RingCounter)
            'Checks if the amount is greater than the current amount of available cells
            If RingAvailableCells > CurrentMostAvailableCells Then
                'If it is, then updates values
                CurrentMostAvailableCells = RingAvailableCells
                AvailableRing = RingCounter
            End If
        Next
        'Returns the calculated ring with most free cells
        Return AvailableRing
    End Function



    'Physically adds the event information to the spread sheet
    Private Function AddEventToWorksheet(ByVal AvailableRing As Integer, ByVal TournamentEventCells As Integer, ByVal TotalColumnCount As Integer, ByVal RingColumnCount As Integer, ByVal StandardEventName As String, ByVal RunningOrder As Excel.Worksheet)
        'Keeps track of whether the event was able to be added to the tournament
        Dim Added As Boolean = False
        'Only runs if there is sufficient space in the given ring row
        If TournamentEventCells < TotalColumnCount - RingColumnCount Then
            'Iterates through each cell required by the new event
            For CellCounter As Integer = 0 To TournamentEventCells - 1
                'Adds the event information to the cell
                RunningOrder.Cells((AvailableRing + 1) * 3, RingColumnCount + 3 + CellCounter).value = StandardEventName
            Next CellCounter
            RingColumnCount -= TournamentEventCells
            'Registers that the event was able to be added to the given ring
            Added = True
        End If
        Return Added
    End Function

    'Calculates when placing an event in a specific tournament timeslot conflicts with the same age range of an event in another tournament
    Private Function EventConflict(ByVal TournamentEvent As String, ByVal RunningOrder As Excel.Worksheet, ByVal RingCount As Integer, ByVal TournamentEventCells As Integer, ByVal RingColumnCount As Integer, ByVal AvailableRing As Integer)
        'Saves whether there has been an event time conflict
        Dim Conflict As Boolean = False
        'Splits each aspect of the event, by \ symbol
        Dim EventPartition As List(Of String) = TournamentEvent.Split("\").ToList
        'Gets the age group for the provided event
        Dim EventGroup As String = EventPartition(5).Replace(".csv", "").Replace(" M", "").Replace(" F", "").Replace(")", "")
        'Iterates through each ring on the spreadsheet
        For RingCounter As Integer = 1 To RingCount
            'Only runs for events other than the one the system is trying to add to
            If RingCounter <> AvailableRing + 1 Then
                'Iterates through each cell of the given ring that the provided event will run on in the available ring
                For CellCounter As Integer = 0 To TournamentEventCells
                    'Checks if the given cell contains the same age group as the provided event. Removes gendered parts from events
                    If RunningOrder.Cells(RingCounter * 3, RingColumnCount + 3 + CellCounter).text.contains(EventGroup) Then
                        'If it does, there has been an event conflict
                        Conflict = True

                    End If
                Next
            End If
        Next
        'Returns the calculated result
        Return Conflict
    End Function


    'Gets all csv category files in a folder
    Private Function GetEventCategories(ByVal TournamentEvent As String)
        Dim CategoriesList As New List(Of String)
        'Defines the file directory for the given event
        Dim EventDirectory As New IO.DirectoryInfo(TournamentEvent)
        'Gets all the csv event files in the directory
        Dim Categories As IO.FileInfo() = EventDirectory.GetFiles("*.csv")
        'Adds each event category to the list
        For Each Category As IO.FileInfo In Categories
            CategoriesList.Add(Category.ToString)
        Next
        Return CategoriesList
    End Function


    Public Shared Function EstimateEventTime(ByVal RootNode As String)
        Dim EventTime As Integer = 0
        'Reads all the users in the CSV file and gets the integer length of it
        Dim Participants As Integer = File.ReadAllLines(RootNode).Length
        'Stops maths errors resulting from logging 0
        If Participants <> 0 Then
            'Gets the round number by finding the log of the participants to the base of 2, then rounding
            Dim RoundCount As Integer = CInt(Math.Log(Participants, 2))
            'Differs for each event
            If RootNode.Contains("Sparring") Then
                'Each round lasts approximately, 2 minutes 30 seconds, with time for changeover
                EventTime = Math.Ceiling((RoundCount * 2.5) / 5) * 5
            ElseIf RootNode.Contains("Patterns") Then
                'Each round lasts approximately, 3 minutes, with time for changeover
                EventTime = Math.Ceiling((RoundCount * 3) / 5) * 5
            ElseIf RootNode.Contains("Self Defence") Then
                'Each round lasts approximately, 2 minutes, with time for changeover
                EventTime = Math.Ceiling((RoundCount * 2) / 5) * 5
            Else
                'Each round lasts approximately, 1 minute, with time for changeover
                EventTime = Math.Ceiling(Participants / 5) * 5
            End If
        End If
        Return EventTime
    End Function

    'Subroutine to convert an excel file to a png file
    Private Sub GenerateExcelImage(ByVal TournamentFile As String)
        'Opens the excel file for Spire.Xls
        Dim ExcelFile As New Spire.Xls.Workbook()
        ExcelFile.LoadFromFile(TournamentFile.Replace("Information", "RunningOrder"))
        'Gets the Running Order sheet
        Dim RunningOrderSheet As Spire.Xls.Worksheet = ExcelFile.Worksheets(0)
        'Saves the Excel File to a png image
        RunningOrderSheet.SaveToImage(TournamentFile.Replace("Information.csv", "RunningOrder.png"))
        'Clears the excel file from the disk
        File.Delete(TournamentFile.Replace("Information", "RunningOrder"))
    End Sub

    'Returns the number of adult events remaining in a tournament
    Private Function AdultEvents(ByVal AllEvents As Dictionary(Of String, Dictionary(Of String, Integer))) As Integer
        Dim AdultEventCounter As Integer = 0
        'Runs for each event and adds one to the counter once an adult event is found
        For Each TournamentEvent In AllEvents
            If TournamentEvent.Key.Contains("Adult") Or TournamentEvent.Key.Contains("Veteran") Then
                AdultEventCounter += 1
            End If
        Next

        Return AdultEventCounter
    End Function

    Private Sub StartDynamicCategories(ByVal TournamentDirectory As String)
        'Gets all folder directories in the given tournament directory
        Dim EventNames As List(Of String) = Directory.GetDirectories(TournamentDirectory).ToList

        'Runs for each tournament event
        For Each TournamentEvent In EventNames
            'Gets all categories for the given event
            Dim EventCategories As List(Of String) = GetEventCategories(TournamentEvent)
            'Runs for each individual event category
            For Each Category In EventCategories
                Dim Participants As Integer = File.ReadAllLines(TournamentEvent + "\" + Category).Length
                If Participants < 3 And Participants > 0 Then
                    Dim AvailableEvent As String = CheckForAvailableEvents(TournamentEvent, Category)
                    If AvailableEvent <> "" Then
                        SwitchCSVFiles(TournamentEvent, Category, AvailableEvent)
                    End If
                End If
            Next
        Next
    End Sub

    'Checks for any available events, taking into account the gender split for events
    Private Function CheckForAvailableEvents(ByVal TournamentEvent As String, ByVal Category As String)
        Dim AvailableEvent As String = ""
        'Checks for available male events
        If Category.Contains(" M") Then
            Dim CategoryList As List(Of String) = {"Children(8-10).csv", "Children(11-14 M).csv", "Junior(14-16 M).csv", "Junior(17-19 M).csv", "Adult(20-40 M).csv", "Veteran(40+ M).csv"}.ToList
            AvailableEvent = CheckAdjacentEvents(CategoryList, TournamentEvent, Category)

            'Checks for available female events
        ElseIf Category.Contains(" F") Then
            Dim CategoryList As List(Of String) = {"Children(8-10).csv", "Children(11-14 F).csv", "Junior(14-16 F).csv", "Junior(17-19 F).csv", "Adult(20-40 F).csv", "Veteran(40+ F).csv"}.ToList
            AvailableEvent = CheckAdjacentEvents(CategoryList, TournamentEvent, Category)

            'Checks for available non gendered events
        Else
            Dim CategoryList As List(Of String) = {"Children(8-10).csv", "Children(8-10).csv", "Children(11-14).csv", "Junior(14-16).csv", "Junior(17-19).csv", "Adult(20-40).csv", "Veteran(40+).csv"}.ToList
            AvailableEvent = CheckAdjacentEvents(CategoryList, TournamentEvent, Category)

        End If
        Return AvailableEvent
    End Function

    'Finds the event from the available categories that has a viable number of participants and is suitable for the new entrants
    Private Function CheckAdjacentEvents(ByVal CategoryList As List(Of String), ByVal TournamentEvent As String, ByVal Category As String)
        Dim CategoryFound As String = ""

        'Gets the index of the event currently being checked
        Dim CurrentCategoryIndex As Integer = CategoryList.IndexOf(Category)

        'Only runs for events other than the Children(8-10) event as these must not change category due to safety concerns
        If CurrentCategoryIndex <> 0 Then
            'Checks for the previous category, excluding the Children(8-10) event due to safety concerns
            Dim StartCategoryIndex As Integer = CurrentCategoryIndex - 1
            If StartCategoryIndex = 0 Then
                StartCategoryIndex = 1
            End If
            'Runs for each adjacent category ie Junior(17-19) and Adult(20-40)
            Do Until StartCategoryIndex >= CurrentCategoryIndex + 1 Or CategoryFound <> ""
                'Gets the number of participants in the current event
                Dim Participants As Integer = File.ReadAllLines(TournamentEvent + "\" + CategoryList(StartCategoryIndex)).Length
                'Runs only for events that are valid, or will be after the new participant is added
                If Participants >= 2 Then
                    'Assigns the category from the category list
                    CategoryFound = CategoryList(StartCategoryIndex)
                End If
                'Skips the event we have already checked
                StartCategoryIndex += 2
            Loop

        End If
        Return CategoryFound
    End Function

    'Transfers the entrants from the old category to the new category
    Private Sub SwitchCSVFiles(ByVal TournamentEvent As String, ByVal OldCategory As String, ByVal NewCategory As String)
        'Creates a 2d list to store cell contents from the CSV file
        Dim OldCategoryCSVContents As New List(Of List(Of String))
        'Reads the old file
        Dim CSVRead As StreamReader = New StreamReader(TournamentEvent + "\" + OldCategory)
        Do Until CSVRead.Peek = -1
            Dim CSVLine As String = CSVRead.ReadLine
            'Gets the data from each row
            Dim CSVLineList As List(Of String) = CSVLine.Split(",").ToList
            'Then adds this to the 2d list
            OldCategoryCSVContents.Add(CSVLineList)
        Loop
        CSVRead.Close()

        'Opens the new category file and appends to it
        Using NewCSVWrite As StreamWriter =
            New StreamWriter(TournamentEvent + "\" + NewCategory, True)
            'Iterates through the 2d list, adding each row and column to the new category file
            For x As Integer = 0 To OldCategoryCSVContents.Count - 1
                Dim Fullline As String = OldCategoryCSVContents(x)(0)
                For i As Integer = 1 To OldCategoryCSVContents(x).Count - 1
                    Fullline = Fullline + "," + OldCategoryCSVContents(x)(i)
                Next
                'Writes the new data
                NewCSVWrite.Write(Fullline)
                'And a line break to continue creating new lines
                NewCSVWrite.WriteLine()
            Next
            NewCSVWrite.Close()
        End Using
    End Sub

    'Uses a predefined tournament due to lack of constant scanning availability
    Private Sub RunningOrder_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        'Opens the tournament information file
        Dim TournamentInfo As StreamReader = New StreamReader("C:\TransferredFiles\Tournaments\Adidas International\Adidas International Information.csv")
        'Seperates each cell
        Dim InfoParts As List(Of String) = TournamentInfo.ReadLine.Split(",").ToList
        'Closes the file
        TournamentInfo.Close()
        'Gets the number of rings available
        Dim RingNumber As Integer = CInt(InfoParts(InfoParts.Count - 1))
        'Splits the opening and closing times
        Dim TournamentTimes As List(Of String) = InfoParts(InfoParts.Count - 2).Split("/").ToList
        'The test data to be used
        Dim TournamentEventGraph As Graph = New Graph(TimeValue(Date.Parse(TournamentTimes(0))), TimeValue(Date.Parse(TournamentTimes(1))), RingNumber)

        'Reads whether the tournament organiser wants to use dynamic categories or not. Converts this to a boolean
        Dim DynamicCategories As Boolean = True
        If InfoParts(3).Contains("No") Then
            DynamicCategories = False
        End If

        If DynamicCategories = True Then
            StartDynamicCategories("C:\TransferredFiles\Tournaments\Adidas International")
        End If

        'Gets all folder directories in the given tournament directory
        Dim EventNames As List(Of String) = Directory.GetDirectories("C:\TransferredFiles\Tournaments\Adidas International").ToList

        'Runs for each tournament event
        For Each TournamentEvent In EventNames
            'Gets all categories for the given event
            Dim EventCategories As List(Of String) = GetEventCategories(TournamentEvent)
            'Adds the event as a node in the graph
            For Each Category In EventCategories
                If EstimateEventTime(TournamentEvent + "\" + Category) <> 0 Then
                    TournamentEventGraph.AddNode(TournamentEvent + "\" + Category)

                End If
            Next
        Next

        'Runs for each event in the graph
        For Each TournamentEvent As KeyValuePair(Of String, Dictionary(Of String, Integer)) In TournamentEventGraph.NodeGraph
            'Calculates the edges for this given event
            TournamentEventGraph.CalculateEdges(TournamentEvent.Key)
        Next
        'Assigns the changes to the graph backup to the original node graph
        TournamentEventGraph.NodeGraph = TournamentEventGraph.GraphBackup

        'Dim NodePath As List(Of String) = TournamentEventGraph.Dijkstra("C:\TransferredFiles\Tournaments\Adidas International\Patterns\Adult(20-40).csv", "C:\TransferredFiles\Tournaments\Adidas International\Special Technique\Veteran(40+).csv")
        'For Each node In NodePath
        '    MsgBox(node)
        'Next
        GenerateRunningOrder("C:\TransferredFiles\Tournaments\Adidas International\Adidas International Information.csv", TournamentEventGraph)
        SendRunningOrderEmail("C:\TransferredFiles\Tournaments\Adidas International\Adidas International Information.csv", EventNames)
    End Sub

    'Gets the email for every tournament entrant, avoiding duplicates
    Private Function GetUserEmails(ByVal TournamentEvents As List(Of String))
        'Holds each user's email
        Dim UserEmailList As New List(Of String)

        'Runs for each tournament event
        For Each TournamentEvent In TournamentEvents
            'Gets all categories for the given event
            Dim EventCategories As List(Of String) = GetEventCategories(TournamentEvent)
            'Gets the user emails from every category
            For Each Category In EventCategories
                'Opens the file in csv mode
                Dim CSVRead As New FileIO.TextFieldParser(TournamentEvent + "\" + Category) With {
                    .Delimiters = New String() {","},
                    .TextFieldType = FileIO.FieldType.Delimited
                }
                'Runs until the end of the file is reached
                While CSVRead.EndOfData = False
                    'Gets the email on the current line
                    Dim Email As String = CSVRead.ReadFields(1)
                    'Only adds to the list if the list doesn't already have the email
                    If Not UserEmailList.Contains(Email) Then
                        UserEmailList.Add(Email)
                    End If
                End While
                CSVRead.Close()
            Next
        Next

        Return UserEmailList
    End Function

    Private Sub SendRunningOrderEmail(ByVal Tournament As String, ByVal TournamentEvents As List(Of String))
        'Gets the user emails from the CSV file
        Dim UserEmails As List(Of String) = GetUserEmails(TournamentEvents)

        'Sends the drawsheet to each entrant
        For Each UserEmail In UserEmails
            'Defines the email contents and sender and recipient information
            Dim DrawSheetEmail As New MailMessage With {
                .From = New MailAddress("TournamentManagerBot@gmail.com")
            }
            DrawSheetEmail.To.Add(UserEmail)
            'Creates the email subject message
            Dim TournamentSplit As List(Of String) = Tournament.Split("\").ToList
            DrawSheetEmail.Subject = "Running Order for: " + TournamentSplit(3)
            'Attatches the generated image to the email
            Dim DrawSheetAttachment As New Attachment(Tournament.Replace("Information.csv", "RunningOrder.png"))
            DrawSheetEmail.Attachments.Add(DrawSheetAttachment)
            'Attempts to send the email
            Try
                'Using Google's SMTP Server

                Dim SMTP As New SmtpClient("smtp.gmail.com") With {
                    .Port = 587,
                    .EnableSsl = True,
                    .Credentials = New NetworkCredential("TournamentManagerBot@gmail.com", "TournamentManager1234")'Logs into the Gmail account
                }
                SMTP.Send(DrawSheetEmail)

            Catch ex As Exception
                Debug.WriteLine("Invalid Email")
            End Try
        Next
    End Sub
End Class

