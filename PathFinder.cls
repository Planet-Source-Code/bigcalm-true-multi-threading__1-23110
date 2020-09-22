VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PathFinder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Option Base 0
#Const LogThread = 1

Private Declare Function GetTickCount Lib "kernel32" () As Long

Private WithEvents mTimer As XTimers.XTimer
Attribute mTimer.VB_VarHelpID = -1
Private mMapData As Variant
Private mTag As String
Private mStartNode As AStarBaseClasses.CoOrdinate
Private mEndNode As AStarBaseClasses.CoOrdinate
Private Type MarkD
    Cost As Long
    Set As Long ' 0 = None, 1 = Open, 2 = Closed
    Direction As Long ' 0 to 3
End Type
Private mMinX As Long, mMaxX As Long, mMinY As Long, mMaxY As Long
Private mIgnoreTicks As Boolean

' This Array is used to store information about the nodes that have already been examined
Private MarkData() As MarkD


Event ImStillWorkingOnIt(ByVal Tag As String, ByRef Cancel As Boolean)
Event Complete(ByVal Tag As String, ByVal Cancelled As Boolean, ByVal Success As Boolean, Path As AStarBaseClasses.Path)
Event AnErrorOccurred(ByVal Tag As String, ByVal Err As ErrObject)
Event InfoMessage(ByVal Tag As String, ByVal Info As String)

Public Property Get ThreadID() As Long
  ThreadID = App.ThreadID
End Property

Public Property Get NumberOnThread() As Long
  NumberOnThread = glngGlobalData
End Property

' Map Data is expected to be a 2-dimensional array of MapNode
Public Function StartFindPath(ByVal MapData As Variant, ByVal StartNode As AStarBaseClasses.CoOrdinate, ByVal EndNode As AStarBaseClasses.CoOrdinate, ByVal Tag As String) As Boolean
  On Error GoTo ErrHandler
  ' First, check what we've been passed is what we're expecting:
  LogIt "Starting to find path"
  If VarType(MapData) And vbArray <> vbArray Then
    LogIt "Error in StartFindPath - You need to pass an array!"
    StartFindPath = False
    Exit Function
  End If
  
  ' Is it a two dimensional array?
'  If UBound(MapData) - LBound(MapData) + 1 <> 2 Then
'    Print #LogFileNum, "StartFindPath Quit 2"
'    StartFindPath = False
'    Exit Function
'  End If
  
  ' Is it an array of mapnode?
  
  'Ok, so let's copy to modular variables
  mMapData = MapData
  mMinX = LBound(mMapData, 1)
  mMaxX = UBound(mMapData, 1)
  mMinY = LBound(mMapData, 2)
  mMaxY = UBound(mMapData, 2)
  mTag = Tag
  Set mStartNode = StartNode
  Set mEndNode = EndNode
  LogIt "Copied passed variables ok", False
  
  ' Enable the timer.  Set to a small value so the Tick is called almost immediately.
  mTimer.Interval = 20
  mIgnoreTicks = False
  LogIt "Timer enabled", False
  mTimer.Enabled = True

  ' and return
  StartFindPath = True
  Exit Function
ErrHandler:
  StartFindPath = False
End Function

Private Sub Class_Initialize()
  glngGlobalData = glngGlobalData + 1
  #If LogThread = 1 Then
  If LogFileNum = 0 Then
    LogFileNum = FreeFile
    Open App.Path & "\pfthreadlog" & App.ThreadID & ".txt" For Output As #LogFileNum
    LogIt "Class Initialising", False
  Else
    LogIt "Installing new task on same thread", False
  End If
  #End If
  Set mTimer = New XTimers.XTimer
End Sub

Private Sub Class_Terminate()
  glngGlobalData = glngGlobalData - 1
  LogIt "Class terminating", False
  #If LogThread = 1 Then
  Close #LogFileNum
  #End If
  mTimer.Enabled = False
  Set mTimer = Nothing
  Erase mMapData
  Set mStartNode = Nothing
  Set mEndNode = Nothing
  Erase MarkData
End Sub

Private Sub mTimer_Tick()
  ' First, turn off the timer.
  mTimer.Enabled = False
  If mIgnoreTicks = False Then
    mIgnoreTicks = True
    LogIt "XTimer Tick", True
    ' Now do the task
    Call FindPath
  End If
End Sub

Private Sub FindPath()
Dim OPENSet As New Heap
' Dim CLOSEDSet As New Heap
Dim AddNode As AStarNode
Dim BestNode As AStarNode
Dim CurrentNode As HeapNode
Dim CurrNode As AStarNode
Dim i As Long
Dim j As Long
Dim k As Long
Dim InsertDirection As Long
Dim BackPath As Path
Dim Path As Path
Dim StartTime As Long
Dim EndTime As Long

    On Error GoTo ErrHandler
    LogIt "Find Path Starting"

    ' Initialize.  0-1ms
    InsertDirection = 1
    ReDim MarkData(mMaxX - mMinX + 1, mMaxY - mMinY + 1)
    For i = mMinX To mMaxX
        For j = mMinY To mMaxY
            MarkData(i, j).Cost = 0
            MarkData(i, j).Set = 0
        Next
    Next
    
    ' Add the initial node to the OPEN set. O(1)
    Set AddNode = New AStarNode
    With AddNode
        .X = mStartNode.X
        .Y = mStartNode.Y
        .gval = 0
        .hval = Cost(.X, .Y, mEndNode.X, mEndNode.Y)
        OPENSet.Add AddNode, .gval + .hval, "R" & .X & "C" & .Y
        MarkData(.X, .Y).Cost = .gval + .hval
        MarkData(.X, .Y).Set = 1
    End With
    
    Set AddNode = Nothing
    
    LogIt "Starting Main Loop"
    
    ' Main loop
    Do While OPENSet.Count > 0 ' If OPENSet.Count = 0 then there is no path
        ' Get Best Node, where "Best Node" is "The one that is most likely to lead to the goal"

        Set CurrentNode = OPENSet.GetLeftMostElement ' O(logN)
        Set BestNode = CurrentNode.Item
        
        ' Note that in this algorithm, CLOSEDSet is never used - I use the MarkData() array
        ' instead.  BUT if you make changes, then the CLOSEDSet may come in useful.
        ' Therefore, uncomment all CLOSEDSet lines if you need it
        'CLOSEDSet.Add CurrentNode.Item, CurrentNode.Value, CurrentNode.ItemKey
        
        ' check if we've reached the goal
        If BestNode.X = mEndNode.X And BestNode.Y = mEndNode.Y Then
            Exit Do ' We have reached the goal - exit and back-track to build path
        End If
        
        ' Add neighbours of current square into the OPEN set.
        ' Need to do this a different way each time to make the path "look" right
        '           0
        '         1  2
        '           3
        ' If you are allowing diagonal movement, then you would have more possible
        ' directions.
        ' e.g.    1     2     3
        '           4           6
        '           7     8    9
        ' You could always do the directions in a random order - as long as it's not the
        ' same each time (you'd end up with odd looking paths and a skewed heap if you did).
        If InsertDirection = 1 Then
            i = 0
            j = 3
        Else
            i = 3
            j = 0
        End If
        For i = i To j Step InsertDirection
            Set AddNode = New AStarNode
            ' Get neighbour...
            With AddNode
                Select Case i
                    Case 0 ' Up Neighbour
                        .Y = BestNode.Y - 1
                        .X = BestNode.X
                    Case 1 ' Left Neighbour
                        .X = BestNode.X - 1
                        .Y = BestNode.Y
                    Case 2 ' Right Neighbour
                        .X = BestNode.X + 1
                        .Y = BestNode.Y
                    Case 3 ' Bottom Neighbour
                        .Y = BestNode.Y + 1
                        .X = BestNode.X
                End Select
                
                If .X < mMinX Or .X > mMaxX Or .Y < mMinY Or .Y > mMaxY Then
                Else
                    If mMapData(.X, .Y).NodeType = 0 Then ' pMap("R" & .X & "C" & .Y).NodeType = 0 Then
                    Else
                        .gval = BestNode.gval + mMapData(.X, .Y).MoveCost 'pMap("R" & BestNode.X & "C" & BestNode.Y).MoveCost
                        .hval = Cost(.X, .Y, mEndNode.X, mEndNode.Y)
                        If MarkData(.X, .Y).Set = 0 Then
                            MarkData(.X, .Y).Cost = .gval + .hval
                            MarkData(.X, .Y).Set = 1
                            MarkData(.X, .Y).Direction = i
                            OPENSet.Add AddNode, .gval + .hval, "R" & .X & "C" & .Y ' O(logN)
                        Else
                            ' It's already in the OPEN set, so we may need to update it.
                            If MarkData(.X, .Y).Cost > .gval + .hval And MarkData(.X, .Y).Set = 1 Then
                               OPENSet.Delete "R" & .X & "C" & .Y, MarkData(.X, .Y).Cost ' O(logN)
                                MarkData(.X, .Y).Cost = .gval + .hval
                                MarkData(.X, .Y).Direction = i
                                OPENSet.Add AddNode, .gval + .hval, "R" & .X & "C" & .Y ' O(logN)
                            Else
                                ' Do nothing because we don't need to assign a higher cost to an already checked node, or we don't want to recheck a closed node
                            End If
                        End If
                    End If
                End If
            End With
            Set AddNode = Nothing
        Next
        InsertDirection = InsertDirection * -1
        If OPENSet.Delete(CurrentNode.ItemKey, CurrentNode.Value) = False Then
            LogIt "Deletion of best node failed", True
        End If
        
        MarkData(BestNode.X, BestNode.Y).Set = 2
        Set AddNode = Nothing
        Set BestNode = Nothing
        Set CurrentNode = Nothing
    Loop
    LogIt "Finished path finding main loop "
    If OPENSet.Count > 0 Then
        ' A path is possible.  6-7ms
        Set Path = New Path
        ' We're going to use the MARK array to construct the path from the end-coordinate back to the start
        ' Because we've stored the DIRECTION that the square was moved into, we can back-track from
        ' the end to the start.  Path.Add inserts the entry at the START of the collection, so we end up with
        ' the Path in the right order.
        i = mEndNode.X
        j = mEndNode.Y
        k = 1
        Do
            If k = 1 Then
              Path.Add "R" & i & "C" & j, i, j, "R" & i & "C" & j
              Path.TotalMoveCost = Path.TotalMoveCost + mMapData(i, j).MoveCost ' pMap("R" & i & "C" & j).MoveCost
            Else
              Path.Add "R" & i & "C" & j, i, j, "R" & i & "C" & j, True
              Path.TotalMoveCost = Path.TotalMoveCost + mMapData(i, j).MoveCost ' pMap("R" & i & "C" & j).MoveCost
            End If
            Select Case ReverseDirection(MarkData(i, j).Direction)
                Case 0 ' Up
                    j = j - 1
                Case 1 ' Left
                    i = i - 1
                Case 2 ' Right
                    i = i + 1
                Case 3 ' Bottom
                    j = j + 1
            End Select
            k = k + 1
        Loop Until i = mStartNode.X And j = mStartNode.Y
        ' add final node
        Path.Add "R" & i & "C" & j, i, j, "R" & i & "C" & j, True
        Path.TotalMoveCost = Path.TotalMoveCost + mMapData(i, j).MoveCost ' pMap("R" & i & "C" & j).MoveCost
    Else
        Set Path = Nothing
    End If
    ' Return the value.
    If OPENSet.Count > 0 Then
      RaiseEvent Complete(mTag, False, True, Path)
    Else
      RaiseEvent Complete(mTag, False, False, Path)
    End If
    ' Tidy up.  1-2ms
    Set OPENSet = Nothing
    'Set CLOSEDSet = Nothing
    Exit Sub
ErrHandler:
  RaiseEvent AnErrorOccurred(mTag, Err)
  RaiseEvent Complete(mTag, False, False, Nothing)
  Exit Sub
End Sub

' Heuristic.
Private Function Cost(pStartX As Long, pStartY As Long, pDestinationX As Long, pDestinationY As Long) As Long
    ' Use Manhattan distance for this demo
    Cost = Abs(pStartX - pDestinationX) + Abs(pStartY - pDestinationY)
    ' Note: If we were allowing simple diagonal movement (i.e. diagonal move costs the same as any other move)
    ' then the formula would be:
    ' Cost = Max(Abs(pStartX - pDestinationX),Abs(pStartY - pDestinationY))
    ' If straight line distance then
    ' Cost = Sqr((pStartX - pDestinationX)^2 + (pStartY - pDestinationY)^2)
End Function

' Simply flips the direction.
' 0 = Up, 1 = Left, 2 = Right, 3 = Down
Private Function ReverseDirection(pDir As Long) As Long
    Select Case pDir
        Case 0
            ReverseDirection = 3
        Case 1
            ReverseDirection = 2
        Case 2
            ReverseDirection = 1
        Case 3
            ReverseDirection = 0
    End Select
End Function

' Also raises info message events
Private Sub LogIt(Description As String, Optional RaiseInfo As Boolean = True)
  If RaiseInfo = True Then
    RaiseEvent InfoMessage(mTag, Description)
  End If
  #If LogThread = 1 Then
  Print #LogFileNum, Description
  #End If
End Sub
