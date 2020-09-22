VERSION 5.00
Begin VB.Form frmAstar 
   Caption         =   "A* Demo"
   ClientHeight    =   7680
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9390
   LinkTopic       =   "Form1"
   ScaleHeight     =   512
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   626
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "Clear Map"
      Height          =   495
      Left            =   7920
      TabIndex        =   10
      Top             =   1440
      Width           =   1335
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   7920
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   1920
      Width           =   1335
   End
   Begin VB.PictureBox Picture3 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   660
      Left            =   1560
      Picture         =   "Helga.frx":0000
      ScaleHeight     =   40
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   40
      TabIndex        =   8
      Top             =   360
      Visible         =   0   'False
      Width           =   660
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   360
      Left            =   1080
      Picture         =   "Helga.frx":1302
      ScaleHeight     =   20
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   20
      TabIndex        =   5
      Top             =   360
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Generate Map"
      Height          =   495
      Left            =   7920
      TabIndex        =   4
      Top             =   960
      Width           =   1335
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   360
      Left            =   360
      Picture         =   "Helga.frx":17F4
      ScaleHeight     =   20
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   20
      TabIndex        =   3
      Top             =   360
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Redraw"
      Height          =   495
      Left            =   7920
      TabIndex        =   2
      Top             =   480
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Find Paths"
      Height          =   495
      Left            =   7920
      TabIndex        =   0
      Top             =   0
      Width           =   1335
   End
   Begin VB.PictureBox picWorking 
      AutoRedraw      =   -1  'True
      Height          =   6615
      Left            =   0
      ScaleHeight     =   437
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   509
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   7695
   End
   Begin VB.Label Label4 
      Caption         =   "Label4"
      Height          =   255
      Left            =   3720
      TabIndex        =   12
      Top             =   7440
      Width           =   3855
   End
   Begin VB.Label Label3 
      Caption         =   "Label3"
      Height          =   255
      Left            =   0
      TabIndex        =   11
      Top             =   7440
      Width           =   3495
   End
   Begin VB.Label Label2 
      Height          =   855
      Left            =   7920
      TabIndex        =   7
      Top             =   2280
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   $"Helga.frx":1CE6
      Height          =   615
      Left            =   0
      TabIndex        =   6
      Top             =   6720
      Width           =   6135
   End
End
Attribute VB_Name = "frmAstar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Code by Jonathan Daniel. Version 1.3 - Multithreaded
' Description:  Demo of the A* pathing algorithm
' No copyright - feel free to copy this code and use it in any way you feel like.
' Any comments, improvements, and bug-fixes to bigcalm@hotmail.com

' Please please please, if you have no idea what A* is, or what it is used for, then
' please read some of the following resources, before mailing me:
' http://theory.stanford.edu/~amitp/GameProgramming/    - Amit Patel's fantastic notes
'       on A* and other games programming stuff.  Truly superb resource.
' http://www.gameai.com     -   Another good resouce for Artificial Intelligence, including some
'               notes on A*
' http://www.gamedev.net    -   If you want your brain fried by Maths, check out these very useful pages (novice programmers should stay away)
' http://www.ccg.leeds.ac.uk/james/aStar/   - a java representation of the A* algorithm
' http://www.utm.edu/cgi-bin/caldwell/tutor/departments/math/graph/intro - If you are not intending to
'       use a tile based system, but a vector based system, this page introduces graph theory (you will
'       need to reduce your vector objects to a connected graph for A* to work out a path).

' This Project implements the A* algorithm for finding a path through terrain, and the
' algorithm is used in games to help find a path from one place to another on a map.
' Different types of A* are used in other types of game - for example, chess game
' opponents use an A* variant called ID-A* (Iterative Deepening).
' There are many possible implementations of A*, of this is only one.
' If you want to include this into a game environment you will have to make several
' modifications.
' Namely:  1) Use your own MAP class.  (If you used an array to hold the map rather than a collection
'                   the A* algorithm would be 30-50% faster).
'               2) If your game is "interactive" instead of "turn taking" you will need to be
'                   able to interrupt path finding, to draw graphics, do AI, etc.  I suggest
'                   giving A* a maximum period of time to work on the problem, start
'                   any unit walking in the general direction (or, get them to scratch their heads),
'                   and give A* a few miliseconds per frame to work out paths.  The reason
'                   behind this is that A* IS COMPUTATIONALLY EXPENSIVE, even though it will always
'                   find a path if there is one.  If there is no path, (e.g. the end is blocked off), then
'                   then A* will search every available square before giving up).  You could even
'                   give A* a maximum search time, in which case, your unit just stays where they
'                   are (and assumes no path if not found in a certain amount of time).
'                   You can also "tune" the A* algorithm by modifying the heuristic estimate of the
'                   distance - by giving an under-estimate, it will find paths quicker, but
'                   there might actually be *better* (less movement needed) paths.
'               3) Adjust the A* to suit your game - if you allow diagonal movement, want to allow
'                  path splicing, etc.  Check out http://theory.stanford.edu/~amitp/GameProgramming/
'                   for possible ways to adapt this A* algorithm to your game (also covers Path Splicing).
'               4) Improve the A* algorithm (use Beam A*, or change the heuristic) - my
'                   algorithm is probably a little too general.  Plus changing the HeapNode Class to having Early-Binding would help
' I would also recommend writing A* in C or C++, as it will be much faster because of
' C's good implementation of heaps ( I've only written this in VB 'cos my C compiler's broken).

' This implementation of A* can cope with different types of terrain (i.e. MoveCost of Map is taken into account), so
' this implementation can find quicker paths if there are Roads, or navigate around Mountains, etc.
' For Route finding software (that finds a road-route from one location to another), A* is also
' used, but instead of considering all roads, it will navigate to a B road, then an A road, then to a motorway,
' and then reverse this if necessary, using several levels of A* pathing to speed up the search.
' This dividing of A* is also used on some "World" map type games, where continents and islands are "precalculated"
' before telling A* to search for a path.


Public Path As Path
Public Path2 As Path
Private CoStart As CoOrdinate
Private CoEnd As CoOrdinate
Private CoStart2 As CoOrdinate
Private CoEnd2 As CoOrdinate
Private PathFinders(0 To 1) As PathFinding.PathFinder
Public Trackers As New Collection

Private Function NewTracker(ByVal ThreadID As Long, _
        ByVal Size As Long, ByVal Tag As String) As PathFinderTracker
        
    Dim pft As New PathFinderTracker
    '
    ' Cache the thread ID of the Coffee object
    '   the tracker will be keeping track of.
    pft.ThreadID = ThreadID
    '
    ' Set the size of the task assigned to the
    '   Coffee object the tracker will track.
    pft.Size = Size
    '
    ' Give the tracker a unique key for the
    '   collection.
    pft.ID = Tag
    '
    ' Put the new tracker into a collection.
    Trackers.Add pft, pft.ID
    '
    ' Return a reference to the new tracker.
    Set NewTracker = pft
End Function

' Provide unique keys for Trackers.
'
Private Function NewKey() As String
    Static lngLastKey As Long
    lngLastKey = lngLastKey + 1
    NewKey = "K" & lngLastKey
End Function

' Find the paths
Private Sub Command1_Click()
Dim StartTime As Long
Dim EndTime As Long
Dim CoOrd As CoOrdinate
Dim Tag As String
Dim Tracker As PathFinderTracker
    
    Set Path = Nothing
    Set Path2 = Nothing
    
    Tag = NewKey
    If PathFinders(0) Is Nothing Then
      Set PathFinders(0) = New PathFinder
      Set Tracker = NewTracker(PathFinders(0).ThreadID, 10, Tag)
      Set Tracker.PathFinder = PathFinders(0)
    End If
    PathFinders(0).StartFindPath MapArr, CoStart, CoEnd, Tag

    
    Tag = NewKey
    If PathFinders(1) Is Nothing Then
      Set PathFinders(1) = New PathFinder
      Set Tracker = NewTracker(PathFinders(1).ThreadID, 10, Tag)
      Set Tracker.PathFinder = PathFinders(1)
    End If
    PathFinders(1).StartFindPath MapArr, CoStart2, CoEnd2, Tag
End Sub

' Draw map
Private Sub Command2_Click()
  RedrawAll
End Sub

Public Sub RedrawAll()
  DrawBoard
  FlipBoard
End Sub

' Generate random map
Private Sub Command3_Click()
Dim i As Long
Dim j As Long
Dim k As Long
Dim MoveCost As Long

    Set Path = Nothing
    Set Path2 = Nothing
    Randomize
    For i = 0 To MaxX
        For j = 0 To MaxY
            If MapArr(i, j) Is Nothing Then
              Set MapArr(i, j) = New AStarBaseClasses.MapNode
            End If
            k = Rnd * 1.5
            If k = 1 Then
                If Rnd > 0.5 Then
                    k = 5 + Rnd * 2
                    Select Case k
                        Case 5
                            MoveCost = 10
                        Case 6
                            MoveCost = 0
                        Case 7
                            MoveCost = 5
                        Case Else
                            MoveCost = 1
                    End Select
                Else
                    MoveCost = 1
                End If
            Else
                MoveCost = 1
            End If
            MapArr(i, j).MoveCost = MoveCost
            MapArr(i, j).NodeType = k
        Next
    Next

    Set CoStart = Nothing
    Set CoStart = New CoOrdinate
    Set CoStart2 = Nothing
    Set CoStart2 = New CoOrdinate
    CoStart.X = 0
    CoStart.Y = 0
    MapArr(0, 0).NodeType = 1
    MapArr(MaxX, MaxY).NodeType = 1
    CoStart2.X = 0
    CoStart2.Y = MaxY
    Set CoEnd = Nothing
    Set CoEnd = New CoOrdinate
    Set CoEnd2 = Nothing
    Set CoEnd2 = New CoOrdinate
    CoEnd.X = MaxX
    CoEnd.Y = MaxY
    CoEnd2.X = MaxX
    CoEnd2.Y = 0
    MapArr(CoStart2.X, CoStart2.Y).NodeType = 1
    MapArr(CoEnd2.X, CoEnd2.Y).NodeType = 1

    Command2_Click
End Sub

Private Sub Command4_Click()
Dim i As Long
Dim j As Long
Dim k As Long
Dim MoveCost As Long

    Set Path = Nothing
    Set Path2 = Nothing
    For i = 0 To MaxX
        For j = 0 To MaxY
            If MapArr(i, j) Is Nothing Then
              Set MapArr(i, j) = New AStarBaseClasses.MapNode
            End If
            MapArr(i, j).MoveCost = 1
            MapArr(i, j).NodeType = 1
        Next
    Next
    MapArr(0, 0).NodeType = 1
    MapArr(MaxX, MaxY).NodeType = 1
    Set CoStart = Nothing
    Set CoStart = New CoOrdinate
    Set CoStart2 = Nothing
    Set CoStart2 = New CoOrdinate
    CoStart.X = 0
    CoStart.Y = 0
    CoStart2.X = 0
    CoStart2.Y = MaxY
    Set CoEnd = Nothing
    Set CoEnd = New CoOrdinate
    Set CoEnd2 = Nothing
    Set CoEnd2 = New CoOrdinate
    CoEnd.X = MaxX
    CoEnd.Y = MaxY
    CoEnd2.X = MaxX
    CoEnd2.Y = 0
    Command2_Click
End Sub

Private Sub Form_Load()
Dim i As Long
Dim j As Long
  
  picWorking.ScaleMode = 3
  picWorking.Width = (MaxY + 1) * 20
  picWorking.Height = (MaxX + 1) * 20
  
    Command3_Click
    Combo1.AddItem "Normal"
    Combo1.AddItem "Mountain"
    Combo1.AddItem "Road"
    Combo1.AddItem "Forest"
    Combo1.AddItem "Impassable"
    Combo1.ListIndex = 0
End Sub

' Erm, I've realised that somewhere I've reversed co-ordinates
' So rows are columns and vice-versa when it comes to displaying the information
' Oops.
Private Sub DrawBoard()
Dim MN As MapNode
Dim CoOrd As CoOrdinate
Dim LastCoOrd As CoOrdinate
Dim i As Long
Dim j As Long
Dim X As IEnumVARIANT

    For i = 0 To MaxX
        For j = 0 To MaxY
        Select Case MapArr(i, j).NodeType
            Case 0
                BitBlt picWorking.hDC, j * 20, i * 20, 20, 20, picWorking.hDC, 0, 0, BLACKNESS
            Case 1
                BitBlt picWorking.hDC, j * 20, i * 20, 20, 20, picWorking.hDC, 0, 0, WHITENESS
            Case 2
                BitBlt picWorking.hDC, j * 20, i * 20, 20, 20, Picture1.hDC, 0, 0, vbSrcCopy
            Case 3
                BitBlt picWorking.hDC, j * 20, i * 20, 20, 20, Picture1.hDC, 0, 0, WHITENESS
                BitBlt picWorking.hDC, j * 20, i * 20, 20, 20, Picture1.hDC, 0, 0, vbSrcInvert
            Case 4
                BitBlt picWorking.hDC, j * 20, i * 20, 20, 20, Picture2.hDC, 0, 0, vbSrcCopy
            Case 5
                BitBlt picWorking.hDC, j * 20, i * 20, 20, 20, Picture3.hDC, 0, 0, vbSrcCopy
            Case 6
                BitBlt picWorking.hDC, j * 20, i * 20, 20, 20, Picture3.hDC, 20, 0, vbSrcCopy
            Case 7
                BitBlt picWorking.hDC, j * 20, i * 20, 20, 20, Picture3.hDC, 20, 20, vbSrcCopy
            Case 8
                BitBlt picWorking.hDC, j * 20, i * 20, 20, 20, Picture3.hDC, 0, 20, vbSrcCopy
        End Select
        Next
    Next
    If Not (Path Is Nothing) Then
        Set LastCoOrd = Nothing
        For Each CoOrd In Path
            If Not (LastCoOrd Is Nothing) Then
                picWorking.Line (LastCoOrd.Y * 20 + 10, LastCoOrd.X * 20 + 10)-(CoOrd.Y * 20 + 10, CoOrd.X * 20 + 10)
            End If
            Set LastCoOrd = CoOrd
        Next
    End If
    If Not (Path2 Is Nothing) Then
      Set LastCoOrd = Nothing
      For Each CoOrd In Path2
        If Not (LastCoOrd Is Nothing) Then
          picWorking.Line (LastCoOrd.Y * 20 + 10, LastCoOrd.X * 20 + 10)-(CoOrd.Y * 20 + 10, CoOrd.X * 20 + 10), vbRed
        End If
        Set LastCoOrd = CoOrd
      Next
    End If
End Sub

Private Sub FlipBoard()
    BitBlt Me.hDC, 0, 0, picWorking.ScaleWidth, picWorking.ScaleHeight, picWorking.hDC, 0, 0, vbSrcCopy
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim MLX As Long
Dim MLY As Long
    MLX = Y \ 20
    MLY = X \ 20
    If MLX < 0 Or MLX > MaxX Then
        Exit Sub
    End If
    If MLY < 0 Or MLY > MaxY Then
        Exit Sub
    End If
    Select Case Combo1.ListIndex
        Case 0
            MapArr(MLX, MLY).NodeType = 1
            MapArr(MLX, MLY).MoveCost = 1
        Case 1
            MapArr(MLX, MLY).NodeType = 5
            MapArr(MLX, MLY).MoveCost = 10
        Case 2
            MapArr(MLX, MLY).NodeType = 6
            MapArr(MLX, MLY).MoveCost = 0
        Case 3
            MapArr(MLX, MLY).NodeType = 7
            MapArr(MLX, MLY).MoveCost = 5
        Case 4
            MapArr(MLX, MLY).NodeType = 0
            MapArr(MLX, MLY).MoveCost = 10
    End Select
    Command2_Click
End Sub
