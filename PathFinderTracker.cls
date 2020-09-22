VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PathFinderTracker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' This class is slaved to frmAStar.  It should not be created by ANYTHING else!
Public ThreadID As Long
Public Size As Long
Public ID As String

Private WithEvents mPathFinder As PathFinding.PathFinder
Attribute mPathFinder.VB_VarHelpID = -1

Public Property Get PathFinder() As PathFinder
  Set PathFinder = mPathFinder
End Property
Public Property Set PathFinder(ByVal vData As PathFinder)
  Set mPathFinder = vData
End Property



Private Sub mPathFinder_AnErrorOccurred(ByVal Tag As String, ByVal Err As ErrObject)
  Debug.Print "Error from multithreaded component: " & Err.Number & " " & Err.Description & " " & Err.Source
End Sub

Private Sub mPathFinder_Complete(ByVal Tag As String, ByVal Cancelled As Boolean, ByVal Success As Boolean, Path As AStarBaseClasses.Path)
  Debug.Print "Pathfinder_complete called " & Tag
  If Cancelled = False And Success = True Then
    If Tag Like "*2" Or Tag Like "*4" Or Tag Like "*6" Or Tag Like "*8" Or Tag Like "*0" Then
      Debug.Print "path2"
      Set frmAstar.Path2 = Path
    Else
      Debug.Print "path1"
      Set frmAstar.Path = Path
    End If
  End If
  frmAstar.RedrawAll
End Sub

Private Sub mPathFinder_ImStillWorkingOnIt(ByVal Tag As String, Cancel As Boolean)
  Debug.Print "it's still working on it"
End Sub

Private Sub mPathFinder_InfoMessage(ByVal Tag As String, ByVal Info As String)
    If Tag Like "*2" Or Tag Like "*4" Or Tag Like "*6" Or Tag Like "*8" Or Tag Like "*0" Then
      frmAstar.Label3.Caption = Info
    Else
      frmAstar.Label4.Caption = Info
    End If
End Sub
