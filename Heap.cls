VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Heap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"HeapNode"
Attribute VB_Ext_KEY = "Member0" ,"HeapNode"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

' Code by Jonathan Daniel 01/11/00
' This project contains my implementation of a HEAP (tree data structure).  In computer science, C programming or
' databases these are found quite often, but I've never been able to find an easy way to implement
' them in Visual Basic.
' Could be used to implement a Heapsort algorithm.
' Comparing data types (all values are given as averages). N = number of elements
'                                   Search                      Insert                                  Delete
' Sorted Array                 O(logN) (?)           Search time + O(N/2)       Search time + O(N/2)
' Ordered Linked List       O(N/2)                 Search time + O(1)          Search time + O(1)
' Heap                            O(logN)               Search time + O(1)          Search time + O(1)
'

Public TopNode As HeapNode
Public Count As Long
Public TreeHeight As Long
Public LeftOrRightDeletion As Long ' Variable is used to vary whether we pick the left or right item to "bring up" the tree during a deletion.
                                                    ' If we were to constantly pick "left" we might end up with a skewed heap.

' Class functions
Private Sub Class_Initialize()
    Set TopNode = Nothing
    TreeHeight = 0
    Count = 0
    LeftOrRightDeletion = 0
End Sub

Private Sub Class_Terminate()
    DestroyHeap
End Sub

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = TopNode.Item
End Property

Public Property Get Item(vntIndexKey As Variant) As HeapNode
Attribute Item.VB_UserMemId = 0
    Set Item = TopNode.Item
End Property

Public Function Delete(ByVal ItemKey As String, ByVal Value As Single) As Boolean
Dim NodeToDelete As HeapNode
Dim ParentNode As HeapNode
Dim FoundAtDepth As Long
Dim ChildNode As HeapNode
Dim SeperateTree As HeapNode
Dim ReinsertNode As HeapNode
Dim LeftOrRightChild As Long
Dim DepthFound As Long

    Set NodeToDelete = FindInHeap(Value, ItemKey)
    If NodeToDelete Is Nothing Then
        Delete = False
        Exit Function
    End If
    Delete = True
    If NodeToDelete.LeftChild Is Nothing Then
        If NodeToDelete.RightChild Is Nothing Then
            If NodeToDelete.ParentNode Is Nothing Then
                ' Deleting top level element
                Set NodeToDelete = Nothing
                Set TopNode = Nothing
            Else
                ' Easy peasy, we're deleting a bottom level element.
                Set ParentNode = NodeToDelete.ParentNode
                If NodeToDelete.Value <= ParentNode.Value Then ' Left child
                    Set ParentNode.LeftChild = Nothing
                Else
                    Set ParentNode.RightChild = Nothing
                End If
                Set NodeToDelete = Nothing
            End If
        Else
            ' We're deleting a node with no left children, but with right children.
            '   5                                                   8
            '  /  \                                                /   \
            ' 4    6                   OR                    6     10
            '         \                                            \
            '           7                                           7
            ' Deleting Node 6.  So, we need to attach 5 to 7, or deleting node 6, so we need to attach 7 to 8
            If NodeToDelete.ParentNode Is Nothing Then
                ' I've got to rejig the top node.
                Set TopNode = NodeToDelete.RightChild
                Set NodeToDelete.RightChild.ParentNode = Nothing
                Set NodeToDelete = Nothing
            Else
                Set ParentNode = NodeToDelete.ParentNode
                Set NodeToDelete.RightChild.ParentNode = ParentNode
                If NodeToDelete.Value <= ParentNode.Value Then ' Left child of parent node
                    Set ParentNode.LeftChild = NodeToDelete.RightChild
                Else    ' Right child of parent node
                    Set ParentNode.RightChild = NodeToDelete.RightChild
                End If
                Set NodeToDelete = Nothing ' Destroys node.
                ' We now have: from the previous examples:
                '    5                                              8
                '  /   \                  OR                    /    \
                ' 4    7                                       7      10
            End If
        End If
    Else
        If NodeToDelete.RightChild Is Nothing Then
                ' We're deleting a node with no right children, but with left children.
            If NodeToDelete.ParentNode Is Nothing Then
                ' I've got to rejig the top node.
                Set TopNode = NodeToDelete.LeftChild
                Set NodeToDelete.LeftChild.ParentNode = Nothing
                Set NodeToDelete = Nothing
            Else
                Set ParentNode = NodeToDelete.ParentNode
                Set NodeToDelete.LeftChild.ParentNode = ParentNode
                If NodeToDelete.Value <= ParentNode.Value Then ' Left child of parent node
                    Set ParentNode.LeftChild = NodeToDelete.LeftChild
                Else
                    Set ParentNode.RightChild = NodeToDelete.LeftChild
                End If
                Set NodeToDelete.LeftChild.ParentNode = ParentNode
                Set NodeToDelete = Nothing ' Destroys node
            End If
        Else
            ' We are deleting an element with both Left and Right Children nodes.
            ' If we delete the Left hand node:
            ' Step 1: Store all right hand children off the left child in a seperate tree
            ' Step 2: Set the Left hand children's Right Child node to be the deleting item's
            '             right hand node.
            ' Step 3: Delete the now unattached top node.
            ' Step 4: Find place for the seperate tree
            ' Step 5: Re-insert
            ' e.g For deleting item 5:
            '       Initial                  Step 1                                   Steps 2&3                       Steps 4&5
            '          5                           5                                         2                                      2
            '        /     \                      /    \                                     /   \                                   /   \
            '      2        8                 2       8       Tree 2                 1     8   Tree 2                   1     8
            '    /   \     /   \              /  \     /  \         \                            /  \       \                            /  \
            '  1     4  7     9          1    4  7    9        4                         7    9      4                         7   9
            '                                                                                                                          /
            '                                                                                                                        4
            ' The reverse is used for right-hand child.
            If LeftOrRightDeletion = 0 Then ' Left node becomes this node
                ' Step 1 - Store
                Set ChildNode = NodeToDelete.LeftChild
                Set SeperateTree = ChildNode.RightChild
                If Not (SeperateTree Is Nothing) Then
                    Set SeperateTree.ParentNode = Nothing
                End If
                ' Step 2 - Move
                Set ChildNode.RightChild = NodeToDelete.RightChild
                Set ChildNode.RightChild.ParentNode = ChildNode
                If NodeToDelete.ParentNode Is Nothing Then
                    Set ChildNode.ParentNode = Nothing
                    Set TopNode = ChildNode
                Else
                    Set ChildNode.ParentNode = NodeToDelete.ParentNode
                    If NodeToDelete.Value <= NodeToDelete.ParentNode.Value Then
                        Set NodeToDelete.ParentNode.LeftChild = ChildNode
                    Else
                        Set NodeToDelete.ParentNode.RightChild = ChildNode
                    End If
                End If
                ' Step 3 - Delete
                Set NodeToDelete = Nothing
                ' Step 4 - Reinsert
                If Not (SeperateTree Is Nothing) Then
                    Set ReinsertNode = FindFreeSpace(SeperateTree.Value, LeftOrRightChild, DepthFound)
                    Set SeperateTree.ParentNode = ReinsertNode
                    If LeftOrRightChild = 0 Then
                        Set ReinsertNode.LeftChild = SeperateTree
                    Else
                        Set ReinsertNode.RightChild = SeperateTree
                    End If
                End If
            Else ' Right node becomes this node
                ' Step 1 - Store
                Set ChildNode = NodeToDelete.RightChild
                Set SeperateTree = ChildNode.LeftChild
                If Not (SeperateTree Is Nothing) Then
                    Set SeperateTree.ParentNode = Nothing
                End If
                ' Step 2 - Move
                Set ChildNode.LeftChild = NodeToDelete.LeftChild
                Set ChildNode.LeftChild.ParentNode = ChildNode
                If NodeToDelete.ParentNode Is Nothing Then
                    Set ChildNode.ParentNode = Nothing
                    Set TopNode = ChildNode
                Else
                    Set ChildNode.ParentNode = NodeToDelete.ParentNode
                    If NodeToDelete.Value <= NodeToDelete.ParentNode.Value Then
                        Set NodeToDelete.ParentNode.LeftChild = ChildNode
                    Else
                        Set NodeToDelete.ParentNode.RightChild = ChildNode
                    End If
                End If
                ' Step 3 - Delete
                Set NodeToDelete = Nothing
                ' Step 4 - Reinsert
                If Not (SeperateTree Is Nothing) Then
                    Set ReinsertNode = FindFreeSpace(SeperateTree.Value, LeftOrRightChild, DepthFound)
                    Set SeperateTree.ParentNode = ReinsertNode
                    If LeftOrRightChild = 0 Then
                        Set ReinsertNode.LeftChild = SeperateTree
                    Else
                        Set ReinsertNode.RightChild = SeperateTree
                    End If
                End If
            End If ' Left or Right move
            LeftOrRightDeletion = LeftOrRightDeletion Xor 1
        End If ' Both left and right nodes filled
    End If ' Left child is nothing
    Count = Count - 1
End Function

Public Function Add(Obj As Object, Value As Single, ItemKey As String) As HeapNode
Dim objHeapNode As New HeapNode
Dim objParentNode As HeapNode
Dim LeftOrRight As Long
Dim InsertionDepth As Long
    
    ' Initialize the heap node that we're going to insert
    With objHeapNode
        .ItemKey = ItemKey
        Set .Item = Obj
        .Value = Value
        Set .LeftChild = Nothing
        Set .RightChild = Nothing
        Set .ParentNode = Nothing
    End With

    ' If this is the first element being added....
    If TopNode Is Nothing Then
        Set TopNode = objHeapNode
        TreeHeight = 1
    Else
        Set objParentNode = FindFreeSpace(Value, LeftOrRight, InsertionDepth)
        With objHeapNode
            Set .ParentNode = objParentNode
        End With
        With objParentNode
            If LeftOrRight = 0 Then
                Set objParentNode.LeftChild = objHeapNode
            Else
                Set objParentNode.RightChild = objHeapNode
            End If
        End With
        If InsertionDepth > TreeHeight Then
            TreeHeight = InsertionDepth
        End If
    End If
    Count = Count + 1
End Function

Public Function FindInHeap(Value As Single, ItemKey As String) As HeapNode
Dim CurrentNode As HeapNode
Dim FoundIt As Boolean

    Set CurrentNode = TopNode
    FoundIt = False
    Do While FoundIt = False
        With CurrentNode
            If Value <= .Value Then
                If Value = .Value And ItemKey = .ItemKey Then
                    ' gotcha
                    Set FindInHeap = CurrentNode
                    FoundIt = True
                Else
                    If .LeftChild Is Nothing Then ' Cant find it.
                        Set FindInHeap = Nothing
                        FoundIt = True
                    Else
                        Set CurrentNode = .LeftChild
                    End If
                End If
            Else
                If .RightChild Is Nothing Then ' cant find it
                    Set FindInHeap = Nothing
                    FoundIt = True
                Else
                    Set CurrentNode = .RightChild
                End If
            End If
        End With
    Loop
End Function
Public Function GetLeftMostElement() As HeapNode
Dim CurrentNode As HeapNode
    Set CurrentNode = TopNode
    If TopNode Is Nothing Then
        Set GetLeftMostElement = Nothing
        Exit Function
    End If
    Do While Not (CurrentNode.LeftChild Is Nothing)
        Set CurrentNode = CurrentNode.LeftChild
    Loop
    Set GetLeftMostElement = CurrentNode
End Function

Public Function FindFreeSpace(ByRef Value As Single, ByRef LeftOrRight As Long, ByRef DepthFound As Long) As HeapNode
Dim CurrentNode As HeapNode
Dim FoundFree As Boolean
Dim SearchDepth As Long
    ' Note that Identical values will always go down the left hand side of the tree
    
    SearchDepth = 1
    Set CurrentNode = TopNode
    FoundFree = False
    Do While FoundFree = False
        With CurrentNode
            If Value <= .Value Then
                If .LeftChild Is Nothing Then
                    FoundFree = True
                    LeftOrRight = 0
                Else
                    Set CurrentNode = .LeftChild
                End If
            Else
                If .RightChild Is Nothing Then
                    FoundFree = True
                    LeftOrRight = 1
                Else
                    Set CurrentNode = .RightChild
                End If
            End If
        End With
        SearchDepth = SearchDepth + 1
    Loop
    Set FindFreeSpace = CurrentNode
    DepthFound = SearchDepth
End Function

Public Function DestroyHeap() As Boolean
Dim CurrentNode As HeapNode
Dim ParentNode As HeapNode
    Set CurrentNode = TopNode
    Do While Count > 0
        With CurrentNode
            If .LeftChild Is Nothing Then
                If .RightChild Is Nothing Then
                    If .ParentNode Is Nothing Then ' Top level element
                        Set CurrentNode = Nothing
                        Set TopNode = Nothing
                    Else
                        Set ParentNode = .ParentNode
                        Set .ParentNode = Nothing
                        If .Value <= ParentNode.Value Then ' Left hand child
                            Set ParentNode.LeftChild = Nothing
                        Else ' Right hand child
                            Set ParentNode.RightChild = Nothing
                        End If
                        Set CurrentNode = Nothing ' just make sure it kills the memory associated with the object
                        Set CurrentNode = ParentNode
                    End If
                    Count = Count - 1
                Else
                    ' Right child is not nothing, therefore progress to that element and check again
                    Set CurrentNode = .RightChild
                End If
            Else
                ' Left child is not nothing, so progress to that node and check again
                Set CurrentNode = .LeftChild
            End If
        End With
    Loop
End Function

'Public Function TrueCount() As Long
'Dim CurrentNode As HeapNode
'    Set CurrentNode = TopNode
'    If CurrentNode Is Nothing Then
'        TrueCount = 0
'        Exit Function
'    End If
'    TrueCount = CountChildren(CurrentNode)
'End Function
'
'Private Function CountChildren(pHN As HeapNode) As Long
'Dim lCount As Long
'    lCount = 0
'    If Not (pHN.LeftChild Is Nothing) Then
'        lCount = lCount + CountChildren(pHN.LeftChild)
'    End If
'    If Not (pHN.RightChild Is Nothing) Then
'        lCount = lCount + CountChildren(pHN.RightChild)
'    End If
'    lCount = lCount + 1
'    CountChildren = lCount
'End Function
'
