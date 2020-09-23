Attribute VB_Name = "ListBoxControl"
'ListBox BAS - Assembled by Phishbowler
'
'

Public Function LISTKIllDuplicates(listBox As listBox) As Long
    Dim A%, B%
    
    blah = 0
    


    For A% = 0 To listBox.ListCount - 1


        For B% = A + 1 To listBox.ListCount - 1
            blah = blah + 1

            If listBox.List(A%) = listBox.List(B%) Then
                DuplicateCount = DuplicateCount + 1
                listBox.RemoveItem B%
                B% = B% - 1
            End If
        Next B%
    Next A%
    LISTKIllDuplicates = DuplicateCount
End Function
Sub ListKillDup(lst As listBox)
    Dim i, Duplicate
    For i = 0 To lst.ListCount - 1
        For Duplicate = 0 To lst.ListCount - 1
        If LCase(lst.List(i)) Like LCase(lst.List(Duplicate)) And i <> Duplicate Then
            lst.RemoveItem (Duplicate)
        End If
        Next Duplicate
    Next i
End Sub



Public Function DuplicatesExist(lst As listBox) As Integer
    Dim A%, B%
    
    NewkillListDuplicates = 0
    


    For A% = 0 To lst.ListCount - 1


        For B% = A + 1 To lst.ListCount - 1
            Duplicates = Duplicates + 1


            If lst.List(A%) = lst.List(B%) Then
            duplicatefound = lst.List(B%)
                Duplicate = True
                               GoTo Done:
            End If
        Next B%
    Next A%
Done:
    If Duplicate = True Then
    'Msg = MsgBox("Please change the duplicate time stamp: " & duplicatefound & "in lyrics file.", vbOKOnly, "Duplicate Time Stamp Found")
    DuplicatesExist = True
    Exit Function
    End If
    DuplicatesExist = False
End Function

Public Sub CopyList2List(CopyFrom As listBox, CopyTo As listBox)
x = 0
'Initalize CopyTo ListBox
Do
CopyTo.AddItem ""
x = x + 1
Loop Until CopyTo.ListCount = CopyFrom.ListCount

'Copy Lists

x = 0
Do
CopyTo.List(x) = CopyFrom.List(x)
x = x + 1
Loop Until x = CopyTo.ListCount


End Sub

Function LISTSearchForSelected(lst As listBox)
If lst.List(0) = "" Then
counterf = 0
GoTo last
End If
counterf = -1

Start:
counterf = counterf + 1
If lst.ListCount = counterf + 1 Then GoTo last
If lst.Selected(counterf) = True Then GoTo last
If couterf = lst.ListCount Then GoTo last
GoTo Start

last:
SearchForSelected = counterf
End Function



Function LISTSearchForIndexbyText(lst As listBox, Text As String)

If lst.ListIndex < 0 Then
LISTSearchForIndexbyText = -1
Exit Function
End If
x = 0
Do
If lst.List(x) = Text Then
LISTSearchForIndexbyText = x
Exit Do
Else
x = x + 1
End If
DoEvents

Loop While lst.ListIndex < lst.ListCount
End Function

Public Function ListDeleteSelectedItem(lst As listBox)
'This is Used as a Function so that it can Delete and
'Return the Deleted Item's Index if Needed

If lst.ListIndex < 0 Then
ListDeleteSelectedItem = -1
Exit Function
End If

x = 0
Do
If lst.Selected(x) = True Then
lst.RemoveItem (x)
ListDeleteSelectedItem = x
Exit Do
Else
x = x + 1
End If
DoEvents

Loop While lst.ListIndex < lst.ListCount


End Function

Public Function ListGetIndexText(AListBox As listBox, ListText As String) As Integer

Dim iIndex As Integer

With AListBox
 For iIndex = 0 To .ListCount - 1
   If .List(iIndex) = ListText Then
    ListGetIndexText = iIndex
    Exit Function
   End If
 Next iIndex
End With

ListGetListIndexText = -2   '  if Item isnt found
'( I didnt want to use -1 as it evaluates to True)

End Function

Sub ListDeleteItembyText(lst As listBox, Item$)
'Find's a Specific Item in List and Deletes It
On Error Resume Next
Do
NoFreeze% = DoEvents()
If LCase$(lst.List(A)) = LCase$(Item$) Then lst.RemoveItem (A)
A = 1 + A
Loop Until A >= lst.ListCount
End Sub
Function ListSelectItembyText(lst As listBox, Item$, SelectItem As Boolean)
'Find's a Specific Item in List and Selects It
'Not Case Sensitive

On Error Resume Next
If lst.ListCount > 0 Then
 Do
  NoFreeze% = DoEvents()
  If LCase$(lst.List(A)) = LCase$(Item$) Then
  If SelectItem = True Then
  lst.Selected(A) = True
  End If
  ListSelectItembyText = A
  Exit Function
  End If
  
  A = 1 + A
 Loop Until A >= lst.ListCount
End If
End Function

Public Function IsDuplicate(Item As String, lst As listBox) As Boolean
'Is string item a duplicate in list?

x = 0
If lst.ListCount > 0 Then

 Do
  If Item = lst.List(x) Then
    'This is a duplicate
  IsDuplicate = True
  Exit Function
  End If
 x = x + 1
 Loop Until x = lst.ListCount

IsDuplicate = False

Else
IsDuplicate = False

End If

End Function


Function List_IsNameListed(Lis As listBox, Name As String) As Boolean
Dim i As Integer
There = False
For i = 0 To Lis.ListCount
l$ = Lis.List(i)
If LCase(Name) = LCase(l$) Then
There = True
End If
Next i
List_IsNameListed = There
End Function

Sub ListSave(Path As String, lst As listBox)
'Ex: Call Save_ListBox("c:\windows\desktop\list.lst", list1)

    Dim Listz As Long
    On Error Resume Next

    Open Path$ For Output As #1
    For Listz& = 0 To lst.ListCount - 1
        Print #1, lst.List(Listz&)
        Next Listz&
    Close #1
End Sub

Sub ListLoad(Path As String, lst As listBox)
'Ex: Call Load_ListBox("c:\windows\desktop\list.lst", list1)

    Dim what As String
    On Error Resume Next

    Open Path$ For Input As #1
    While Not EOF(1)
        Input #1, what$
        DoEvents
        lst.AddItem what$
    Wend
    Close #1
End Sub

Public Sub ListLock(List1 As listBox, List2 As listBox, LockFirst As Boolean)
'If Lockfirst is set to True then
'First list box will be locked to the second one's movements
'Put this procedure call in a Timer
    
    Static PrevTI_TrackList
    Dim TpIndexCaptureTime
    Dim PrevTICaptureTime
      
If List1.ListCount = 0 Or List2.ListCount = 0 Then Exit Sub
      
      'Lock First to Second List
If LockFirst = True Then
      'Get the index for the first item in the visible list
      TpIndexCaptureTime = List2.TopIndex
      
      'See if the top index has changed
      If TpIndexCaptureTime <> PrevTICaptureTime Then
         'Set the top index of List1 equal to List2 so that the list boxes
         'scroll to the same relative position
         If TpIndexCaptureTime > List1.ListCount Then Exit Sub
         List1.TopIndex = TpIndexCaptureTime
         
         'Keep track of the current top index
         PrevTICaptureTime = TpIndexCaptureTime
       End If
      'Select the item in the same relative position in both list boxes
      If List2.ListIndex <> List1.ListIndex Then
       If List2.ListIndex > List1.ListCount Then Exit Sub
                 List1.ListIndex = List2.ListIndex

      End If
Else
'Lock Second to First List
      'Get the index for the first item in the visible list
      TpIndexCaptureTime = List1.TopIndex
      
      'See if the top index has changed
      If TpIndexCaptureTime <> PrevTICaptureTime Then
         'Set the top index of List1 equal to List2 so that the list boxes
         'scroll to the same relative position
         List2.TopIndex = TpIndexCaptureTime
         'Keep track of the current top index
         PrevTICaptureTime = TpIndexCaptureTime
       End If
      'Select the item in the same relative position in both list boxes
      If List1.ListIndex <> List2.ListIndex Then
         List2.ListIndex = List1.ListIndex
      End If



End If




End Sub
