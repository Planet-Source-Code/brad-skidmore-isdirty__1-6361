<div align="center">

## IsDirty


</div>

### Description

(Update 3/2/2000 I forgot to add the Enum To be used with the MyFormValuesOnLoad() Array But its in there now:)

This code will aide in determining if Data has changed on a form. It serves many purposes. 1. To tell wether or not it is approriate to prompt user if they want to save changes they have made. 2. if the user wants to Reset or Undo changes they have made to a single text box, checkbox, or combobox Or all Controls at once(Works for control arrays as well).
 
### More Info
 
Need to place a Variant array on each form you want this functionality (To store all the values of those three different control types whith the data after you have loaded the form)

The Function isDirty will return true if any Value on your form is different from the original values on Load of the form. As well the Isdirty can change a single Control to its original value or All Controls.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Brad Skidmore](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/brad-skidmore.md)
**Level**          |Intermediate
**User Rating**    |4.5 (9 globes from 2 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Custom Controls/ Forms/  Menus](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/custom-controls-forms-menus__1-4.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/brad-skidmore-isdirty__1-6361/archive/master.zip)





### Source Code

```
'This Code Is placed in a Common Module so All forms Can Access it
'This Enum used to Navigate the MyFormValuesOnLoad Array (-1,0,1,2)
Public Enum ValueType
  NotControlArray = -1
  MyName
  MyTextOrValue
  MyIndex
End Enum
'These are Constants for Use in calling IsDirty
Public Const RESET_VALUES As Boolean = True
Public Const RESET_ACTIVE_CONTROL As Boolean = True
''''''''''''''''''''''''''''''''''''''''''''''
'This Code Is placed in a Common Module so All forms Can Access it
Public Sub FormatData(MyForm As Form, MyFormValuesOnLoad As Variant)
 'BGS 8/10/1999
 'A. formats data in all controls for MyForm
 'depending upon the control type and what its tag property says
 'B. Then it places all the control names and their values into
 'a dynamic two dimensional variant array MyFormValuesOnLoad to be used later.
 'The IsDirty boolean function will use this variant array to tell whether
 'changes were made, as well as reset the values on the form if the user
 'desires to do so.
 On Error GoTo EH
 Dim MyControl As Control
 Dim MyControlCount As Integer
 MyControlCount = 0
 'A. formats data in all controls for MyForm
 'depending upon the control type and what its tag property says
 Screen.MousePointer = vbHourglass
 For Each MyControl In MyForm.Controls
  'Put data formating code here
  '
  '
  '
  '
  'End Format Code
  If TypeOf MyControl Is TextBox Or TypeOf MyControl Is CheckBox Or TypeOf MyControl Is ComboBox Then
   MyControlCount = MyControlCount + 1
  End If
 Next
  'B. Then it places all the control names and their values into
 'a dynamic two dimensional variant array MyFormValuesOnLoad to be used later.
 'The IsDirty boolean function will use this variant array to tell whether
 'changes were made, as well as reset the values on the form if the user
 'desires to do so.
 ReDim MyFormValuesOnLoad(MyName To MyIndex, 1 To MyControlCount)
 MyControlCount = 0
 For Each MyControl In MyForm.Controls
  If TypeOf MyControl Is TextBox Or TypeOf MyControl Is CheckBox Or TypeOf MyControl Is ComboBox Then
   MyControlCount = MyControlCount + 1
   MyFormValuesOnLoad(MyName, MyControlCount) = MyControl.Name
   If TypeOf MyControl Is TextBox Then
    MyFormValuesOnLoad(MyTextOrValue, MyControlCount) = MyControl.Text
   ElseIf TypeOf MyControl Is CheckBox Then
    MyFormValuesOnLoad(MyTextOrValue, MyControlCount) = MyControl.Value
   ElseIf TypeOf MyControl Is ComboBox Then
    MyFormValuesOnLoad(MyTextOrValue, MyControlCount) = MyControl.ListIndex
   End If
   If isControlArray(MyForm, MyControl) Then
    MyFormValuesOnLoad(MyIndex, MyControlCount) = MyControl.Index
   Else
    MyFormValuesOnLoad(MyIndex, MyControlCount) = NotControlArray
   End If
  End If
 Next
 Screen.MousePointer = vbDefault
 Exit Sub
EH:
 Screen.MousePointer = vbDefault
 MsgBox Err.Description & " in Form " & MyForm.Name, , "FormatData"
End Sub
''''''''''''''''''''''''''''''''''
'This Code Is placed in a Common Module so All forms Can Access it
Public Function isControlArray(MyForm As Form, MyControl As Control) As Boolean
 'BGS 8/1/1999 Added this function to determin if a Control is part of
 'a control array or not. I had to do this because VB does not have a
 'function that figures this out IsArray does not work on Control Arrays
 On Error GoTo EH
 Dim MyCount As Integer
 Dim CheckMyControl As Control
 For Each CheckMyControl In MyForm.Controls
  If CheckMyControl.Name = MyControl.Name Then
   MyCount = MyCount + 1
  End If
 Next
 isControlArray = MyCount - 1
 Exit Function
EH:
 MsgBox Err.Description & "in Form " & MyForm.Name, , "isControlArray"
End Function
''''''''''''''''''''''''''''''''''
'This Code Is placed in a Common Module so All forms Can Access it
Public Function IsDirty(MyForm As Form, MyFormValuesOnLoad As Variant, Optional Reset As Boolean, Optional ResetActiveControl As Boolean, Optional MyActiveControl As Control) As Boolean
 'BGS 8/8/1999 IsDirty for Forms with TextBoxes, CheckBoxes, and ComboBoxes
 'Checks all the Controls on Myform and compares their values to what is in
 'MyFormValuesOnLoad Variant Array.
 'First the Function checks the type of each Control, if they are a TexBox CheckBox
 'or ComboBox then it will continue on. Continuing, it will check to see if the
 'Control in question is a Control array or not. IF it is then the function will
 'compare each Name in the MyFormValuesOnLoad Variant array, When then name matches
 'the one in the Array, then it will compare the Index. When both name and the Index
 'match , then it will check the TypeOf of the Control in Question. If it is a TexBox
 'then the function will compare the .Text to the MyTextOrValue in the Array. If it matches then It
 'is "Not Dirty" so the Boolean variable bIsDirty remains False. (***Note if the Boolean Variable
 'Reset is set to True Then All Controls will be set back to their previous value stored in the Array.
 'Or if ResetActiveControl is Set to True, Then ONLY the Control which currently has Focus would be reset to
 'the previous value stored in the Array. ***) The function does the exact same thing for
 'the CheckBox and ComboBox controls but uses the .Value and .ListIndex instead of the .Text .
 'IF the Control in question is not a control array then the function does the exact same
 'thing as above but leaves out checking to make sure the index matches the Array since it
 'does not have that property.
 On Error GoTo EH
 Dim MyControl As Control
 Dim MyControlCount As Integer
 Dim MyActCtrlName As String
 Dim MyActCtrlIndex As Integer
 Dim bIsDirty As Boolean
 Screen.MousePointer = vbHourglass
 If ResetActiveControl Then
  If isControlArray(MyForm, MyActiveControl) Then
   MyActCtrlIndex = MyActiveControl.Index
  End If
  MyActCtrlName = MyActiveControl.Name
 End If
 For Each MyControl In MyForm.Controls
  If TypeOf MyControl Is TextBox Or TypeOf MyControl Is CheckBox Or TypeOf MyControl Is ComboBox Then
   With MyControl
    If isControlArray(MyForm, MyControl) Then
     For MyControlCount = 1 To UBound(MyFormValuesOnLoad, 2)
      If MyFormValuesOnLoad(MyName, MyControlCount) = .Name Then
       If MyFormValuesOnLoad(MyIndex, MyControlCount) = .Index Then
        If TypeOf MyControl Is TextBox Then
         If MyFormValuesOnLoad(MyTextOrValue, MyControlCount) <> .Text Then
          bIsDirty = True
          If Reset Then
           .Text = MyFormValuesOnLoad(MyTextOrValue, MyControlCount)
          End If
          If ResetActiveControl Then
           If .Name = MyActCtrlName And .Index = MyActCtrlIndex Then
            .Text = MyFormValuesOnLoad(MyTextOrValue, MyControlCount)
            Screen.MousePointer = vbDefault
            Exit Function
           End If
          End If
          Exit For
         End If
        ElseIf TypeOf MyControl Is CheckBox Then
         If MyFormValuesOnLoad(MyTextOrValue, MyControlCount) <> .Value Then
          bIsDirty = True
          If Reset Then
           .Value = MyFormValuesOnLoad(MyTextOrValue, MyControlCount)
          End If
           If ResetActiveControl Then
           If .Name = MyActCtrlName And .Index = MyActCtrlIndex Then
            .Value = MyFormValuesOnLoad(MyTextOrValue, MyControlCount)
            Screen.MousePointer = vbDefault
            Exit Function
           End If
          End If
          Exit For
         End If
        ElseIf TypeOf MyControl Is ComboBox Then
         If MyFormValuesOnLoad(MyTextOrValue, MyControlCount) <> .ListIndex Then
          bIsDirty = True
          If Reset Then
           .ListIndex = MyFormValuesOnLoad(MyTextOrValue, MyControlCount)
          End If
          If ResetActiveControl Then
           If .Name = MyActCtrlName And .Index = MyActCtrlIndex Then
            .ListIndex = MyFormValuesOnLoad(MyTextOrValue, MyControlCount)
            Screen.MousePointer = vbDefault
            Exit Function
           End If
          End If
          Exit For
         End If
        End If
       End If
      End If
     Next
    Else
     For MyControlCount = 1 To UBound(MyFormValuesOnLoad, 2)
      If MyFormValuesOnLoad(MyName, MyControlCount) = .Name Then
       If TypeOf MyControl Is TextBox Then
        If MyFormValuesOnLoad(MyTextOrValue, MyControlCount) <> .Text Then
         bIsDirty = True
         If Reset Then
          .Text = MyFormValuesOnLoad(MyTextOrValue, MyControlCount)
         End If
         If ResetActiveControl Then
          If .Name = MyActCtrlName Then
           .Text = MyFormValuesOnLoad(MyTextOrValue, MyControlCount)
           Screen.MousePointer = vbDefault
           Exit Function
          End If
         End If
         Exit For
        End If
       ElseIf TypeOf MyControl Is CheckBox Then
        If MyFormValuesOnLoad(MyTextOrValue, MyControlCount) <> .Value Then
         bIsDirty = True
         If Reset Then
          .Value = MyFormValuesOnLoad(MyTextOrValue, MyControlCount)
         End If
         If ResetActiveControl Then
          If .Name = MyActCtrlName Then
           .Value = MyFormValuesOnLoad(MyTextOrValue, MyControlCount)
           Screen.MousePointer = vbDefault
           Exit Function
          End If
         End If
         Exit For
        End If
       ElseIf TypeOf MyControl Is ComboBox Then
        If MyFormValuesOnLoad(MyTextOrValue, MyControlCount) <> .ListIndex Then
         bIsDirty = True
         If Reset Then
          .ListIndex = MyFormValuesOnLoad(MyTextOrValue, MyControlCount)
         End If
         If ResetActiveControl Then
          If .Name = MyActCtrlName Then
           .ListIndex = MyFormValuesOnLoad(MyTextOrValue, MyControlCount)
           Screen.MousePointer = vbDefault
           Exit Function
          End If
         End If
         Exit For
        End If
       End If
      End If
     Next
    End If
   End With
  End If
 Next
 Screen.MousePointer = vbDefault
 IsDirty = bIsDirty
 Exit Function
EH:
 Screen.MousePointer = vbDefault
 MsgBox Err.Description & " in Form " & MyForm.Name, , "IsDirty"
End Function
''''''''''''''''''''''''''''''''''''''''
'This is the Click event for a ToolBar with Buttons you could use on your form
'I used a tool bar because the Active Control such as a Textbox or whatever will
'Remain Active even though you click on the ToolBar Button. This is Handy to know
'if you want to reset Just the Active Textbox to its Original Value.
Private Sub tbrReset_ButtonClick(ByVal Button As MSComctlLib.Button)
'BGS 8/17/99
 On Error GoTo EH
 Select Case Button.Key
  Case "ResetAll"
   If IsDirty(Me, mValuesOnLoad) Then
    Select Case MsgBox("Are you sure you want to Reset All Values ?", vbYesNo + vbQuestion, " Reset to Previous Values")
     Case vbYes
      Call IsDirty(Me, mValuesOnLoad, RESET_VALUES)
     Case vbNo
      Exit Sub
    End Select
   Else
    'MsgBox "Could not find Any Changes to Reset", vbInformation, "Reset"
   End If
  Case "ResetActive"
   Call IsDirty(Me, mValuesOnLoad, , RESET_ACTIVE_CONTROL, Me.ActiveControl)
 End Select
 Exit Sub
EH:
 MsgBox Err.Description & " in Form " & Me.Name, , "ResetToolBar_ButtonClick"
End Sub
''''''''''''''''''''''''''''''''
'This Goes in your Form as a Mod level Variable. it will be used to Store
'All the Values of TextBoxes, CheckBoxes, and ComboBoxes on Load
Private mValuesOnLoad() As Variant
```

