# Portfolio
These files aren't guaranteed to be fully function, as the final debugged code may be property of my past clients. 

# Smart Picker Code
```
Option Explicit

Public WithEvents txtBox As MSForms.TextBox
Public WithEvents lstBox As MSForms.ListBox
Const m_sNames As String = "Names!$A$1:$A$69"


Private Sub lstBox_Click()
' process when item in list is selected
    [testcell] = lstBox
End Sub


Private Sub txtBox_Change()
' process when text is typed into textbox
    Update_ListBox
End Sub


Sub Update_ListBox()
' load listbox with all names containing typed letters
    Const sMismatch As String = "no"
    
    ' get array indicating all matching and non-matching items.
    ' for each item, array contains "no" if no match, or person-name if match

    ' for example, if "g" is typed, parses to:
    ' TRANSPOSE(IFERROR(IF(SEARCH("*g*", Names!$A$1:$A$69), Names!$A$1:$A$69),"no"))

    ' returns array with 1 for each match, error-code for all others
    ' SEARCH("*g*", Names!$A$1:$A$69)

    ' substitutes the person-name for each match
    ' IF( [previous-step] Names!$A$1:$A$69)

    ' substitute "no" for all error-codes (mismatches)
    ' IFERROR( [previous step], "no")

    ' transpose, needed for Filter and listbox
    ' TRANSPOSE ( [previous step] )
   
    Dim arbMatches()
    arbMatches = Application.Evaluate("TRANSPOSE(IFERROR(IF(SEARCH(""*" & txtBox.Value & "*"", " _
        & m_sNames & "), " & m_sNames & "),""" & sMismatch & """))")
    
    Dim arSelected
    arSelected = Filter(arbMatches, sMismatch, False)
    lstBox.List = arSelected
End Sub


Private Sub UserForm_Initialize()
' set up form
    ' must programmatically add controls, to ensure good rendering on Mac/Windows
    Const lWidth As Long = 200
    Width = lWidth + 5
    
    ' set up textbox
    Set txtBox = Controls.Add("Forms.TextBox.1", "TextBox1")
    With txtBox
        .Width = lWidth
        .Font.Size = 12
        .Height = 20
        .TabIndex = 0
    End With
    
    ' set up listbox
    Set lstBox = Controls.Add("Forms.ListBox.1", "ListBox1")
    With lstBox
        .Width = lWidth
        
        ' must load array variable first. Fails if try to pass sheet fx directly into .List
        ' must transpose a vertical range for .List
        Dim arList
        arList = WorksheetFunction.Transpose(Sheet1.Range(m_sNames))
        .List = arList
        .Top = 22
        .Font.Size = 11
        
        ' this sequence is required, in this exact order, to ensure last item in list is visible.
        .IntegralHeight = False
        DoEvents
        ' must subtract height of titlebar and borders, so list doesn't fall off form.
        .Height = Me.Height - 47
        .IntegralHeight = True
    End With

    txtBox.SetFocus
End Sub
```

