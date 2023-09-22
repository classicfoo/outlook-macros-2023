Private WithEvents myInspectors As Outlook.inspectors

Private Sub Application_Startup()
    ' Set up an event handler for new inspectors
    Set myInspectors = Application.inspectors
End Sub

Private Sub myInspectors_NewInspector(ByVal Inspector As Inspector)
    ' Check the type of the newly created inspector
    Select Case Inspector.currentItem.Class
        Case olTask
            t = Inspector
            ' If task start date and task due date is empty, set the dates to today
            If t.StartDate And t.DueDate = "1/01/4501" Then
                t.StartDate = Date
                t.DueDate = Date
            End If
            ' If category is nothing
            If t.Categories = "" Then
                t.Categories = "MH"
            End If
    End Select
End Sub

