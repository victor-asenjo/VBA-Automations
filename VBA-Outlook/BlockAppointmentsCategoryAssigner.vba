Option Explicit

Private WithEvents Items As Outlook.Items
Private lastEvent As Object

Private Sub Application_Startup()
    ' Initialize variables and event handler
    InitializeCategoryAssignerVariables
    InitializeHandler
End Sub

Private Sub InitializeHandler()
    ' Set up Outlook folder and items for monitoring
    Dim olNamespace As Outlook.NameSpace
    Dim olFolder As Outlook.Folder
    
    Set olNamespace = Application.GetNamespace("MAPI")
    Set olFolder = olNamespace.GetDefaultFolder(olFolderCalendar)
    
    Set Items = olFolder.Items
End Sub

Private Sub InitializeCategoryAssignerVariables()
    ' Initialize lastEvent variable
    Set lastEvent = Nothing
End Sub

Private Sub Items_ItemAdd(ByVal Item As Object)
    ' Handle item add events
    If TypeOf Item Is Outlook.AppointmentItem Then
        OpenAndCloseOutlookCalendarWindow Item
    End If
End Sub

Private Sub Items_ItemRemove()
    ' Reset variables when items are removed
    InitializeCategoryAssignerVariables
End Sub

Sub OpenAndCloseOutlookCalendarWindow(ByVal Item As Object)
    ' Hidden window process
    Dim olApp As Object
    Dim olNamespace As Object
    Dim olFolder As Object
    
    ' Create a new instance of Outlook
    Set olApp = CreateObject("Outlook.Application")
    ' Get the MAPI namespace
    Set olNamespace = olApp.GetNamespace("MAPI")
    
    Set olFolder = olNamespace.GetDefaultFolder(olFolderCalendar)
    
    ' Check if the active explorer's current folder is "Calendar" and set view accordingly
    If Application.ActiveExplorer.CurrentFolder = "Calendar" Then
        Application.ActiveExplorer.CurrentView = "List"
    End If
    
    ' Process the new appointment item
    ProcessNewAppointmentItem Item
    
    ' Check and set the current view back to "Calendar"
    If Application.ActiveExplorer.CurrentFolder = "Calendar" Then
        Application.ActiveExplorer.CurrentView = "Calendar"
    End If
    
    Dim explorers As Object
    Set explorers = olApp.explorers
    
End Sub

Private Sub ProcessNewAppointmentItem(ByVal Item As Outlook.AppointmentItem)
    ' Process new appointment item
    If Not lastEvent Is Nothing Then
        CompareLastAndNewEvent Item
    Else
        ShowPopupWindow Item
    End If
End Sub

Private Sub CompareLastAndNewEvent(ByVal Item As Outlook.AppointmentItem)
    ' Handle recent events list
    ' Calculate the time difference in seconds
    Dim difTimeInSeconds As Long
    difTimeInSeconds = DateDiff("s", lastEvent.CreationTime, Item.CreationTime)
    
    ' Check if the difference is less than 5 seconds
    If difTimeInSeconds < 5 Then
        ' Assign the category of the last event to the current one
        Item.categories = lastEvent.categories
        
        ' Save the event only if it has been modified
        If Not Item.Saved Then
            Item.Save
        End If
        
        ' Update the last event for the next one
        Set lastEvent = Item
    Else
        ' If the time difference is greater, ask for the category of the new set of appointments
        ShowPopupWindow Item
    End If
End Sub

Private Sub ShowPopupWindow(ByVal objEvent As Outlook.AppointmentItem)
    ' Show a popup window for category selection
    Dim categories As String
    categories = InputBox("Would you like to assign a category to the appointment(s)? Please type your category below:")

    ' Assign the category to all recently created events
    If categories <> "" Then
        objEvent.categories = categories
        
        ' Save the event only if it has been modified
        If Not objEvent.Saved Then
            objEvent.Save
        End If
    End If
    
    ' Update the last event with the current one
    Set lastEvent = objEvent
End Sub
