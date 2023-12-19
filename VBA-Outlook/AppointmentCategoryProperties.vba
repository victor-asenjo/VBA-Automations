Option Explicit

Private WithEvents Items As Outlook.Items
Private isProcessing As Boolean

Private Sub Application_Startup()
    ' Initialize variables and event handler
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

Private Sub Items_ItemChange(ByVal Item As Object)
    ' Handle item change events
    If Not isProcessing Then
        Dim Appointment As Outlook.AppointmentItem
        Set Appointment = Item
        
        ' Check if the category is "Holiday" and take appropriate action
        If InStr(1, Appointment.categories, "Holiday", vbTextCompare) > 0 Then
            HandleHolidayAppointment Appointment
        End If
        
        ' Add more conditions for other categories if needed
        ' Copy the block above and make a new function for other categories
        'If InStr(1, Appointment.Categories, "OtherCategory", vbTextCompare) > 0 Then
        '    HandleOtherCategoryAppointment Appointment
        'End If
        
    End If
End Sub

Private Sub HandleHolidayAppointment(ByVal Appointment As Outlook.AppointmentItem)
    ' Handle actions for appointments with "Holiday" category
    ' Disable the macro during processing
    isProcessing = True
    
    ' Modify appointment options
    If Not Appointment.AllDayEvent Then
        Appointment.AllDayEvent = True
    End If
    
    If Appointment.BusyStatus <> olOutOfOffice Then
        Appointment.BusyStatus = olOutOfOffice
    End If
    
    ' Save the appointment only if it has been modified
    If Not Appointment.Saved Then
        Appointment.Save
    End If
    
    ' Re-enable the macro
    isProcessing = False
End Sub
