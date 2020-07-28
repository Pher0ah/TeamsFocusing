'
' Description: Set a single calendar item to reflect "focusing" time in Microsoft Teams
'
' Author     : Hany Elkady
' Reference  : https://pher0ah.blogspot.com/2020/07/TeamsFocusingStatus.html
' Version    : 0.0.1
' Last Edited: 27/07/2020
'
Sub MakeCalendarItemFocusTime()
 
 Dim oItems As Items
 Dim myCurrentItem As AppointmentItem
 Dim dndStatus As Boolean
 
 ' This is the initial reference to an appointment collection.
 Set oItems = Outlook.Application.Session.GetDefaultFolder(olFolderCalendar).Items
 
 ' Get Currently Selected Item
 Set myCurrentItem = Outlook.Application.ActiveExplorer.Selection.Item(1)

 
 ' Check if user has selected an appointment
 If myCurrentItem.Class = olAppointment Then
   On Error Resume Next
   dndStatus = myCurrentItem.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/string/{6ED8DA90-450B-101B-98DA-00AA003F1305}/IsDoNotDisturbTime")
   If dndStatus = False Then
     'Set do not disturb flag
     myCurrentItem.PropertyAccessor.SetProperty "http://schemas.microsoft.com/mapi/string/{6ED8DA90-450B-101B-98DA-00AA003F1305}/IsDoNotDisturbTime", True
     myCurrentItem.PropertyAccessor.SetProperty "http://schemas.microsoft.com/mapi/string/{6ED8DA90-450B-101B-98DA-00AA003F1305}/IsBookedFreeBlocks", True
     Debug.Print " Item dndStatus Added"
 
     'Save item
     myCurrentItem.Save
   Else
     MsgBox "This item is already set to DnD", vbOKOnly, "Information"
   End If
 Else
   MsgBox "You have not selected a calendar item", vbOKOnly, "Error"
 End If
End Sub
