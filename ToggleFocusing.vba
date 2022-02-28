Option Explicit

'
' Description: Set a single calendar item to reflect "focusing" time in Microsoft Teams
'
' Author     : Hany Elkady
' Reference  : https://pher0ah.blogspot.com/2020/07/TeamsFocusingStatus.html
' Version    : 0.0.2
' Created    : 27/07/2020
' Last Edited: 28/02/2022
'
Sub ToggleCalItemFocusing()
 
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
    
    'Get Current Status
    dndStatus = myCurrentItem.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/string/{6ED8DA90-450B-101B-98DA-00AA003F1305}/IsDoNotDisturbTime")
    
    'Toggle Status
    If dndStatus = False Then
      'Set do not disturb flag
      myCurrentItem.PropertyAccessor.SetProperty "http://schemas.microsoft.com/mapi/string/{6ED8DA90-450B-101B-98DA-00AA003F1305}/IsDoNotDisturbTime", True
      myCurrentItem.PropertyAccessor.SetProperty "http://schemas.microsoft.com/mapi/string/{6ED8DA90-450B-101B-98DA-00AA003F1305}/IsBookedFreeBlocks", True
      If InStr(1, myCurrentItem.Subject, "[FOCUS]") <> 1 Then
        myCurrentItem.Subject = "[FOCUS] " & myCurrentItem.Subject
      End If
 
      'Save item
      myCurrentItem.Save
     
      'Notify User
      MsgBox "This event is now set to focusing", vbOKOnly, "Information"
    Else
      'Unset do not disturb flag
      myCurrentItem.PropertyAccessor.SetProperty "http://schemas.microsoft.com/mapi/string/{6ED8DA90-450B-101B-98DA-00AA003F1305}/IsDoNotDisturbTime", False
      myCurrentItem.PropertyAccessor.SetProperty "http://schemas.microsoft.com/mapi/string/{6ED8DA90-450B-101B-98DA-00AA003F1305}/IsBookedFreeBlocks", False
      If InStr(1, myCurrentItem.Subject, "[FOCUS]") = 1 Then
        myCurrentItem.Subject = Replace(myCurrentItem.Subject, "[FOCUS] ", "", 1, -1, vbTextCompare)
      End If
    
      'Save item
      myCurrentItem.Save
     
      'Notify User
      MsgBox "This event is now set to normal", vbOKOnly, "Information"
    End If
  Else
    MsgBox "You have not selected a calendar item", vbOKOnly, "Error"
  End If

End Sub
