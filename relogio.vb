' relogio.bas 

Sub OnOffButton.Click (index as integer)
    If (index = 1 ) then
       AlarmOn = true
    else
       AlarmOn = False
    End If
End Sub

Sub OnOffItem_Click()
    If (alarmOn) then
       AlarmOn = False ´Muda o alarme
       OnOffItem.Caption = "Alarm Off"
    Else
       AlarmOn = True ´Muda o Ala
       OnOffItem.Caption = "Alarm On"
    End If
End Sub

Sub OnOffItem_Click() ´Relogio despertador com menu
    If (AlarmOn then
       AlarmOn = Fals ´Muda o alarme
       AlarmOnItem.Visible = True
       AlarmOffItem = False
    End If
End Sub    

Sub OnItem_Click()
    If (AlarmOn) Then
       AlarmOn = False
       OnItem.Checked = False
    Else
       AlarmOn = True
       OnItem.Checked = True
    End If
End Sub

