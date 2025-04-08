'This script served as logic for toggling the visibility of the on-call technician
'on the main video wall display screens. This was to enable a quick response to network
'related anomolies that arose during normal operations.

Dim OnCallVisible as boolean

If OnCallVisible = true then
	OnCall_Title_Text.Visible = 1
	OnCall_PnkBox_Border.Visible = 1
	OnCall_Name_Text.Visible = 1
	OnCall_Date_Text.Visible = 1
	OnCall_Date_Line.Visible = 1
	OnCall_Refresh_Button.Visible = 1
	OnCall_Title_Text.Enable = 1
	OnCall_PnkBox_Border.Enable = 1
	OnCall_Name_Text.Enable = 1
	OnCall_Date_Text.Enable = 1
	OnCall_Date_Line.Enable = 1
	OnCall_Refresh_Button.Enable = 1
ElseIf OnCallVisible = false then
	OnCall_Title_Text.Visible = 0
	OnCall_PnkBox_Border.Visible = 0
	OnCall_Name_Text.Visible = 0
	OnCall_Date_Text.Visible = 0
	OnCall_Date_Line.Visible = 0
	OnCall_Refresh_Button.Visible = 0
	OnCall_Title_Text.Enable = 0
	OnCall_PnkBox_Border.Enable = 0
	OnCall_Name_Text.Enable = 0
	OnCall_Date_Text.Enable = 0
	OnCall_Date_Line.Enable = 0
	OnCall_Refresh_Button.Enable = 0
	OnCall_Name_text.Text = "Name"
End If

'--------------------------------------------------------------------------------------

Dim ToggleClick As integer

ToggleClick = 0

ToggleClick = ToggleClick + 1

If ToggleClick > 1 then 
	ToggleClick = 0
End If

If ToggleClick = 1 then
	OnCallOn
ElseIf ToggleClick = 0 Then 
	OnCallOff
End If

'----------------------------------------------------------------------------------------------

Function OnCallOn()
	OnCall_Title_Text.Visible = 1
	OnCall_PnkBox_Border.Visible = 1
	OnCall_Name_Text.Visible = 1
	OnCall_Date_Text.Visible = 1
	OnCall_Date_Line.Visible = 1
	OnCall_Refresh_Button.Visible = 1
	OnCall_Title_Text.Enable = 1
	OnCall_PnkBox_Border.Enable = 1
	OnCall_Name_Text.Enable = 1
	OnCall_Date_Text.Enable = 1
	OnCall_Date_Line.Enable = 1
	OnCall_Refresh_Button.Enable = 1
End Function

Function OnCallOff()
	OnCall_Title_Text.Visible = 0
	OnCall_PnkBox_Border.Visible = 0
	OnCall_Name_Text.Visible = 0
	OnCall_Date_Text.Visible = 0
	OnCall_Date_Line.Visible = 0
	OnCall_Refresh_Button.Visible = 0
	OnCall_Title_Text.Enable = 0
	OnCall_PnkBox_Border.Enable = 0
	OnCall_Name_Text.Enable = 0
	OnCall_Date_Text.Enable = 0
	OnCall_Date_Line.Enable = 0
	OnCall_Refresh_Button.Enable = 0
	OnCall_Name_text.Text = "Name"
End Function
