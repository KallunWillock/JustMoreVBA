Attribute VB_Name = "modUserForm_CustomMSGBOX"


Sub AA()
CustomMSG

End Sub
Sub CustomMSG()
Dim msg As frmCustomMSG
Set msg = New frmCustomMSG
Dim BTNS As Variant
BTNS = Array("Fields", "Properties", "Template")
msg.Title = "VBADracula"
msg.MainInstruction = "Code Generation"
msg.Contents = "What would you like to generate code on?"
msg.Icon = IconType.Critical
msg.Buttons = BTNS
msg.PrimaryColour = rgb(120, 20, 150)
'msg.BGColour = BACKDARK
msg.Show vbModal
End Sub

