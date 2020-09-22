Attribute VB_Name = "modSystem"
Sub Main()
' Startup functions - This function is executed at runtime
SetupInitialEnvironment
ShowInitialForm
End Sub
Function SetupInitialEnvironment()
' Setup the 'Connection Not Established' environment
frmMain.btn_connect.Enabled = True
frmMain.btn_connect.Default = True
frmMain.btn_disconnect.Enabled = False
frmMain.btn_send.Enabled = False
frmMain.fra_connectiondetails.Enabled = True
frmMain.txt_address.Enabled = True
frmMain.txt_incoming.Enabled = False
frmMain.txt_port.Enabled = True
frmMain.txt_send.Enabled = False
frmMain.lbl_address.Enabled = True
frmMain.lbl_port.Enabled = True
frmMain.lbl_incoming.Enabled = False
frmMain.lbl_send.Enabled = False
frmMain.lbl_status.Enabled = True
frmMain.lbl_status_status = True
frmMain.lbl_status_status.BackColor = &HC0&
frmMain.lbl_status_status.Caption = "Disconnected"
End Function
Function ShowInitialForm()
'Load the form into memory
Load frmMain
' Show it on screen
frmMain.Show
End Function
Function SetupTCPClosureEnvironment()
' Setup the 'Connection Not Established' environment
frmMain.btn_connect.Enabled = True
frmMain.btn_connect.Default = True
frmMain.btn_disconnect.Enabled = False
frmMain.btn_send.Enabled = False
frmMain.fra_connectiondetails.Enabled = True
frmMain.txt_address.Enabled = True
frmMain.txt_incoming.Enabled = False
frmMain.txt_port.Enabled = True
frmMain.txt_send.Enabled = False
frmMain.lbl_address.Enabled = True
frmMain.lbl_port.Enabled = True
frmMain.lbl_incoming.Enabled = False
frmMain.lbl_send.Enabled = False
frmMain.lbl_status.Enabled = True
frmMain.lbl_status_status = True
frmMain.lbl_status_status.BackColor = &HC0&
frmMain.lbl_status_status.Caption = "Disconnected"
frmMain.txt_address.SetFocus
End Function
Function SetupTCPConnectEnvironment()
' Setup the 'Connection Established' environment
frmMain.btn_connect.Enabled = False
frmMain.btn_disconnect.Enabled = True
frmMain.btn_send.Enabled = True
frmMain.btn_send.Default = True
frmMain.fra_connectiondetails.Enabled = True
frmMain.txt_address.Enabled = False
frmMain.txt_incoming.Enabled = True
frmMain.txt_port.Enabled = False
frmMain.txt_send.Enabled = True
frmMain.lbl_address.Enabled = False
frmMain.lbl_port.Enabled = False
frmMain.lbl_incoming.Enabled = True
frmMain.lbl_send.Enabled = True
frmMain.lbl_status.Enabled = True
frmMain.lbl_status_status = True
frmMain.lbl_status_status.BackColor = &HC000&
frmMain.lbl_status_status.Caption = "Connected"
frmMain.txt_send.SetFocus
End Function
