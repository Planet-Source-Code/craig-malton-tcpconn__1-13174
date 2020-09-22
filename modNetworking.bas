Attribute VB_Name = "modNetworking"
Function TCPConnect(conn_address, conn_port)
With frmMain.wsk_conn
' Close the Winsock (Just in case!)
.Close
' Attempt the connection
.Connect conn_address, conn_port
End With
End Function
Function TCPSend(DataToSend)
' Actually send the data
frmMain.wsk_conn.SendData DataToSend
End Function
