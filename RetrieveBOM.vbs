If Not IsObject(application) Then
   Set SapGuiAuto  = GetObject("SAPGUI")
   Set application = SapGuiAuto.GetScriptingEngine
End If
If Not IsObject(connection) Then
   Set connection = application.Children(0)
End If
If Not IsObject(session) Then
   Set session    = connection.Children(0)
End If
If IsObject(WScript) Then
   WScript.ConnectObject session,     "on"
   WScript.ConnectObject application, "on"
End If

'* This is the main section that was recorded by SAP GUI recording software. If SAP changes this must be
'* updated. 
session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").text = "zse16n"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/ctxtGD-TAB").text = "mast"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/chkGS_SELFIELDS-MARK[5,5]").selected = false
session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/chkGS_SELFIELDS-MARK[5,6]").selected = false
session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/chkGS_SELFIELDS-MARK[5,7]").selected = false
session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/chkGS_SELFIELDS-MARK[5,8]").selected = false
session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/chkGS_SELFIELDS-MARK[5,9]").selected = false
session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/chkGS_SELFIELDS-MARK[5,10]").selected = false
session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/chkGS_SELFIELDS-MARK[5,11]").selected = false
session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/chkGS_SELFIELDS-MARK[5,12]").selected = false
session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/ctxtGS_SELFIELDS-LOW[2,2]").text = "5108"
session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/ctxtGS_SELFIELDS-LOW[2,2]").setFocus
session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/ctxtGS_SELFIELDS-LOW[2,2]").caretPosition = 4
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/btnPUSH[4,1]").setFocus
session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/btnPUSH[4,1]").press
session.findById("wnd[1]/tbar[0]/btn[24]").press
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[1]/tbar[0]/btn[8]").press
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]/usr/cntlRESULT_LIST/shellcont/shell").setCurrentCell -1,""

'* This section jumps down to the bottom of the document in order to load new data. 
Rows = session.findByID("wnd[0]/usr/cntlRESULT_LIST/shellcont/shell").RowCount
session.findById("wnd[0]/usr/cntlRESULT_LIST/shellcont/shell").firstVisibleRow = (Rows - 1) / 4
session.findById("wnd[0]/usr/cntlRESULT_LIST/shellcont/shell").selectAll
session.findById("wnd[0]/usr/cntlRESULT_LIST/shellcont/shell").firstVisibleRow = (Rows - 1) / 2
session.findById("wnd[0]/usr/cntlRESULT_LIST/shellcont/shell").selectAll
session.findById("wnd[0]/usr/cntlRESULT_LIST/shellcont/shell").firstVisibleRow = Rows - 1
session.findById("wnd[0]/usr/cntlRESULT_LIST/shellcont/shell").selectAll

'* This section selects all of the data and copies it to the clipboard. 
session.findById("wnd[0]/usr/cntlRESULT_LIST/shellcont/shell").selectAll
session.findById("wnd[0]/usr/cntlRESULT_LIST/shellcont/shell").contextMenu
session.findById("wnd[0]/usr/cntlRESULT_LIST/shellcont/shell").selectContextMenuItemByText "Copy Text"

'* This section resets SAP. 
session.findById("wnd[0]/tbar[0]/btn[15]").press
session.findById("wnd[0]/tbar[0]/btn[15]").press
