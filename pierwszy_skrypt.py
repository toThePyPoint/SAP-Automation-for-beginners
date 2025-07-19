import win32com.client

# Połączenie z uruchomioną aplikacją i sesją SAP
SapGuiAuto = win32com.client.GetObject("SAPGUI")
application = SapGuiAuto.GetScriptingEngine
connection = application.Children(0)
session = connection.Children(0)

# ---------------------------------------------------------------
#  🔽 Od tego momentu zaczyna się faktyczna interakcja z SAP GUI:
#     (Składnia jest niemal identyczna jak w VBS)

session.findById("wnd[0]").maximize()
session.findById("wnd[0]/tbar[0]/okcd").text = "/nmd04"
session.findById("wnd[0]").sendVKey(0)
session.findById("wnd[0]/usr/tabsTAB300/tabpF01/ssubINCLUDE300:SAPMM61R:0301/ctxtRM61R-MATNR").text = "991111"
session.findById("wnd[0]/usr/tabsTAB300/tabpF01/ssubINCLUDE300:SAPMM61R:0301/ctxtRM61R-WERKS").text = "2101"
session.findById("wnd[0]/usr/tabsTAB300/tabpF01/ssubINCLUDE300:SAPMM61R:0301/ctxtRM61R-WERKS").setFocus()
session.findById("wnd[0]/usr/tabsTAB300/tabpF01/ssubINCLUDE300:SAPMM61R:0301/ctxtRM61R-WERKS").caretPosition = 4
session.findById("wnd[0]").sendVKey(0)
