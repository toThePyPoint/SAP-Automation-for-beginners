import win32com.client
import sys

try:
    # Pobranie uruchomionej aplikacji SAP GUI
    SapGuiAuto = win32com.client.GetObject("SAPGUI")
    if not isinstance(SapGuiAuto, win32com.client.CDispatch):
        # Ta linia jest dla bezpieczeństwa, zazwyczaj nie jest konieczna
        sys.exit("Nie można było uzyskać dostępu do obiektu SAPGUI.")

    application = SapGuiAuto.GetScriptingEngine
    if not isinstance(application, win32com.client.CDispatch):
        sys.exit("Nie można było uzyskać dostępu do silnika skryptowego.")

    # Połączenie z pierwszą otwartą sesją
    connection = application.Children(0)
    if not isinstance(connection, win32com.client.CDispatch):
        sys.exit("Nie znaleziono aktywnego połączenia.")

    session = connection.Children(0)
    if not isinstance(session, win32com.client.CDispatch):
        sys.exit("Nie znaleziono aktywnej sesji.")

    # ---------------------------------------------------------------
    #  🔽 Od tego momentu zaczyna się faktyczna interakcja z SAP GUI:

    session.findById("wnd[0]").maximize()
    session.findById("wnd[0]/tbar[0]/okcd").text = "/nmd04"
    session.findById("wnd[0]").sendVKey(0)
    session.findById("wnd[0]/usr/tabsTAB300/tabpF01/ssubINCLUDE300:SAPMM61R:0301/ctxtRM61R-MATNR").text = "991111"
    session.findById("wnd[0]/usr/tabsTAB300/tabpF01/ssubINCLUDE300:SAPMM61R:0301/ctxtRM61R-WERKS").text = "2101"
    session.findById("wnd[0]/usr/tabsTAB300/tabpF01/ssubINCLUDE300:SAPMM61R:0301/ctxtRM61R-WERKS").setFocus()
    session.findById("wnd[0]/usr/tabsTAB300/tabpF01/ssubINCLUDE300:SAPMM61R:0301/ctxtRM61R-WERKS").caretPosition = 4
    session.findById("wnd[0]").sendVKey(0)

except Exception as e:
    print(f"Wystąpił błąd: {e}")
    print("Upewnij się, że SAP Logon jest uruchomiony i jesteś zalogowany do systemu.")

finally:
    # Opcjonalne: zwolnienie obiektów COM
    session = None
    connection = None
    application = None
    SapGuiAuto = None