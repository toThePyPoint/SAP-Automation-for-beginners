import time

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

    # ---------------------------------------------------------------
    # Okno nr 1: Transakcja COHV
    session.findById("wnd[0]").maximize()
    session.findById("wnd[0]/tbar[0]/okcd").text = "/ncohv"
    session.findById("wnd[0]").sendVKey(0)
    session.findById("wnd[0]").sendVKey(17)  # CTRL + F5
    session.findById("wnd[1]/usr/txtV-LOW").text = "PLAN_LU_ZRI"
    session.findById("wnd[1]/usr/txtENAME-LOW").text = ""  # Wyczyszczenie pola z nazwą użytkownika
    session.findById("wnd[1]").sendVKey(0)
    session.findById("wnd[1]").sendVKey(8)
    session.findById("wnd[0]").sendVKey(8)

    session.createSession()  # "Nowe okno GUI"
    time.sleep(1)  # Tutaj dodajemy 1s pauzę w programie, aby nowe okno "zdążyło się uruchomić"

    # ---------------------------------------------------------------
    # Okno nr 2: Transakcja MD04 (pomijamy fragment kodu związany z wczytywaniem wariantu)
    session2 = connection.Children(1)

    session2.findById("wnd[0]/tbar[0]/okcd").text = "/nmd04"
    session2.findById("wnd[0]").sendVKey(0)

    session2.createSession()  # "Nowe okno GUI"
    time.sleep(1)  # Tutaj dodajemy 1s pauzę w programie, aby nowe okno "zdążyło się uruchomić"

    # ---------------------------------------------------------------
    # Okno nr 3: Transakcja COHV (nie uruchamiamy kolejnego okna GUI)
    session3 = connection.Children(2)

    session3.findById("wnd[0]/tbar[0]/okcd").text = "/ncohv"
    session3.findById("wnd[0]").sendVKey(0)
    session3.findById("wnd[0]").sendVKey(17)  # CTRL + F5
    session3.findById("wnd[1]/usr/txtV-LOW").text = "PLAN_LU_ZAR"
    session3.findById("wnd[1]/usr/txtENAME-LOW").text = ""  # Wyczyszczenie pola z nazwą użytkownika
    session3.findById("wnd[1]").sendVKey(0)
    session3.findById("wnd[1]").sendVKey(8)
    session3.findById("wnd[0]").sendVKey(8)


except Exception as e:
    print(f"Wystąpił błąd: {e}")
    print("Upewnij się, że SAP Logon jest uruchomiony i jesteś zalogowany do systemu.")

finally:
    # Opcjonalne: zwolnienie obiektów COM
    session = None
    connection = None
    application = None
    SapGuiAuto = None
