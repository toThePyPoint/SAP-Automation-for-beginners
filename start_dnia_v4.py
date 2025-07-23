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

    def otworz_transakcje(numer_sesji, nazwa_transakcji, nazwa_wariantu):
        session = connection.Children(numer_sesji)
        session.findById("wnd[0]").maximize()
        session.findById("wnd[0]/tbar[0]/okcd").text = nazwa_transakcji  # w ten sposób pobieramy wartość przypisaną do klucza 'transakcja'
        session.findById("wnd[0]").sendVKey(0)
        if slownik['wariant']:
            # Program zrealizuje poniższe linie tylko, jeśli użytkownik podał jakiś wariant
            session.findById("wnd[0]").sendVKey(17)  # CTRL + F5
            session.findById("wnd[1]/usr/txtV-LOW").text = nazwa_wariantu  # podajemy nazwę wariantu
            session.findById("wnd[1]/usr/txtENAME-LOW").text = ""
            session.findById("wnd[1]").sendVKey(0)
            session.findById("wnd[1]").sendVKey(8)
            session.findById("wnd[0]").sendVKey(8)


    # === TWOJA KONFIGURACJA ===
    # Wpisz tutaj transakcje i warianty, które chcesz uruchomić
    zadania_do_uruchomienia = [
        {'transakcja': 'COHV', 'wariant': 'PLAN_LU_ZRI'},  # Słownik dla okna nr 1
        {'transakcja': 'MD04', 'wariant': None},  # Słownik dla okna nr 2
        {'transakcja': 'COHV', 'wariant': 'PLAN_LU_ZAR'},  # Słownik dla okna nr 3
        # Możesz dodać więcej!
    ]

    numer_okna = 0

    for slownik in zadania_do_uruchomienia:

        wariant = slownik['wariant']
        transakcja = slownik['transakcja']

        # Wywołujemy funkcję, która realizuje operacje, które wcześniej mnieliśmy w pętli
        otworz_transakcje(numer_sesji=numer_okna, nazwa_transakcji=transakcja, nazwa_wariantu=wariant)

        if numer_okna < len(zadania_do_uruchomienia) - 1:
            # Po ostatnim oknie nie tworzymy nowej sesji
            session.createSession()  # "Nowe okno GUI"
            time.sleep(1)  # Tutaj dodajemy 1s pauzę w programie, aby nowe okno "zdążyło się uruchomić"
            numer_okna += 1  # zwiększamy numer sesji o 1 po każdej iteracji pętli

except Exception as e:
    print(f"Wystąpił błąd: {e}")
    print("Upewnij się, że SAP Logon jest uruchomiony i jesteś zalogowany do systemu.")

finally:
    # Opcjonalne: zwolnienie obiektów COM
    session = None
    connection = None
    application = None
    SapGuiAuto = None
