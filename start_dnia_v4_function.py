import time

import win32com.client
import sys

czas_start = time.time()
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

    def otworz_transakcje_i_wczytaj_wariant(numer_sesji, nazwa_transakcji, nazwa_wariantu):
        print(f"{(time.time() - czas_start):.2f}s: Wczytuję transakcję wariant {nazwa_wariantu} w transakcji {nazwa_transakcji} w oknie: {numer_sesji}")
        session = connection.Children(numer_sesji)
        session.findById("wnd[0]").maximize()
        session.findById("wnd[0]/tbar[0]/okcd").text = nazwa_transakcji
        session.findById("wnd[0]").sendVKey(0)
        if nazwa_wariantu:
            # Program zrealizuje poniższe linie tylko, jeśli użytkownik podał jakiś wariant
            session.findById("wnd[0]").sendVKey(17)  # CTRL + F5
            session.findById("wnd[1]/usr/txtV-LOW").text = nazwa_wariantu  # podajemy nazwę wariantu
            session.findById("wnd[1]/usr/txtENAME-LOW").text = ""
            session.findById("wnd[1]").sendVKey(0)
            session.findById("wnd[1]").sendVKey(8)
            session.findById("wnd[0]").sendVKey(8)


    # === TWOJA KONFIGURACJA ===
    # Wpisz tutaj transakcje i warianty, które chcesz uruchomić
    # zadania_do_uruchomienia = [
    #     {'transakcja': 'COHV', 'wariant': 'PLAN_LU_ZRI'},  # Słownik dla okna nr 1
    #     {'transakcja': 'MD04', 'wariant': None},  # Słownik dla okna nr 2
    #     {'transakcja': 'COHV', 'wariant': 'PLAN_LU_ZAR'},  # Słownik dla okna nr 3
    #     # Możesz dodać więcej!
    # ]
    zadania_do_uruchomienia = [
        {'transakcja': 'COHV', 'wariant': 'PLAN_LU_ZAR'},
        {'transakcja': 'COHV', 'wariant': 'PLAN_LU_ZAR'},
        {'transakcja': 'COHV', 'wariant': 'PLAN_LU_ZAR'},
        {'transakcja': 'COHV', 'wariant': 'PLAN_LU_ZAR'},
        {'transakcja': 'COHV', 'wariant': 'PLAN_LU_ZAR'},
        {'transakcja': 'COHV', 'wariant': 'PLAN_LU_ZAR'},
    ]

    numer_okna = 0

    for slownik in zadania_do_uruchomienia:

        # w ten sposób pobieramy wartości przypisane do kluczy: 'transakcja' oraz 'wariant'
        # a następnie przypisujemy je do zmiennych wariant oraz transakcja, które będą argumentami do naszej funkcji
        wariant = slownik['wariant']
        transakcja = slownik['transakcja']

        # Wywołujemy funkcję, która realizuje operacje, które wcześniej mnieliśmy w pętli
        otworz_transakcje_i_wczytaj_wariant(numer_sesji=numer_okna, nazwa_transakcji=transakcja, nazwa_wariantu=wariant)

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

print(f"Czas wykonywania skryptu w podejściu sekwencyjnym: {(time.time() - czas_start):.2f}")
