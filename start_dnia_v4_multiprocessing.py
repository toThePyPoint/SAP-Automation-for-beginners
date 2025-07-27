import time
import multiprocessing  # importujemy niezbędną bibliotekę
import win32com.client

czas_start = time.time()


def otworz_transakcje_i_wczytaj_wariant(numer_sesji, nazwa_transakcji, nazwa_wariantu, start):
    # Inicjalizacja COM w nowym procesie
    print(f"{(time.time() - start):.2f}s: Wczytuję transakcję wariant {nazwa_wariantu} w transakcji {nazwa_transakcji} w oknie: {numer_sesji}")
    SapGuiAuto = win32com.client.GetObject("SAPGUI")
    application = SapGuiAuto.GetScriptingEngine
    connection = application.Children(0)
    session = connection.Children(numer_sesji)

    session.findById("wnd[0]").maximize()
    session.findById("wnd[0]/tbar[0]/okcd").text = nazwa_transakcji
    session.findById("wnd[0]").sendVKey(0)

    if nazwa_wariantu:
        session.findById("wnd[0]").sendVKey(17)  # CTRL + F5
        session.findById("wnd[1]/usr/txtV-LOW").text = nazwa_wariantu
        session.findById("wnd[1]/usr/txtENAME-LOW").text = ""
        session.findById("wnd[1]").sendVKey(0)
        session.findById("wnd[1]").sendVKey(8)
        session.findById("wnd[0]").sendVKey(8)


if __name__ == "__main__":
    # Inicjalizacja COM w procesie głównym
    SapGuiAuto = win32com.client.GetObject("SAPGUI")
    application = SapGuiAuto.GetScriptingEngine
    connection = application.Children(0)
    session = connection.Children(0)

    # === TWOJA KONFIGURACJA ===
    zadania_do_uruchomienia = [
        {'transakcja': 'COHV', 'wariant': 'PLAN_LU_ZAR'},
        {'transakcja': 'COHV', 'wariant': 'PLAN_LU_ZAR'},
        {'transakcja': 'COHV', 'wariant': 'PLAN_LU_ZAR'},
        {'transakcja': 'COHV', 'wariant': 'PLAN_LU_ZAR'},
        {'transakcja': 'COHV', 'wariant': 'PLAN_LU_ZAR'},
        {'transakcja': 'COHV', 'wariant': 'PLAN_LU_ZAR'},
    ]

    numer_okna = 0
    procesy = []  # Tworzymy pustą listę, w której będziemy przechowywać nasze procesy

    for slownik in zadania_do_uruchomienia:
        wariant = slownik['wariant']
        transakcja = slownik['transakcja']

        # Tworzymy nowy proces, który uruchomi daną transakcję w osobnym oknie SAP
        proces = multiprocessing.Process(
            target=otworz_transakcje_i_wczytaj_wariant,
            args=(numer_okna, transakcja, wariant, czas_start)
        )
        procesy.append(proces)  # Dodajemy proces do listy, by później móc na niego zaczekać
        proces.start()  # Uruchamiamy proces (czyli otwieranie i konfigurację okna)

        if numer_okna < len(zadania_do_uruchomienia) - 1:
            session.createSession()
            time.sleep(1)
            numer_okna += 1

    # ⏳ Główny program czeka, aż wszystkie okna SAP zakończą swoje zadania
    for proces in procesy:
        proces.join()

    print(f"Czas wykonywania skryptu w podejściu multiprocessing: {(time.time() - czas_start):.2f}")
