import time
import win32com.client
import subprocess
import os
import sys


class NazwySystemowSAP:
    SYSTEM_P11 = "P11 Single Sign-On [ERP PRD]"
    SYSTEM_K11 = "K11 [ERP QAS]"


def otworz_sap():
    # Ścieżka do pliku wykonywalnego SAP GUI (np. saplogon.exe)
    sciezka_do_sap_gui = r"C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe"

    # Sprawdzenie, czy plik istnieje
    if os.path.exists(sciezka_do_sap_gui):
        # Uruchomienie SAP GUI
        subprocess.Popen(sciezka_do_sap_gui)
    else:
        print(f"Błąd: Nie znaleziono SAP GUI pod ścieżką {sciezka_do_sap_gui}")

    # Krótkie wstrzymanie, by GUI zdążyło się załadować
    time.sleep(2)


def zaloguj_do_sap(system_sap):
    # Inicjalizacja silnika SAP GUI Scripting
    sap_gui_auto = win32com.client.GetObject("SAPGUI")
    aplikacja = sap_gui_auto.GetScriptingEngine

    # Nawiązanie połączenia z SAP na podstawie identyfikatora systemu
    # (nie trzeba podawać danych logowania, jeśli działa SSO)
    polaczenie = aplikacja.OpenConnection(system_sap, True)


def otworz_transakcje_i_wczytaj_wariant(numer_sesji, nazwa_transakcji, nazwa_wariantu):
    # Inicjalizacja COM w nowym procesie
    SapGuiAuto = win32com.client.GetObject("SAPGUI")
    application = SapGuiAuto.GetScriptingEngine
    connection = application.Children(0)
    session = connection.Children(numer_sesji)

    session.findById("wnd[0]").maximize()
    session.findById("wnd[0]/tbar[0]/okcd").text = nazwa_transakcji
    session.findById("wnd[0]").sendVKey(0)
    print(f"Transakcja {nazwa_transakcji} została uruchomiona.")

    if nazwa_wariantu:
        session.findById("wnd[0]").sendVKey(17)  # CTRL + F5
        session.findById("wnd[1]/usr/txtV-LOW").text = nazwa_wariantu
        session.findById("wnd[1]/usr/txtENAME-LOW").text = ""
        session.findById("wnd[1]").sendVKey(0)
        session.findById("wnd[1]").sendVKey(8)
        session.findById("wnd[0]").sendVKey(8)
        print(f"Wariant {nazwa_wariantu} został wczytany.")


if __name__ == "__main__":

    # Tutaj otwieramy SAP-a i logujemy się do systemu
    otworz_sap()
    zaloguj_do_sap(NazwySystemowSAP.SYSTEM_P11)

    # Inicjalizacja COM w procesie głównym
    SapGuiAuto = win32com.client.GetObject("SAPGUI")
    application = SapGuiAuto.GetScriptingEngine
    connection = application.Children(0)
    session = connection.Children(0)

    # numer_okna = 0
    # transakcja = "COHV"
    # wariant = "PLAN_LU_ZAR"
    argumenty_tekst = str(sys.argv[1])  # tym razem argumenty mają postać "transakcja,wariant,numer_okna"
    argumenty_lista = argumenty_tekst.split(',')  # dzielimy tekst na listę z trzema elementami
    # pobieramy poszczególne elementy z listy
    transakcja = argumenty_lista[0]  # pobieramy nazwę transakcji do zmiennej 'trnsakcja'
    wariant = argumenty_lista[1]  # pobieramy nazwę wariantu...
    numer_okna = int(argumenty_lista[2])  # konwertujemy tekst na liczbę całkowitą - integer
    # print(numer_okna)
    # print(transakcja)
    # print(wariant)

    try:
        # Program próbuje uruchomić tę część kodu
        otworz_transakcje_i_wczytaj_wariant(numer_sesji=numer_okna, nazwa_transakcji=transakcja, nazwa_wariantu=wariant)
        sys.exit(0)  # pomyślne zakończenie programu
    except Exception as e:
        # Ta część kodu zostanie wykonana tylko, jeśli wystąpi wyjątek
        sys.exit(1)  # kończmy program z kodem zamknięcia 1 - co oznacza błąd





# if __name__ == "__main__":
#     print(sys.argv)
#     time.sleep(10)
#     input("Wciśnij enter aby kontynuować...")
