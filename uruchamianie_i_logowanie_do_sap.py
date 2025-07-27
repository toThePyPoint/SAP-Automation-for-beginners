import os
import time
import subprocess
import win32com.client


class NazwySystemowSAP:
    """
    Klasa przechowująca nazwy systemów SAP, z którymi możemy się połączyć.
    Dzięki temu możesz łatwo przełączać się np. pomiędzy systemem produkcyjnym i testowym
    Możesz zdefiniować tyle nazw, ile potrzebujesz
    """
    SYSTEM_Prod = "Nazwa_systemu_prod"
    SYSTEM_Test = "Nazwa_systemu_test"


def otworz_sap():
    """
    Uruchamia program SAP Logon, jeśli znajduje się w domyślnej lokalizacji.
    """
    sciezka_do_sap_gui = r"C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe"

    if os.path.exists(sciezka_do_sap_gui):
        # Uruchomienie SAP GUI
        subprocess.Popen(sciezka_do_sap_gui)
    else:
        print(f"Błąd: Nie znaleziono SAP GUI pod ścieżką {sciezka_do_sap_gui}")

    # Odczekanie chwili, aby GUI mogło się poprawnie załadować
    time.sleep(2)


def zaloguj_do_sap(system_sap):
    """
    Łączy się z wybranym systemem SAP za pomocą SAP GUI Scripting.

    Parametry:
    system_sap (str): Nazwa systemu SAP (zgodna z SAP Logon, np. "P11 Single Sign-On [ERP PRD]")

    Zwraca:
    connection: Obiekt połączenia z systemem SAP
    """
    # Inicjalizacja silnika SAP GUI Scripting
    sap_gui_auto = win32com.client.GetObject("SAPGUI")
    aplikacja = sap_gui_auto.GetScriptingEngine

    # Otwarcie połączenia z systemem
    # w dotychczasowej pracy ze skryptami spotkałeś się z tym obiektem pod angielską nazwą "connection"
    polaczenie = aplikacja.OpenConnection(system_sap, True)

    return polaczenie


# -----------------------------
# Przykład użycia programu
# -----------------------------

if __name__ == "__main__":
    # Krok 1: Uruchom SAP GUI
    otworz_sap()

    # Krok 2: Zaloguj się do wybranego systemu (tu: system produkcyjny)
    polaczenie = zaloguj_do_sap(NazwySystemowSAP.SYSTEM_Prod)

    # Dalej możesz wykonać inne operacje na obiekcie `polaczenie`, np. otworzyć transakcję
