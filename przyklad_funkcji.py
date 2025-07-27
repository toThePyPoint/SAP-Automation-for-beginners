import time


def nazwa_funkcji(parametr1, parametr2):
    # ciało funkcji – czyli instrukcje, które mają się wykonać
    print(f"Witaj, {parametr1}! Twoje hasło to: {parametr2}")


nazwa_funkcji("Kamil", "Python!")
# input()


# konstrucja programu dwóch raportów
def otw_trans_i_wczyt_war(numer_sesji, nazwa_transakcji, nazwa_wariantu):
    # funkcja otwierająca transakcję i wczytująca wariant
    pass


def eksp_danych_do_excl(sciezka_do_pliku, nazwa_pliku):
    # funkcja eksportująca dane do excela
    pass


def pobierz_dane_produkcji(numer_sesji, nazwa_transakcji, nazwa_wariantu, sciezka_do_pliku, nazwa_pliku):
    # funkcja, która pobiera dane dotyczące produkcji
    otw_trans_i_wczyt_war(numer_sesji=numer_sesji, nazwa_transakcji=nazwa_transakcji, nazwa_wariantu=nazwa_wariantu)
    eksp_danych_do_excl(sciezka_do_pliku=sciezka_do_pliku, nazwa_pliku=nazwa_pliku)


def pobierz_dane_sprzedazy(numer_sesji, nazwa_transakcji, nazwa_wariantu, sciezka_do_pliku, nazwa_pliku):
    # funkcja, która pobiera dane dotyczące sprzedaży
    otw_trans_i_wczyt_war(numer_sesji=numer_sesji, nazwa_transakcji=nazwa_transakcji, nazwa_wariantu=nazwa_wariantu)
    eksp_danych_do_excl(sciezka_do_pliku=sciezka_do_pliku, nazwa_pliku=nazwa_pliku)


# Przygotowanie argumentów dla obydwu raportów
tr_sp = "COHV"
tr_pr = "COHV"
war_sp = "wariant_sprzedaż"
war_prod = "wariant_produkcja"
num_okna = 0
sciezka_do_pliku = "C:/Dysk/Raporty/"
nazwa_sp = "dane_sprzdaz.xlsx"
nazwa_prod = "dane_produkcja.xlsx"


# Uruchomienie całego programu
pobierz_dane_produkcji(numer_sesji=num_okna, nazwa_transakcji=tr_pr, nazwa_wariantu=war_prod,
                       sciezka_do_pliku=sciezka_do_pliku, nazwa_pliku=nazwa_prod)
pobierz_dane_sprzedazy(numer_sesji=num_okna, nazwa_transakcji=tr_sp, nazwa_wariantu=war_sp,
                       sciezka_do_pliku=sciezka_do_pliku, nazwa_pliku=nazwa_sp)
