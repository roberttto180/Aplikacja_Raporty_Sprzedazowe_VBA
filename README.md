# Aplikacja Raporty Sprzedażowe VBA
Aplikacja w Excelu korzystająca z VBA do wczytywania danych i generowania raportów sprzedażowych.

Najpierw aplikacja zczytuje dane z plików tekstowych z folderu dane (najpierw trzeba wypakować pliki z zipa, ma być w taki sposób jak na rysunku foto0.png i foto0.5.png). Czyli przykładowo, po ściągnięciu aplikacji sprzedażowej (pliku: Aplikacja_Raporty_Sprzedażowe_VBA.xlsm) umieszczamy go w Pobrane na komputerze. Następnie ściągamy plik dane.zip i umieszczamy go też w Pobrane. Dalej ten plik dane.zip rozpakowujemy i powstaje nowy folder o nazwie dane. Zmieniamy jego nazwę na dane0. Wchodzimy do tego folderu dane0 i kopiujemy (klawisze CTRL+C) z niego folder dane. Dalej w Pobrane wklejamy (klawisze CTRL+V) skopiowany plik. Plik dane0 możemy usunąć, bo nie będzie już potrzebny. Chodziło o to, aby po wejściu w folder dane, nie był widoczny kolejny folder dane, tylko już same pliki tekstowe.

Wczytane dane znajdują się w arkuszu Dane.
Potem w arkuszu Interfejs_uzytkownika pojawia się widok jak na foto1.png.
Możemy wybrać kategorię, dla której ma być wygenerowany raport(domyślnie jest utworzony jako nowy arkusz Excela).
W każdym raporcie zawarte są proste obliczenia z tabeli znajdującej się w arkuszu Dane
oraz tabele przestawne wraz z wykresami przestawnymi.
Jeśli chcemy raporty dlw wszystkich elementów kategorii, to nie zaznaczamy żadnego elementu z pola wyboru.
Na końcu mamy też możliwość wyboru czy chcemy, aby raport był też utworzony dodatkowo jako plik PDF.

