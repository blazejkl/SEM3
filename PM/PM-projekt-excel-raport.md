# Projekt PM Excel, temat 21 - Regresja wielomianowa z użyciem metody Gaussa

**Autorzy:**  
Jan Kochaniak, Piotr Cywoniuk, Błażej Klepacki


---

## 1. Cel projektu

Celem projektu jest zastosowanie metod rozwiązywania układów równań liniowych do zagadnienia regresji wielomianowej. Chcieliśmy pokazać, że wyznaczenie „optymalnego” wielomianu dopasowanego do danych można sprowadzić do rozwiązania układu równań liniowych i zrealizować to w praktyce w Excelu, korzystając z własnych funkcji w VBA.

---

## 2. Opis problemu i model matematyczny

Dany jest zbiór punktów pomiarowych $(x_i, y_i)$ dla $i = 1, \dots, n$.  
Szukamy wielomianu stopnia $d$

$$
p(x) = a_0 + a_1 x + a_2 x^2 + \dots + a_d x^d,
$$

który jak najlepiej przybliża dane.  
Stosujemy klasyczne kryterium najmniejszych kwadratów:

$$
S(a_0, \dots, a_d) = \sum_{i=1}^n (y_i - p(x_i))^2.
$$

Warunek minimalizacji prowadzi do układu równań liniowych:

$$
A \cdot a = b,
$$

gdzie macierz $A$ zawiera sumy potęg $x_i$, a wektor $b$ sumy postaci $\sum x_i^k y_i$.  
Rozwiązaniem jest wektor współczynników wielomianu: $a = (a_0, \dots, a_d)$.

---

## 3. Metoda numeryczna – Gauss z częściowym wyborem elementu głównego

Do rozwiązania układu $A \cdot a = b$ zastosowaliśmy eliminację Gaussa z częściowym wyborem elementu głównego (pivoting).

W każdym kroku:

- wybieramy największy (wartość bezwzględna) element w kolumnie od wiersza bieżącego w dół,
- zamieniamy wiersze tak, aby największy element był na przekątnej,
- wykonujemy klasyczną eliminację,
- kończymy podstawianiem wstecznym.

Wybór elementu głównego poprawia stabilność obliczeń i zapobiega dzieleniu przez liczby bliskie zeru.

---

## 4. Wskaźniki jakości dopasowania

Oprócz współczynników wielomianu obliczyliśmy dwa podstawowe wskaźniki jakości dopasowania.

### **SSE – suma kwadratów błędów**

$$
\text{SSE} = \sum_{i=1}^n (y_i - \hat y_i)^2,
$$

gdzie $\hat y_i = p(x_i)$.

### **R² – współczynnik determinacji**

Najpierw liczymy średnią:

$$
\bar y = \frac{1}{n} \sum_{i=1}^n y_i.
$$

Zmienność całkowita:

$$
\text{SST} = \sum_{i=1}^n (y_i - \bar y)^2.
$$

Zmienność niewyjaśniona:

$$
\text{SSE} = \sum_{i=1}^n (y_i - \hat y_i)^2.
$$

Współczynnik determinacji:

$$
R^2 = 1 - \frac{\text{SSE}}{\text{SST}}.
$$

Wartość $R^2$ bliska 1 oznacza dobre dopasowanie modelu.

---

## 5. Implementacja w Excelu i VBA

Projekt wykonaliśmy w Excelu z wykorzystaniem VBA.  
Arkusz zawiera:

- kolumna A: wartości $x$  
- kolumna B: wartości $y$  
- kolumna C: wartości $\hat y$ obliczane automatycznie  
- komórka E1: stopień wielomianu $d$  
- komórki od D2: współczynniki $a_0, a_1, \dots, a_d$  
- komórki G1–G2: nagłówek i wartość SSE  
- komórki H1–H2: nagłówek i wartość $R^2$

W projekcie zaimplementowaliśmy następujące funkcje:

### **1. `GaussTab(A As Variant) As Variant`**
Rozwiązanie układu liniowego metodą Gaussa z częściowym wyborem elementu głównego.

### **2. `WspolczynnikiWielomianu(xR, yR, stopien)`**
Tworzy układ równań normalnych na podstawie danych, następnie wywołuje `GaussTab`.

### **3. `PolyEval(x, coeffs)`**
Oblicza wartość wielomianu w punkcie $x$.

### **4. `SSE(yR, yHatR)`**
Liczy sumę kwadratów błędów.

### **5. `R2(yR, yHatR)`**
Liczy współczynnik determinacji $R^2$.

### **6. `RegresjaPrzycisk()`**
Makro podpięte do przycisku.  
Po kliknięciu:

- pobiera dane z kolumn A i B,
- oblicza współczynniki wielomianu,
- wpisuje wartości $\hat y$,
- wylicza SSE i $R^2$,
- tworzy/odświeża wykres z danymi i wielomianem.

Wszystkie obliczenia wykonywane są przez kod, zgodnie z wymaganiem projektu, aby nie używać formuł arkuszowych do generowania wykresów.

---

## 6. Przykładowe wyniki

Dla przykładowych danych z lekkim szumem i wielomianu stopnia $d = 2$ otrzymaliśmy:

- współczynniki w przybliżeniu zgodne z oczekiwanym przebiegiem danych,
- wartość SSE około kilkuset jednostek,
- wartość $R^2 \approx 0.86$,  
co oznacza, że model wyjaśnia ok. 86% zmienności obserwacji.

Na wykresie widać, że wielomian przechodzi „po środku” punktów pomiarowych, zgodnie z zasadą metody najmniejszych kwadratów.

---

## 7. Wnioski

W projekcie pokazaliśmy praktyczne połączenie:

- metod numerycznych (eliminacja Gaussa z pivotingiem),  
- statystyki (regresja wielomianowa),  
- i narzędzi informatycznych (VBA w Excelu).

Zrealizowany arkusz spełnia wszystkie wymagania:

- oblicza współczynniki wielomianu,
- umożliwia wybór stopnia,
- liczy wskaźniki jakości dopasowania,
- generuje wykres danych i funkcji,
- działa jednym przyciskiem.

Projekt pokazał, że nawet proste metody numeryczne można skutecznie zastosować do realnej analizy danych i łatwo zaimplementować w Excelu.

