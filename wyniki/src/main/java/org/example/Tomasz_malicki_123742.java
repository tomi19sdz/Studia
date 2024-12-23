package org.example;// pakiet do zarządzania klasami

import java.util.ArrayList; // import biblioteki aray listy
import java.util.List;
import java.util.Scanner; // import biblioteki Scanner do pobierania danych od użytkownika

import org.apache.poi.ss.usermodel.*; //import biblioteki służącej do odczytywania i zapisywania plików formacie EXCEL
import org.apache.poi.xssf.usermodel.XSSFWorkbook; // import biblioteki, która służy do pracy z plikami EXCEL

import java.io.FileInputStream; // praca na pliku "otwieranie"
import java.io.FileOutputStream; // praca na pliku "zapisywanie"
import java.io.File; // pozwala na pracę z plikami czy katalogami
import java.io.IOException; // praca na wyjątkach programu
import java.io.FileNotFoundException;

public class Tomasz_malicki_123742 { // nazwa głównej klasy programu
    public static String selectedFile; // Deklaracja zmiennej selectedFile jako public

    // Konstruktor klasy
    public Tomasz_malicki_123742(String fileName) {
        this.selectedFile = fileName;
    }

    //Metoda zwracająca wartość selectedFile
    public String getSelectedFile() {
        return selectedFile;
    }
    // statyczna publiczna klasa i nie zwraca żadnej wartości
    public static void main(String[] args) throws FileNotFoundException, IOException{
        System.out.println("Tomasz Malicki 123742 Podstawy programowania");
        // pobiera dane od użytkownika
        Scanner sc = new Scanner(System.in);

        mainLoop:// etykieta do której odwołuje sie instrukcja continue
        while (true) { // ta pętla wykonuje program do momentu, gdy użytkownik nie zakończy działania
            System.out.println("Wybierz plik, na którym chcesz pracować");
            String[] optionsplik = {"Piłka_nożna", "Siatkówka", "Piłka_ręczna"};
            for (int k = 0; k < optionsplik.length; k++) { // pętla for do wyświetlania możliwości przez tablice opcji
                System.out.println((k + 1) + ". " + optionsplik[k]); // wyświetlanie listy opcji
            }

            // Pobieranie wyboru pliku od użytkownika
            System.out.print("Wpisz numer pliku: ");
            int plikChoice = sc.nextInt(); // pobieranie danych od użytkownika
            sc.nextLine();// przejscie do następnej lini

            if (plikChoice > 0 && plikChoice<= optionsplik.length) { //sprawdzenie poprawności wyboru
                String selectedFile = optionsplik[plikChoice - 1] + ".xlsx"; // przypisanie wybranego pliku
                System.out.println("Wybrałeś plik: \"" + selectedFile + "\"");



                //podmenu dla wybranego pliku
                boolean running = true; // Zmienna do kontroli pętli

                while (true) { // ta pętla wykonuje program do momentu gdy użytkownik nie zakończy działania
                    // tworzenie opcji wyboru
                    String[] options = {"Zapisz nowe spotkanie do pliku", "Otwórz zapisany plik", "Edytuj spotkanie", "Edytuj wiersz w kolumnie", "Wróć do menu wyboru pliku", "Koniec programu"}; //





                    // Wyświetlanie listy wyboru

                    System.out.println("Wybierz opcje:"); //wyrzuca na ekran dane w ""
                    for (int i = 0; i < options.length; i++) { // pętla for do wyświetlania możliwości przez tablice opcji
                        System.out.println((i + 1) + "." + options[i]);// wyświetlanie listy opcji "i + 1" służy aby wartość zaczynała się od 1 a nie od 0
                    }
                    // Pobieranie danych od użytkownika
                    System.out.print("Wpisz numer wyboru:");
                    int choice = sc.nextInt(); // pobieranie danych od użytkownika tylko w liczbach całkowitych
                    sc.nextLine(); // przejście do następnej lini

                    // Sprawdzenie poprawności wyboru
                    if (choice > 0 && choice <= options.length) { // warunek logiczny sprawdzający czy wartość podane przez użytkownika jest większa od 0
                        System.out.println("Wybrałeś: " + options[choice - 1]); // potwierdzenie wybrania opcji z wydrukowanie opcji na ekranie

                        // opóźnienie wykonywania procesu
                        try {
                            Thread.sleep(2000);
                        } catch (InterruptedException e) {
                            e.printStackTrace();
                        }


                        // Tworzenie obiektu klasy Tomasz_malicki_123742
                        Tomasz_malicki_123742 aObject = new Tomasz_malicki_123742(selectedFile);

                        // Przekazywanie obiektu aObject do klasy ReadExcel
                        ReadExcel.readExcelFile(aObject);
                    } else {
                        System.out.println("Niepoprawny wybór.");


                    }
                    if (choice == 6) { // wybór 5 opcji z listy
                        System.out.println("Wybrano zakończenie programu");
                        System.exit(0);
                    }
                    // Pierwszy wybór opcji
                    else if (choice == 1) {
                        System.out.println("Uzupełnij dane do zapisania");

                        // opóźnienie wykonywania procesu
                        try {
                            Thread.sleep(2000);
                        } catch (InterruptedException e) {
                            e.printStackTrace();
                        }
                        // Dane pobierane od użytkownika w celu uzupełnienia w Exelu
                        System.out.println("Podaj pierwszą drużynę");
                        String Klub1 = sc.nextLine();
                        System.out.println("Podaj wynik spotkania");
                        String Wynik = sc.nextLine();
                        System.out.println("Podaj drugą drużynę");
                        String Klub2 = sc.nextLine();
                        File file = new File(selectedFile);
                        Workbook workbook;
                        Sheet sheet;
                        // warunek jeżeli plik istnieje
                        if (file.exists()) {

                            // Otwiera istniejący plik

                            try (FileInputStream fileInputStream = new FileInputStream(file)) {
                                workbook = new XSSFWorkbook(fileInputStream);//"fileInputStream" jest używany do odczytania danych z pliku
                                sheet = workbook.getSheet(selectedFile);// pobiera arkusz o nazwie selectedFile
                                if (sheet == null) { //sprawdza czy zmienna sheet jest równa null, jesli tak arkusz nie istnieje

                                }
                            } catch (IOException e) { // kod do przechwytywania wyjatku typu IOException
                                e.printStackTrace();
                                return; // kończy wykonywanie bierzącej metody
                            }
                        } else {

                            // Utwórz nowy plik i arkusz jeżeli nie istnieje

                            workbook = new XSSFWorkbook();
                            sheet = workbook.createSheet(selectedFile);// tworzy plik o nazwie wybranej przez użytkownika

                            // Tworzenie nagłówków w kolumnie i wierszu

                            Row headerRow = sheet.createRow(0);
                            Cell headerCell = headerRow.createCell(0);
                            headerCell.setCellValue("Indeks");
                            Cell headerCell1 = headerRow.createCell(1);
                            headerCell1.setCellValue("Drużyna Pierwsza");
                            Cell headerCell2 = headerRow.createCell(2);
                            headerCell2.setCellValue("Wynik");
                            Cell headerCell3 = headerRow.createCell(3);
                            headerCell3.setCellValue("Drużyna Druga");
                        }

                        // Znajdź ostatni wiersz i dodaj nowe dane

                        int lastRowNum = sheet.getLastRowNum();// Automatyczne znalezienie ostatniego wiersza i zaindeksowanie nowego kolejnego
                        Row newRow = sheet.createRow(lastRowNum + 1); //tworzenie kolejnego wiersza po ostatnim znalezionym
                        newRow.createCell(0).setCellValue(lastRowNum + 1);//Tworzenie kolejnego indeksu po ostatnim znalezionym
                        newRow.createCell(1).setCellValue(Klub1);
                        newRow.createCell(2).setCellValue(Wynik);
                        newRow.createCell(3).setCellValue(Klub2);

                        // Zapisz zmiany do pliku

                        try (FileOutputStream fileOutputStream = new FileOutputStream(file)) {//tworzy obiekt który potrzebny jest do zapisania danych w pliku
                            workbook.write(fileOutputStream);// zapis danych do pliku
                            workbook.close();//zamkniecie obiektu workbook

                            // Kod służy do obsługi wyjątków
                        } catch (IOException e) {
                            e.printStackTrace();
                        }
                        int createdRowNum = (int) newRow.getCell(0).getNumericCellValue(); // Pobranie wartości indeksu
                        System.out.println("Dane zostały zapisane do pliku " + selectedFile + " w wierszu " + createdRowNum);
                        // Opóznienie programu
                        try {
                            Thread.sleep(2000);
                        } catch (InterruptedException e) {
                            e.printStackTrace();
                        }

                        System.out.println(createdRowNum + "." + " " + Klub1 + " " + Wynik + " " + Klub2); //Wypisanie zapisanej wartości

                        // Opóznienie programu
                        try {
                            Thread.sleep(3000);
                        } catch (InterruptedException e) {
                            e.printStackTrace();
                        }
                        //Opcja numer 2
                    } else if (choice == 2) {
                        System.out.println("Wybrałeś otworzenie pliku");// Wypisanie opcji wybranej
                        String filePath = "\"" + selectedFile + "\""; //ścieżka do pliku który należy otworzyć

                        try {
                            File file = new File(selectedFile);// otworzenie pliku
                            if (file.exists()) { // sprawdzenie czy plik istnieje

                                // Uruchomienie programu Excel z podanym plikiem
                                Runtime.getRuntime().exec("cmd /c start excel " + selectedFile);
                                System.out.println("Otwieranie pliku Excel: " + selectedFile);
                            } else {
                                System.out.println("Plik nie istnieje: " + selectedFile);// przypadkowa w której plik nie istnieje
                            }
                        } catch (IOException e) {
                            e.printStackTrace();

                            System.out.println("Wystąpił błąd podczas otwierania pliku: " + selectedFile); // Wypisanie na ekranie informacji z błędem
                        }

                    } else if (choice == 3) { // wybór 3 opcji z listy

                        String filePath = selectedFile; // ścieżka do pliku edytowanego
                        try (FileInputStream fis = new FileInputStream(filePath);
                            Workbook workbook = new XSSFWorkbook(fis)) {

                        }
                        System.out.println("Podaj wiersz do edycji");//Wypisanie za pytania do użytkownika
                        int rowNum = sc.nextInt();
                        sc.nextLine(); // znak nowej lini
                        System.out.println("Podaj nową pierwszą drużynę: ");//Wypisanie za pytania do użytkownika
                        String newKlub1 = sc.nextLine();// dane wpisane przez użytkownika
                        System.out.println("Podaj nowy wynik: ");//Wypisanie za pytania do użytkownika
                        String newWynik = sc.nextLine();// dane wpisane przez użytkownika
                        System.out.println("Podaj nową drugą drużynę: ");//Wypisanie za pytania do użytkownika
                        String newKlub2 = sc.nextLine();// dane wpisane przez użytkownika

                        try (FileInputStream fis = new FileInputStream(filePath);// Tworzy strumień wejściowy, który otwiera plik określony przez filePath
                             Workbook workbook = new XSSFWorkbook(fis)) { //Tworzy obiekt Workbook z biblioteki Apache POI, który reprezentuje plik Excel odczytywany za pomocą strumienia fis
                            Sheet sheet = workbook.getSheet(selectedFile); //Pobiera arkusz o nazwie "Mecze" z pliku Excel.

                            if (rowNum > 0 && rowNum <= sheet.getLastRowNum()) { //sprawdza, czy wartość rowNum mieści się w określonym zakresie.
                                Row row = sheet.getRow(rowNum);
                                if (row == null) {
                                    row = sheet.createRow(rowNum);// jeśli wiersz nie istnieje to go tworzy
                                }
                                row.createCell(1).setCellValue(newKlub1);
                                row.createCell(2).setCellValue(newWynik);
                                row.createCell(3).setCellValue(newKlub2);

                                try (FileOutputStream fos = new FileOutputStream(filePath)) {
                                    workbook.write(fos);
                                    System.out.println("Dane w wierszu numer " + rowNum + " zostały zaktualizowane");
                                }
                            } else {
                                System.out.println("Podany numer wiersza jest nie prawidłowy"); // Wypisanie informacji o nie prawidłowym wierszu
                            }

                        } catch (IOException e) { // Tworzenie wyjatku instrukcji wejścia wyjścia
                            e.printStackTrace();
                        }
                    } else if (choice == 4) { // wybór 4 opcji z listy
                        String filePath = selectedFile;
                        System.out.print("Podaj numer wiersza do edycji: ");
                        int rowNum = sc.nextInt(); // numer wiersza do edycji (indeks zaczyna się od 0)
                        sc.nextLine();

                        List<String> lista = new ArrayList<>(); // Wylistowanie nazw kolumn aby łatwiej było napisać nazwę kolumny
                        lista.add("1.Drużyna Pierwsza");
                        lista.add("2.Wynik");
                        lista.add("3.Drużyna Druga");
                        for (String kolumna : lista) {
                            System.out.println(kolumna);//wyświetlanie kolumny list
                        }

                        boolean validColumn = false;
                        String columnName = "";
                        while (!validColumn) { // Pętla będzie wykonywana tak długo, jak validColumn jest false.
                            System.out.print("Podaj nazwę kolumny do edycji: ");
                            columnName = sc.nextLine(); // nazwa kolumny do zmiany, użytkownik sam wpisuje

                            try (FileInputStream fis = new FileInputStream(filePath);
                                 Workbook workbook = new XSSFWorkbook(fis)) {
                                Sheet sheet = workbook.getSheet(selectedFile);
                                if (sheet != null) {
                                    Row headerRow = sheet.getRow(0); // założenie, że nagłówki są w pierwszym wierszu
                                    int colNum = -1;
                                    for (Cell cell : headerRow) {
                                        if (cell.getStringCellValue().equalsIgnoreCase(columnName)) {
                                            colNum = cell.getColumnIndex();
                                            validColumn = true;
                                            break;
                                        }
                                    }

                                    if (colNum == -1) {
                                        System.out.println("Kolumna o nazwie '" + columnName + "' nie istnieje. Spróbuj ponownie."); //wypisanie ze kolumna nie istnieje
                                    } else {
                                        System.out.println("Kolumna o nazwie '" + columnName + "' została znaleziona na indeksie " + colNum); // wypisanie ze kolumna o konkretnej nazwie została znaleziona

                                        Row row = sheet.getRow(rowNum);
                                        if (row == null) { // sprawdzenie czy wiersz istnieje
                                            row = sheet.createRow(rowNum);// tworzenie nowego wiersza na wybranej pozycji
                                        }

                                        Cell cell = row.getCell(colNum);
                                        if (cell == null) {// sprawdzanie czy istnieje komórka
                                            cell = row.createCell(colNum);// tworzenie komórki jeśli nie istnieje
                                        }
                                        System.out.println("Wpisz nową wartość do kolumny");
                                        String newValue = sc.nextLine();// pobranie danych od użytkownika
                                        cell.setCellValue(newValue); // ustawienie nowej wartości komórki

                                        try (FileOutputStream fos = new FileOutputStream(filePath)) {
                                            workbook.write(fos); // zapisanie zmian do pliku
                                            System.out.println("Kolumna została zaktualizowana.");//Wypisanie potwierdzenia zapisania danych
                                        }
                                    }
                                }
                            } catch (IOException e) {
                                e.printStackTrace();
                            }
                        }
                    } else if (choice == 5) { // wybór 5 opcji z listy
                        System.out.println("Powrót do menu wyboru pliku");
                        continue mainLoop;
                    }


                }


            }
        }
    }

}

