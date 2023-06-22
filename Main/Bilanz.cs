using Spire.Xls;
using System;
using System.Collections.Generic;
using System.Data.Common;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Runtime.Serialization.Formatters.Binary;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace Main
{
    [Serializable]
    public class Bilanz : Kontenplan
    {
        private readonly string s_aktivKontenColor = "#fc94f4";
        private readonly string s_passivKontenColor = "#fcc404";
        private readonly string s_ertragsKontenColor = "#94d454";
        private readonly string s_aufwandsKontenColor = "#04b4f4";
        private readonly string s_eroeffnungsUndAbschlussKontenColor = "#fcccac";
        private readonly string s_GuvKontenColor = "#fcfc04";

        private readonly string DEFAULT_SHEET_NAME = "Hauptbuchung";

        private readonly List<Buchungssatz> buchungen;

        private Konto[] konten;

        private Konto[] lieferantenKonten;
        private Konto[] kundenKonten;

        private decimal anfangsbestand = 0;
        private int year = -1;
        // Anfangsbestand POSITIV -> HABEN in ebk
        // Anfangsbestand NEGATIV -> SOLL in ebk

        public Bilanz()
        {
            buchungen = new List<Buchungssatz>();

            konten = new Konto[Length()];
            lieferantenKonten = new Konto[GetLieferantenKontenAmount()];
            kundenKonten = new Konto[GetKundenKontenAmount()];

            for (int i = 0; i < konten.Length; i++)
            {
                konten[i] = new();
                Konto currentKonto = konten[i];

                currentKonto.SetDescription(GetDescription(GetKontonummer(i)));
            }
        }

        public Bilanz(int year)
            : this()
        {
            if (year > DateTime.Now.Year || year < 0)
            {
                throw new Exception("Ungueltiges Jahr");
            }
            this.year = year;
        }

        public void AddBuchung(Buchungssatz buchung)
        {
            buchungen.Add(buchung);

            AddToBilanz(buchung);
        }

        private void AddToBilanz(Buchungssatz buchung)
        {
            int[] sollKonten = buchung.GetSollKonten();
            int[] habenKonten = buchung.GetHabenKonten();

            decimal[] sollWerte = buchung.GetSollWerte();
            decimal[] habenWerte = buchung.GetHabenWerte();

            for (int i = 0; i < sollKonten.Length; i++)
            {
                if (sollKonten[i] == 0) continue;

                decimal wert = sollWerte[i];
                int konto = sollKonten[i];

                if (IsLieferantenKonto(konto))
                {
                    if (lieferantenKonten[konto - LIEFERANTEN_KONTO_MIN] == null)
                    {
                        lieferantenKonten[konto - LIEFERANTEN_KONTO_MIN] = new();
                    }

                    lieferantenKonten[konto - LIEFERANTEN_KONTO_MIN].AddSollBetrag(wert, buchung.GetDate(), habenKonten);
                    lieferantenKonten[konto - LIEFERANTEN_KONTO_MIN].SetKontoNummer(konto);

                    continue;
                }
                else if (IsKundenKonto(konto))
                {
                    if (kundenKonten[konto - KUNDEN_KONTO_MIN] == null)
                    {
                        kundenKonten[konto - KUNDEN_KONTO_MIN] = new();
                    }

                    kundenKonten[konto - KUNDEN_KONTO_MIN].AddSollBetrag(wert, buchung.GetDate(), habenKonten);
                    kundenKonten[konto - KUNDEN_KONTO_MIN].SetKontoNummer(konto);

                    continue;
                }

                int kontoPosition = GetPosition(konto);

                konten[kontoPosition].AddSollBetrag(wert, buchung.GetDate(), habenKonten.Where(x => x != 0).ToArray());
            }

            for (int i = 0; i < habenKonten.Length; i++)
            {
                if (habenKonten[i] == 0) continue;

                decimal wert = habenWerte[i];
                int konto = habenKonten[i];

                if (IsLieferantenKonto(konto))
                {
                    if (lieferantenKonten[konto - LIEFERANTEN_KONTO_MIN] == null)
                    {
                        lieferantenKonten[konto - LIEFERANTEN_KONTO_MIN] = new();
                    }

                    lieferantenKonten[konto - LIEFERANTEN_KONTO_MIN].AddHabenBetrag(wert, buchung.GetDate(), sollKonten);
                    lieferantenKonten[konto - LIEFERANTEN_KONTO_MIN].SetKontoNummer(konto);

                    continue;
                }
                else if (IsKundenKonto(konto))
                {
                    if (kundenKonten[konto - KUNDEN_KONTO_MIN] == null)
                    {
                        kundenKonten[konto - KUNDEN_KONTO_MIN] = new();
                    }

                    kundenKonten[konto - KUNDEN_KONTO_MIN].AddHabenBetrag(wert, buchung.GetDate(), sollKonten);
                    kundenKonten[konto - KUNDEN_KONTO_MIN].SetKontoNummer(konto);

                    continue;
                }

                int kontoPosition = GetPosition(konto);

                konten[kontoPosition].AddHabenBetrag(wert, buchung.GetDate(), sollKonten.Where(x => x != 0).ToArray());
            }
        }

#pragma warning disable CS0108
        public void Print()
        {
            for (int i = 0; i < konten.Length; i++)
            {
                if (konten[i].GetEintraege() <= 0)
                {
                    continue;
                }

                PrintKontonummer(konten[i].GetKontoNummer());

                Console.WriteLine("\n-------------------------\n");
            }

            Console.WriteLine("\n==============================\n==============================\n");

            Console.WriteLine("== LIEFERANTENKONTEN ==");

            for (int i = 0; i < lieferantenKonten.Length; i++)
            {
                Konto currKonto = lieferantenKonten[i];

                if (currKonto != null)
                {
                    PrintKontonummer(lieferantenKonten[i].GetKontoNummer());

                    Console.WriteLine("\n-------------------------\n");
                }
            }

            Console.WriteLine("\n==============================\n==============================\n");

            Console.WriteLine("== KUNDENKONTEN ==");

            for (int i = 0; i < kundenKonten.Length; i++)
            {
                if (kundenKonten[i] != null)
                {
                    PrintKontonummer(kundenKonten[i].GetKontoNummer());

                    Console.WriteLine("\n-------------------------\n");
                }
            }

            Console.WriteLine("\n==============================\n==============================\n");
        }

        public void ErfolgskontenAbschlieszen(int year)
        {
            if (year > DateTime.Now.Year || year < 0)
            {
                throw new Exception("Ungueltiges Jahr");
            }
            else if (this.year == -1)
            {
                this.year = year;
            }

            for (int i = 0; i < konten.Length; i++)
            {
                int currKontennummer = konten[i].GetKontoNummer();

                if (IsErfolgskonto(currKontennummer))
                {
                    decimal saldo = konten[i].CalculateSaldo();

                    if (saldo > 0)
                    {
                        // Saldo im SOLL
                        Buchungssatz buchungssatz = new($"{konten[i].GetKontoNummer()} {Math.Abs(konten[i].CalculateSaldo())} EUR" +
                            $" / {SearchFor("GuV")} {Math.Abs(konten[i].CalculateSaldo())} EUR");
                        buchungssatz.SetDatum(year, Monat.December, 31);
                        this.AddBuchung(buchungssatz);
                    }
                    else if (saldo < 0)
                    {
                        // Saldo im HABEN
                        Buchungssatz buchungssatz = new($"{SearchFor("GuV")} {Math.Abs(konten[i].CalculateSaldo())} EUR /" +
                            $" {konten[i].GetKontoNummer()} {Math.Abs(konten[i].CalculateSaldo())} EUR");
                        buchungssatz.SetDatum(year, Monat.December, 31);
                        this.AddBuchung(buchungssatz);
                    }
                } 
            }
        }

        public void ErfolgskontenAbschlieszen(int year, decimal endbestand)
        {
            decimal diff = endbestand - anfangsbestand;

            if (diff > 0)
            {
                Buchungssatz endbestandUmbuchen = new($"{SearchFor("Handelswarenvorrat")} {diff} EUR / {SearchFor("Handelswareneinsatz 20%")} {diff} EUR");
                endbestandUmbuchen.SetDatum(year, Monat.December, 31);
                this.AddBuchung(endbestandUmbuchen);
            }
            else if (diff < 0)
            {
                Buchungssatz endbestandUmbuchen = new($"{SearchFor("Handelswareneinsatz 20%")} {Math.Abs(diff)} EUR / {SearchFor("Handelswarenvorrat")} {Math.Abs(diff)} EUR");
                endbestandUmbuchen.SetDatum(year, Monat.December, 31);
                this.AddBuchung(endbestandUmbuchen);
            }

            this.ErfolgskontenAbschlieszen(year);
        }

        public void SteuernUmbuchen(string date)
        {
            int ustKontennummer = SearchFor("Umsatzsteuer");
            int vstKontennummer = SearchFor("Vorsteuer");

            int ustZahllastKontennummer = SearchFor("Umsatzsteuer-Zahllast");

            decimal ustSaldo = konten[GetPosition(ustKontennummer)].CalculateSaldo();
            decimal vstSaldo = konten[GetPosition(vstKontennummer)].CalculateSaldo();

            var dateParts = ConvertStringToDatum(date);

            int day = dateParts.Day;
            Monat month = dateParts.Monat;
            int year = dateParts.Year;

            Buchungssatz ustBuchungssatz;
            Buchungssatz vstBuchungssatz;

            if (ustSaldo > 0)
            {
                // Ust-Saldo im HABEN
                ustBuchungssatz = new Buchungssatz($"{ustKontennummer} {Math.Abs(ustSaldo)} EUR / {ustZahllastKontennummer} {Math.Abs(ustSaldo)} EUR");
                ustBuchungssatz.SetDatum(year, month, day);
                AddBuchung(ustBuchungssatz);
            }
            else if (ustSaldo < 0)
            {
                //Ust-Saldo im SOLL
                ustBuchungssatz = new Buchungssatz($"{ustZahllastKontennummer} {Math.Abs(ustSaldo)} EUR / {ustKontennummer} {Math.Abs(ustSaldo)} EUR");
                ustBuchungssatz.SetDatum(year, month, day);
                AddBuchung(ustBuchungssatz);
            }

            if (vstSaldo < 0)
            {
                // Vst-Saldo im HABEN
                vstBuchungssatz = new Buchungssatz($"{ustZahllastKontennummer} {Math.Abs(vstSaldo)} EUR / {vstKontennummer} {Math.Abs(vstSaldo)} EUR");
                vstBuchungssatz.SetDatum(year, month, day);
                AddBuchung(vstBuchungssatz);
            }
            else if (vstSaldo > 0)
            {
                // Vst-Saldo im SOLL
                vstBuchungssatz = new Buchungssatz($"{vstKontennummer} {Math.Abs(vstSaldo)} EUR / {ustZahllastKontennummer} {Math.Abs(vstSaldo)} EUR");
                vstBuchungssatz.SetDatum(year, month, day);
                AddBuchung(vstBuchungssatz);
            }
        }

        private static (int Day, Monat Monat, int Year) ConvertStringToDatum(string datum)
        {
            string[] dateParts = datum.Split('.');
            if (dateParts.Length != 3)
            {
                throw new Exception("Falsches Datum");
            }

            int day = Convert.ToInt32(dateParts[0]);

            int i_month = Convert.ToInt32(dateParts[1]);

            if (i_month <= 0 || i_month > 12)
            {
                throw new Exception("Falsches Datum");
            }

            Monat month = new();

            switch (i_month)
            {
                case 1: month = Monat.January; break;
                case 2: month = Monat.February; break;
                case 3: month = Monat.March; break;
                case 4: month = Monat.April; break;
                case 5: month = Monat.May; break;
                case 6: month = Monat.June; break;
                case 7: month = Monat.July; break;
                case 8: month = Monat.August; break;
                case 9: month = Monat.September; break;
                case 10: month = Monat.October; break;
                case 11: month = Monat.November; break;
                case 12: month = Monat.December; break;
            }

            int year = Convert.ToInt32(dateParts[2]);

            return (Day: day, Monat: month, Year: year);
        }

        public void CreateSchlussbilanz(int year)
        {
            int schlussBilanzKonto = SearchFor("Schlussbilanzkonto");

            for (int i = 0; i < konten.Length; i++)
            {
                if (konten[i] == null) continue;
                if (konten[i].GetEintraege() <= 0) continue;

                if (IsErfolgskonto(konten[i].GetKontoNummer())) continue;
                if (konten[i].GetKontoNummer() == schlussBilanzKonto) continue;

                decimal saldo = konten[i].CalculateSaldo();

                if (saldo > 0)
                {
                    // Saldo im SOLL
                    Buchungssatz buchungssatz = new($"{konten[i].GetKontoNummer()} {saldo} EUR / {schlussBilanzKonto} {saldo} EUR");
                    buchungssatz.SetDatum(year, Monat.December, 31);
                    AddBuchung(buchungssatz);
                }
                else if (saldo < 0)
                {
                    // Saldo im HABEN
                    Buchungssatz buchungssatz = new($"{schlussBilanzKonto} {Math.Abs(saldo)} EUR / {konten[i].GetKontoNummer()} {Math.Abs(saldo)} EUR");
                    buchungssatz.SetDatum(year, Monat.December, 31);
                    AddBuchung(buchungssatz);
                }
            }
        }

        public void PrintSchlussbilanz()
        {
            int schlussBilanzKonto = SearchFor("Schlussbilanzkonto");

            PrintKontonummer(schlussBilanzKonto);
        }

        public void PrintGuV()
        {
            int guvKonto = SearchFor("GuV");

            PrintKontonummer(guvKonto);
        }

        public void PrintKontonummer(int kontonummer)
        {
            for (int i = 0; i < konten.Length; i++)
            {
                if (konten[i].GetKontoNummer() == kontonummer)
                {
                    konten[i].Print();

                    decimal saldo = konten[i].CalculateSaldo();

                    if (saldo < 0)
                    {
                        // Saldo ist im HABEN
                        Console.WriteLine($"\n\nSaldo: (Haben) {saldo * -1:f2}");
                    }
                    else if (saldo > 0)
                    {
                        // Saldo ist im SOLL
                        Console.WriteLine($"\nSaldo: (Soll) {saldo:f2}");
                    }
                    else
                    {
                        Console.WriteLine($"\nSaldo: {saldo:f2}");
                    }

                    decimal sollSum = 0, habenSum = 0;

                    for (int j = 0; j < konten[i].GetSollWerte().Count(); j++)
                    {
                        sollSum += konten[i].GetSollWerte()[j];
                    }

                    for (int j = 0; j < konten[i].GetHabenWerte().Count(); j++)
                    {
                        habenSum += konten[i].GetHabenWerte()[j];
                    }

                    Console.WriteLine($"Summe: {sollSum} EUR | {habenSum} EUR");

                    return;
                }
            }

            if (IsLieferantenKonto(kontonummer))
            {
                for (int i = 0; i < lieferantenKonten.Length; i++)
                {
                    if (lieferantenKonten[i] != null && lieferantenKonten[i].GetKontoNummer() == kontonummer)
                    {
                        lieferantenKonten[i].Print();
                        decimal saldo = lieferantenKonten[i].CalculateSaldo();

                        if (saldo < 0)
                        {
                            // Saldo ist im HABEN
                            Console.WriteLine($"\n\nSaldo: (Haben) {saldo * -1:f2}");
                        }
                        else if (saldo > 0)
                        {
                            // Saldo ist im SOLL
                            Console.WriteLine($"\nSaldo: (Soll) {saldo:f2}");
                        }
                        else
                        {
                            Console.WriteLine($"\nSaldo: {saldo:f2}");
                        }

                        return;
                    }
                }
            }

            if (IsKundenKonto(kontonummer))
            {
                for (int i = 0; i < kundenKonten.Length; i++)
                {
                    if (kundenKonten[i] != null && kundenKonten[i].GetKontoNummer() == kontonummer)
                    {
                        kundenKonten[i].Print();
                        decimal saldo = kundenKonten[i].CalculateSaldo();

                        if (saldo < 0)
                        {
                            // Saldo ist im HABEN
                            Console.WriteLine($"\n\nSaldo: (Haben) {saldo * -1:f2}");
                        }
                        else if (saldo > 0)
                        {
                            // Saldo ist im SOLL
                            Console.WriteLine($"\nSaldo: (Soll) {saldo:f2}");
                        }
                        else
                        {
                            Console.WriteLine($"\nSaldo: {saldo:f2}");
                        }

                        return;
                    }
                }
            }
        }

        public void ExportData(string filepath)
        {
            string jsonData = JsonSerializer.Serialize(this);

            using (StreamWriter writer = new StreamWriter(filepath))
            {
                writer.Write(jsonData);
            }

            Console.WriteLine($"Export completed. Total bytes written: {new FileInfo(filepath).Length.ToString("0#,0")}");
        }

        public static Bilanz ImportData(string filepath)
        {
            string jsonData;

            using (StreamReader reader = new StreamReader(filepath))
            {
                jsonData = reader.ReadToEnd();
            }

            Bilanz? bilanz = JsonSerializer.Deserialize<Bilanz>(jsonData);
            return bilanz!;
        }

        public void AddKonto(Konto konto)
        {
            int kontoNummer = konto.GetKontoNummer();

            foreach (Konto currKonto in konten)
            {
                if (currKonto.GetKontoNummer() == kontoNummer)
                {
                    for (int i = 0; i < konto.GetSollWerte().Length; i++)
                    {
                        currKonto.AddSollBetrag(konto.GetSollWerte()[i], konto.GetSollDatum()[i], konto.GetSollGegenkonten()[i].ToArray());
                    }

                    for (int i = 0; i < konto.GetHabenWerte().Length; i++)
                    {
                        currKonto.AddHabenBetrag(konto.GetHabenWerte()[i], konto.GetHabenDatum()[i], konto.GetHabenGegenkonten()[i].ToArray());
                    }
                }
            }
        }

        public void Open()
        {
            for (int i = 0; i < konten.Length; i++)
            {
                if (SearchFor("EBK") == konten[i].GetKontoNummer())
                {
                    if (konten[i] == null)
                    {
                        throw new Exception("EBK-Konto wurde nicht initialisiert");
                    }

                    if (konten[i].GetSollWerte().Length + konten[i].GetHabenWerte().Length <= 0)
                    {
                        throw new Exception("EBK-Konto hat keine Eintraege");
                    }

                    decimal[] habenWerte = konten[i].GetHabenWerte();
                    decimal[] sollWerte = konten[i].GetSollWerte();

                    List<List<int>> habenGegenkonten = konten[i].GetHabenGegenkonten().Distinct().ToList();
                    List<List<int>> sollGegenkonten = konten[i].GetSollGegenkonten().Distinct().ToList();

                    List<Buchungssatz> umbuchungen = new List<Buchungssatz>();

                    for (int j = 0; j < sollGegenkonten.Count; j++)
                    {
                        Buchungssatz umbuchung = new($"{konten[i].GetKontoNummer()} {sollWerte[j]} EUR / {sollGegenkonten[j][0]} {sollWerte[j]} EUR");
                        umbuchungen.Add(umbuchung);

                        if (sollGegenkonten[j][0] == SearchFor("Handelswarenvorrat"))
                        {
                            this.anfangsbestand -= sollWerte[j];
                        }
                    }

                    for (int j = 0; j < habenGegenkonten.Count; j++)
                    {
                        Buchungssatz umbuchung = new($"{habenGegenkonten[j][0]} {habenWerte[j]} EUR / {konten[i].GetKontoNummer()} {habenWerte[j]} EUR");
                        umbuchungen.Add(umbuchung);

                        if (habenGegenkonten[j][0] == SearchFor("Handelswarenvorrat"))
                        {
                            this.anfangsbestand += habenWerte[j];
                        }
                    }

                    konten[i].Clear();

                    foreach(Buchungssatz umbuchung in umbuchungen)
                    {
                        this.AddBuchung(umbuchung);
                    }

                    return;
                }
            }
        }

        public void WriteToExcelFile(string filepath)
        {
            const int TOTAL_COLUMNS = 3;
            const char START_VERTIKAL_REIHE = 'A';

            const double SPALTEN_BREITE = 14;
            const double REIHEN_HOEHE = 25;

            const int EXTRA_LINES_PER_KONTO = 2;

            Workbook workbook = new Workbook();

            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = DEFAULT_SHEET_NAME;

            Konto[] exportKonten = konten.Where(x => x.GetEintraege() > 0).ToArray();

            int kontenPerCol = exportKonten.Length / TOTAL_COLUMNS;
            int rest = exportKonten.Length % TOTAL_COLUMNS;

            int currExportKontenIndex = 0;

            char vertikalReihe = START_VERTIKAL_REIHE; // VERTIKAL = COL
            int horizontalReihe = 2;  // HORIZONTAL = ROW

            for (int i = 0; i < exportKonten.Length; i++)
            {
                if (currExportKontenIndex >= exportKonten.Length)
                {
                    continue;
                }

                int kontenKlasse = (int)exportKonten[currExportKontenIndex].GetKontenklasse();
                int kontenNummer = exportKonten[currExportKontenIndex].GetKontoNummer();

                int kontenEintraege = exportKonten[currExportKontenIndex].GetEintraege();

                #region CellLayout
                Color? currColor = GetKontenklassenColor(kontenNummer);

                string headerRange = $"{(char)(vertikalReihe + 1)}{horizontalReihe}:{(char)(vertikalReihe + 2)}{horizontalReihe}";
                string lineAboveHeader = $"{(char)(vertikalReihe)}{horizontalReihe - 1}:{(char)(vertikalReihe + 3)}{horizontalReihe - 1}";

                sheet.Range[headerRange].Merge();
                sheet.Range[headerRange].IsWrapText = true;
                sheet.Range[headerRange].Style.Font.IsBold = true;

                sheet.Range[headerRange].Style.HorizontalAlignment = HorizontalAlignType.Center;

                sheet.SetValue(horizontalReihe, (int)vertikalReihe - START_VERTIKAL_REIHE + 2, $"{exportKonten[currExportKontenIndex].GetKontoNummer()} {exportKonten[currExportKontenIndex].GetDescription()}");
                

                string kontenRange = "";
                if (currColor != null)
                {
                    kontenRange = $"{(char)vertikalReihe}{horizontalReihe}:{(char)(vertikalReihe + 3)}{horizontalReihe + exportKonten[i].GetEintraege() + EXTRA_LINES_PER_KONTO}";
                    sheet.Range[kontenRange].Style.Color = (Color)currColor!;
                }

                // Zellen schwarzen Rahmen geben
                string currKontoRange = $"{(char)(vertikalReihe)}{(int)horizontalReihe + 1}:{(char)(vertikalReihe + 3)}{(int)horizontalReihe + EXTRA_LINES_PER_KONTO + exportKonten[i].GetEintraege()}"; // + 2 wegen summe und header i guess

                CellRange cell = sheet.Range[currKontoRange];

                cell.Style.Borders[BordersLineType.EdgeLeft].LineStyle = LineStyleType.Thin;
                cell.Style.Borders[BordersLineType.EdgeLeft].Color = Color.Black;
                cell.Style.Borders[BordersLineType.EdgeRight].LineStyle = LineStyleType.Thin;
                cell.Style.Borders[BordersLineType.EdgeRight].Color = Color.Black;
                cell.Style.Borders[BordersLineType.EdgeTop].LineStyle = LineStyleType.Thin;
                cell.Style.Borders[BordersLineType.EdgeTop].Color = Color.Black;
                cell.Style.Borders[BordersLineType.EdgeBottom].LineStyle = LineStyleType.Thin;
                cell.Style.Borders[BordersLineType.EdgeBottom].Color = Color.Black;

                // Add Überschriften und Einträge
                sheet.SetValue(horizontalReihe + 1, vertikalReihe - 'A' + 1, "Datum");
                sheet.SetValue(horizontalReihe + 1, vertikalReihe - 'A' + 2,"Gegenkonto");
                sheet.SetValue(horizontalReihe + 1, vertikalReihe - 'A' + 3, "Soll");
                sheet.SetValue(horizontalReihe + 1, vertikalReihe - 'A' + 4, "Haben");

                sheet.Rows[horizontalReihe - 1].RowHeight = REIHEN_HOEHE;

                #endregion

                #region FillCells

                int horizontalReiheCopy = horizontalReihe;

                if (!string.IsNullOrEmpty(kontenRange))
                {
                    decimal[] sollWerte = exportKonten[currExportKontenIndex].GetSollWerte();
                    string[] sollDatums = exportKonten[currExportKontenIndex].GetSollDatum();
                    List<List<int>> sollGegenkonten = exportKonten[currExportKontenIndex].GetSollGegenkonten();

                    decimal[] habenWerte = exportKonten[currExportKontenIndex].GetHabenWerte();
                    string[] habenDatums = exportKonten[currExportKontenIndex].GetHabenDatum();
                    List<List<int>> habenGegenkonten = exportKonten[currExportKontenIndex].GetHabenGegenkonten();

                    int year = -1;
                    if (sollDatums.Length > 0 && !string.IsNullOrEmpty(sollDatums[0]))
                    {
                        year = Convert.ToInt32(sollDatums[0].Split(".")[sollDatums[0].Split(".").Length - 1]);
                    }
                    else if (habenDatums.Length > 0 && !string.IsNullOrEmpty(habenDatums[0]))
                    {
                        year = Convert.ToInt32(habenDatums[0].Split(".")[habenDatums[0].Split(".").Length - 1]);
                    }

                    List<Entry> mergedEntries = MergeAndSortEntriesByDate(exportKonten[currExportKontenIndex]);

                    for (int k = 0; k < mergedEntries.Count; k++)
                    {
                        Entry entry = mergedEntries[k];

                        decimal wert = entry.Wert;
                        string? datum = entry.Datum!;
                        List<int>? gegenkonten = entry.Gegenkonten!;
                        bool isSoll = entry.IsSoll;

                        if (string.IsNullOrEmpty(datum))
                        {
                            if (gegenkonten.Contains(SearchFor("(EBK)")))
                            {
                                datum = $"1.1.{this.year}";
                            }
                        }
                    }

                    for (int k = 0; k < mergedEntries.Count; k++)
                    {
                        if (kontenEintraege <= 0)
                        {
                            continue;
                        }

                        Entry entry = mergedEntries[k];

                        decimal wert = entry.Wert;
                        string? datum = entry.Datum!;
                        List<int>? gegenkonten = entry.Gegenkonten!;
                        bool isSoll = entry.IsSoll;

                        string gegenkontenMsg = "";

                        for (int j = 0; j < gegenkonten.Count; j++)
                        {
                            gegenkontenMsg += $"{gegenkonten[j]}";
                            if (j < gegenkonten.Count - 1)
                                gegenkontenMsg += ", ";
                        }

                        sheet.SetValue(horizontalReihe + 2, vertikalReihe - 'A' + 1, datum);
                        sheet.SetValue(horizontalReihe + 2, vertikalReihe - 'A' +  2, gegenkontenMsg);

                        sheet.SetNumber(horizontalReihe + 2, vertikalReihe - 'A' + 1 + (isSoll ? 2 : 3), (double)wert);
                        sheet.Range[horizontalReihe + 2, vertikalReihe - 'A' + 1 + (isSoll ? 2 : 3)].NumberFormat = "€\" \"#,##0.00";

                        if (kontenNummer == SearchFor("GuV"))
                        {
                            Color? erfolgsKontenColor = GetKontenklassenColor(Convert.ToInt32(gegenkontenMsg.Split(",")[0]));

                            if (gegenkontenMsg.Split(",")[0].Equals(SearchFor("SBK")))
                            {
                                erfolgsKontenColor = Color.Transparent;
                            }

                            sheet.Range[$"{(char)(vertikalReihe)}{horizontalReihe + 2}:{(char)(vertikalReihe + 3)}{horizontalReihe + 2}"].Style.Color = (Color)erfolgsKontenColor!;
                        }

                        horizontalReihe++;
                    }

                    string cellBottomRight = kontenRange.Split(":")[kontenRange.Split("").Length];

                    string cellBottomRightNumber = string.Join("", Regex.Matches(cellBottomRight, @"\d+"));

                    string cellBottomRightOneLeft = (char)(kontenRange.Split(":")[kontenRange.Split("").Length][0] - (char) 1) + cellBottomRightNumber;

                    sheet.Range[cellBottomRightOneLeft].Formula = 
                        $"SUM({GetCellPosition(horizontalReihe + 1, vertikalReihe - 'A' + 1 + 2)}"
                        + $":{GetCellPosition(horizontalReihe - kontenEintraege + 2, vertikalReihe - 'A' + 1 + 2)})";
                    sheet.Range[cellBottomRightOneLeft].NumberFormat = "€\" \"#,##0.00";

                    sheet.Range[cellBottomRight].Formula =
                        $"SUM({GetCellPosition(horizontalReihe + 1, vertikalReihe - 'A' + 1 + 3)}"
                        + $":{GetCellPosition((horizontalReihe - kontenEintraege) + 2, vertikalReihe - 'A' + 1 + 3)})";
                    sheet.Range[cellBottomRight].NumberFormat = "€\" \"#,##0.00";
                }

                #endregion

                horizontalReihe = horizontalReiheCopy;

                vertikalReihe += (char)5;

                if ((i + 1) % 3 == 0)
                {
                    int horizontalReiheInc = Math.Max(exportKonten[i].GetEintraege() + 2, Math.Max(exportKonten[i - 1].GetEintraege() + 2, exportKonten[i - 2].GetEintraege() + 2));

                    horizontalReihe += horizontalReiheInc;
                    horizontalReihe += EXTRA_LINES_PER_KONTO;

                    vertikalReihe = 'A';
                }

                if (i != 0)
                {
                    sheet.Range[lineAboveHeader].Style.Color = Color.Transparent;
                }

                currExportKontenIndex++;
            }

            for (int i = 0; i < sheet.Columns.Length; i++)
            {
                sheet.Columns[i].ColumnWidth = SPALTEN_BREITE;
            }

            workbook.SaveToFile(filepath);

            System.Diagnostics.Process.Start("cmd.exe", "/c " + filepath);
        }

        private static List<Entry> MergeAndSortEntriesByDate(Konto konto)
        {
            List<Entry> result = new List<Entry>();

            decimal[] sollWerte = konto.GetSollWerte();
            string[] sollDatums = konto.GetSollDatum();
            List<List<int>> sollGegenkonten = konto.GetSollGegenkonten();

            decimal[] habenWerte = konto.GetHabenWerte();
            string[] habenDatums = konto.GetHabenDatum();
            List<List<int>> habenGegenkonten = konto.GetHabenGegenkonten();

            for (int k = 0; k < sollWerte.Length; k++)
            {
                Entry entry = new Entry
                {
                    Wert = sollWerte[k],
                    Datum = sollDatums[k],
                    Gegenkonten = sollGegenkonten[k],
                    IsSoll = true
                };
                result.Add(entry);
            }

            for (int k = 0; k < habenWerte.Length; k++)
            {
                Entry entry = new Entry
                {
                    Wert = habenWerte[k],
                    Datum = habenDatums[k],
                    Gegenkonten = habenGegenkonten[k],
                    IsSoll = false
                };
                result.Add(entry);
            }

            result = SortEntries(result);

            return result;
        }

        private static List<Entry> SortEntries(List<Entry> entries)
        {
            List<Entry> result = new List<Entry>();

            foreach (Entry entry in entries)
            {
                result.Add(entry);
            }

            result.Sort((a, b) =>
            {
                DateTime? dateA = null;
                DateTime? dateB = null;

                if (!string.IsNullOrEmpty(a.Datum))
                {
                    DateTime.TryParseExact(a.Datum, new[] { "d.M.yyyy", "dd.M.yyyy", "d.MM.yyyy", "dd.MM.yyyy" },
                        CultureInfo.InvariantCulture, DateTimeStyles.None, out var parsedDateA);
                    dateA = parsedDateA;
                }

                if (!string.IsNullOrEmpty(b.Datum))
                {
                    DateTime.TryParseExact(b.Datum, new[] { "d.M.yyyy", "dd.M.yyyy", "d.MM.yyyy", "dd.MM.yyyy" },
                        CultureInfo.InvariantCulture, DateTimeStyles.None, out var parsedDateB);
                    dateB = parsedDateB;
                }

                if (dateA.HasValue && dateB.HasValue)
                {
                    return dateA.Value.CompareTo(dateB.Value);
                }
                else if (dateA.HasValue)
                {
                    return -1;
                }
                else if (dateB.HasValue)
                {
                    return 1;
                }
                else
                {
                    return 0;
                }
            });

            return result;
        }
        
        private Color? GetKontenklassenColor(int kontoNummer)
        {
            int kontenklasse = Convert.ToInt32(Convert.ToString(kontoNummer)[0]) - '0';

            if (kontoNummer == SearchFor("Guv"))
            {
                return ColorTranslator.FromHtml(s_GuvKontenColor);
            }
            else if (kontoNummer == SearchFor("SBK"))
            {
                return Color.Transparent;
            }

            if (0 <= kontenklasse && kontenklasse <= 2)
            {
                return ColorTranslator.FromHtml(s_aktivKontenColor);
            }
            else if (kontenklasse == 3)
            {
                return ColorTranslator.FromHtml(s_passivKontenColor);
            }

            else if (kontenklasse == 4 || (kontenklasse == 8 && 0 <= kontoNummer.ToString()[1] && kontoNummer.ToString()[1] <= 2))
            {
                return ColorTranslator.FromHtml(s_ertragsKontenColor);
            }
            else if ((5 <= kontenklasse && kontenklasse <= 7) || (kontenklasse == 8 && kontoNummer.ToString()[1] > 2))
            {
                return ColorTranslator.FromHtml(s_aufwandsKontenColor);
            }
            else if (kontenklasse == 9)
            {
                return ColorTranslator.FromHtml(s_eroeffnungsUndAbschlussKontenColor);
            }
            else
            {
                return null; // unbekannte kontennummer
            }
        }

        public string GetCellPosition(int row, int col)
        {
            char columnLetter = (char)('A' + (col - 1));
            string cellPosition = $"{columnLetter}{row}";

            return cellPosition;
        }
    }
}
