using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Serialization.Formatters.Binary;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Threading.Tasks;

namespace Main
{
    [Serializable]
    public class Bilanz : Kontenplan
    {
        private readonly List<Buchungssatz> buchungen;

        private Konto[] konten;

        private Konto[] lieferantenKonten;
        private Konto[] kundenKonten;

        private decimal anfangsbestand = 0;
        // Anfangsbestand POSITIV -> HABEN in ebk
        // Anfangsbestand NEGATIV -> SOLL in ebk

        public Bilanz()
        {
            buchungen = new List<Buchungssatz>();

            konten = new Konto[new Kontenplan().Length];
            lieferantenKonten = new Konto[new Kontenplan().LieferantenKontenAmount];
            kundenKonten = new Konto[new Kontenplan().KundenKontenAmount];

            for (int i = 0; i < konten.Length; i++)
            {
                konten[i] = new();
                Konto currentKonto = konten[i];

                currentKonto.SetDescription(GetDescription(GetKontonummer(i)));
            }
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

        public void Print()
        {
            for (int i = 0; i < konten.Length; i++)
            {
                if (konten[i].Eintraege <= 0)
                {
                    continue;
                }

                PrintKontonummer(konten[i].KontoNummer);

                Console.WriteLine("\n-------------------------\n");
            }

            Console.WriteLine("\n==============================\n==============================\n");

            Console.WriteLine("== LIEFERANTENKONTEN ==");

            for (int i = 0; i < lieferantenKonten.Length; i++)
            {
                Konto currKonto = lieferantenKonten[i];

                if (currKonto != null)
                {
                    PrintKontonummer(lieferantenKonten[i].KontoNummer);

                    Console.WriteLine("\n-------------------------\n");
                }
            }

            Console.WriteLine("\n==============================\n==============================\n");

            Console.WriteLine("== KUNDENKONTEN ==");

            for (int i = 0; i < kundenKonten.Length; i++)
            {
                if (kundenKonten[i] != null)
                {
                    PrintKontonummer(kundenKonten[i].KontoNummer);

                    Console.WriteLine("\n-------------------------\n");
                }
            }

            Console.WriteLine("\n==============================\n==============================\n");
        }

        public void ErfolgskontenAbschlieszen()
        {
            for (int i = 0; i < konten.Length; i++)
            {
                int currKontennummer = konten[i].KontoNummer;

                if (IsErfolgskonto(currKontennummer))
                {
                    decimal saldo = konten[i].Saldo;

                    if (saldo > 0)
                    {
                        // Saldo im SOLL
                        konten[GetPosition(SearchFor("GuV"))].AddHabenBetrag(Math.Abs(konten[i].Saldo), $"31.12.{DateTime.Now.Year}", new int[1] { konten[i].KontoNummer });
                    }
                    else if (saldo < 0)
                    {
                        // Saldo im HABEN
                        konten[GetPosition(SearchFor("GuV"))].AddSollBetrag(Math.Abs(konten[i].Saldo), $"31.12.{DateTime.Now.Year}", new int[1] { konten[i].KontoNummer });
                    }
                }
            }

            // konten[GetPosition(9890)].Print();
        }

        public void ErfolgskontenAbschlieszen(decimal endbestand)
        {
            decimal diff = endbestand - anfangsbestand;

            if (diff > 0)
            {
                Buchungssatz endbestandUmbuchen = new($"{SearchFor("Handelswarenvorrat")} {diff} EUR / {SearchFor("Handelswareneinsatz 20%")} {diff} EUR");
                this.AddBuchung(endbestandUmbuchen);
            }
            else if (diff < 0)
            {
                Buchungssatz endbestandUmbuchen = new($"{SearchFor("Handelswareneinsatz 20%")} {Math.Abs(diff)} EUR / {SearchFor("Handelswarenvorrat")} {Math.Abs(diff)} EUR");
                this.AddBuchung(endbestandUmbuchen);
            }

            this.ErfolgskontenAbschlieszen();
        }

        public void SteuernUmbuchen(string date)
        {
            int ustKontennummer = SearchFor("Umsatzsteuer");
            int vstKontennummer = SearchFor("Vorsteuer");

            int ustZahllastKontennummer = SearchFor("Umsatzsteuer-Zahllast");

            decimal ustSaldo = konten[GetPosition(ustKontennummer)].Saldo;
            decimal vstSaldo = konten[GetPosition(vstKontennummer)].Saldo;

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
                if (konten[i].Eintraege <= 0) continue;

                if (IsErfolgskonto(konten[i].KontoNummer)) continue;
                if (konten[i].KontoNummer == schlussBilanzKonto) continue;

                decimal saldo = konten[i].Saldo;

                if (saldo > 0)
                {
                    // Saldo im SOLL
                    Buchungssatz buchungssatz = new($"{konten[i].KontoNummer} {saldo} EUR / {schlussBilanzKonto} {saldo} EUR");
                    buchungssatz.SetDatum(year, Monat.December, 31);
                    AddBuchung(buchungssatz);
                }
                else if (saldo < 0)
                {
                    // Saldo im HABEN
                    Buchungssatz buchungssatz = new($"{schlussBilanzKonto} {Math.Abs(saldo)} EUR / {konten[i].KontoNummer} {Math.Abs(saldo)} EUR");
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
                if (konten[i].KontoNummer == kontonummer)
                {
                    konten[i].Print();

                    decimal saldo = konten[i].Saldo;

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

                    for (int j = 0; j < konten[i].SollWerte.Count(); j++)
                    {
                        sollSum += konten[i].SollWerte[j];
                    }

                    for (int j = 0; j < konten[i].HabenWerte.Count(); j++)
                    {
                        habenSum += konten[i].HabenWerte[j];
                    }

                    Console.WriteLine($"Summe: {sollSum} EUR | {habenSum} EUR");

                    return;
                }
            }

            if (IsLieferantenKonto(kontonummer))
            {
                for (int i = 0; i < lieferantenKonten.Length; i++)
                {
                    if (lieferantenKonten[i] != null && lieferantenKonten[i].KontoNummer == kontonummer)
                    {
                        lieferantenKonten[i].Print();
                        decimal saldo = lieferantenKonten[i].Saldo;

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
                    if (kundenKonten[i] != null && kundenKonten[i].KontoNummer == kontonummer)
                    {
                        kundenKonten[i].Print();
                        decimal saldo = kundenKonten[i].Saldo;

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
            BinaryFormatter formatter = new BinaryFormatter();

            using (FileStream stream = new FileStream(filepath, FileMode.Create))
            {
                formatter.Serialize(stream, this);

                Console.WriteLine($"Export completed. Total bytes written:  {stream.Position.ToString("0#,0")}");
            }
        }

        public Bilanz ImportData(string filepath)
        {
            BinaryFormatter formatter = new BinaryFormatter();

            using (FileStream stream = new FileStream(filepath, FileMode.Open))
            {
                Bilanz? bilanz = formatter.Deserialize(stream) as Bilanz;
                return bilanz!;
            }
        }

        public void AddKonto(Konto konto)
        {
            int kontoNummer = konto.KontoNummer;

            foreach (Konto currKonto in konten)
            {
                if (currKonto.KontoNummer == kontoNummer)
                {
                    for (int i = 0; i < konto.SollWerte.Length; i++)
                    {
                        currKonto.AddSollBetrag(konto.SollWerte[i], konto.SollDatum[i], konto.SollGegenkonten[i].ToArray());
                    }

                    for (int i = 0; i < konto.HabenWerte.Length; i++)
                    {
                        currKonto.AddHabenBetrag(konto.HabenWerte[i], konto.HabenDatum[i], konto.HabenGegenkonten[i].ToArray());
                    }
                }
            }
        }

        public void Open()
        {
            for (int i = 0; i < konten.Length; i++)
            {
                if (SearchFor("EBK") == konten[i].KontoNummer)
                {
                    if (konten[i] == null)
                    {
                        throw new Exception("EBK-Konto wurde nicht initialisiert");
                    }

                    if (konten[i].SollWerte.Length + konten[i].HabenWerte.Length <= 0)
                    {
                        throw new Exception("EBK-Konto hat keine Eintraege");
                    }

                    decimal[] habenWerte = konten[i].HabenWerte;
                    decimal[] sollWerte = konten[i].SollWerte;

                    List<List<int>> habenGegenkonten = konten[i].HabenGegenkonten.Distinct().ToList();
                    List<List<int>> sollGegenkonten = konten[i].SollGegenkonten.Distinct().ToList();

                    List<Buchungssatz> umbuchungen = new List<Buchungssatz>();

                    for (int j = 0; j < sollGegenkonten.Count; j++)
                    {
                        Buchungssatz umbuchung = new($"{konten[i].KontoNummer} {sollWerte[j]} EUR / {sollGegenkonten[j][0]} {sollWerte[j]} EUR");
                        umbuchungen.Add(umbuchung);

                        if (sollGegenkonten[j][0] == SearchFor("Handelswarenvorrat"))
                        {
                            this.anfangsbestand -= sollWerte[j];
                        }
                    }

                    for (int j = 0; j < habenGegenkonten.Count; j++)
                    {
                        Buchungssatz umbuchung = new($"{habenGegenkonten[j][0]} {habenWerte[j]} EUR / {konten[i].KontoNummer} {habenWerte[j]} EUR");
                        umbuchungen.Add(umbuchung);

                        if (habenGegenkonten[j][0] == SearchFor("Handelswarenvorrat"))
                        {
                            this.anfangsbestand += habenWerte[j];
                        }
                    }

                    konten[i].Clear();

                    foreach (Buchungssatz umbuchung in umbuchungen)
                    {
                        this.AddBuchung(umbuchung);
                    }

                    return;
                }
            }
        }
    }
}
