using System;
using System.Diagnostics;
using System.Diagnostics.CodeAnalysis;
using System.Drawing;
using Spire;
using Spire.Xls;
using Spire.Xls.Core;

namespace Main
{
    internal class Program
    {
        static void Main(string[] args)
        {
            Kontenplan kontenplan = new();

            AutoBuchungssatz autobuchungssatz = new();

            Stopwatch stopwatch = new Stopwatch();
            stopwatch.Start();

            Konto ebk = new Konto();
            ebk.SetKontoNummer(9800);
            ebk.AddHabenBetrag(86_137, "1.1.2020", new int[] { 1600 });
            ebk.AddHabenBetrag(21_674, "1.1.2020", new int[] { 2000 });
            ebk.AddHabenBetrag(5_960, "1.1.2020", new int[] { 2700 });
            ebk.AddHabenBetrag(13_123, "1.1.2020", new int[] { 2800 });

            ebk.AddSollBetrag(102_459.80m, "1.1.2020", new int[] { 9000 });
            ebk.AddSollBetrag(23_784.20m, "1.1.2020", new int[] { 3300 });
            ebk.AddSollBetrag(650, "1.1.2020", new int[] { 3520 });

            Buchungssatz a = new("2000 28800 EUR / 4000 24000 EUR 3500 4800 EUR");
            a.SetDatum(2020, Monat.January, 01);

            Buchungssatz b = new("3300 4104 EUR / 2800 4021,92 EUR 5880 68,4 EUR 2500 13,68 EUR");
            b.SetDatum(2020, Monat.January, 03);

            Buchungssatz c = autobuchungssatz.KundenSkonto(a, 3, 2800);
            c.SetDatum(2020, Monat.January, 17);

            Buchungssatz d = new("4400 300 EUR 3500 60 EUR / 2000 360 EUR");
            d.SetDatum(2020, Monat.January, 21);

            Buchungssatz e = new("5000 1746 EUR 2500 349,2 EUR / 2700 2095,2 EUR");
            e.SetDatum(2020, Monat.January, 31);

            Bilanz bilanz = new();

            bilanz.AddKonto(ebk);

            bilanz.Open();

            bilanz.AddBuchung(a);
            bilanz.AddBuchung(b);
            bilanz.AddBuchung(c);
            bilanz.AddBuchung(d);
            bilanz.AddBuchung(e);

            bilanz.ErfolgskontenAbschlieszen(2023, 73_107m);

            bilanz.SteuernUmbuchen("31.05.2023");
            bilanz.ErfolgskontenAbschlieszen(2023);

            bilanz.CreateSchlussbilanz(2023);
            bilanz.Print();
            bilanz.PrintKontonummer(kontenplan.SearchFor("GuV"));

            bilanz.WriteToExcelFile("bilanz.xlsx");

            stopwatch.Stop();
            Console.WriteLine("\nElapsed seconds: " + stopwatch.ElapsedMilliseconds / 1000.0);
        }
    }
}
