using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Main
{
    [Serializable]
    public class Kontenplan
    {
        public readonly string unknownKontennummer = "Unbekannt";

        public readonly int LIEFERANTEN_KONTO_MIN = 33000;
        public readonly int LIEFERANTEN_KONTO_MAX = 33999;

        public readonly int KUNDEN_KONTO_MIN = 20000;
        public readonly int KUNDEN_KONTO_MAX = 20999;

        readonly string[,] kontenplan = new string[,]
        {
            // Klasse 0
            {"0110", "Patentrechte, Lizenzen"},
            {"0120", "Datenverarbitungsprogramme"},
            {"0200", "Unbebaute Grundstücke"},
            {"0210", "bebaute Grundstücke"},
            {"0300", "Gebäude"},
            {"0400", "Maschinen"},
            {"0450", "Sonstige Betriebsanlagen"},
            {"0500", "Werkzeuge"},
            {"0600", "Betriebs- und Geschäftsausstattung"},
            {"0620", "Büromaschinen, EDV-Anlagen"},
            {"0630", "Pkw und Kombis"},
            {"0640", "Lkw"},
            {"0800", "Beteiligungen"},

            // Klasse 1
            {"1000", "Bezugsverrechnung"},
            {"1100", "Rohstoffvorrat"},
            {"1250", "Vorrat Ersatzteile"},
            {"1300", "Hilfsstoffvorrat"},
            {"1340", "Vorrat Verpackungsmaterial"},
            {"1350", "Vorrat Betriebsstoffe"},
            {"1360", "Vorrat Heizöl"},
            {"1365", "Vorrat Schmiermittel"},
            {"1370", "Vorrat Reinigungsmaterial"},
            {"1390", "Vorrat Büromaterial"},
            {"1400", "Unfertige Erzeugnisse"},
            {"1500", "Fertige Erzeugnisse"},
            {"1600", "Handelswarenvorrat"},

            // Klasse 2
            {"2000", "Lieferforderungen"},
            {"2300", "Sonstige Forderungen"},
            {"2320", "Gegebene Darlehen"},
            {"2380", "Guthaben bei Lieferanten"},
            {"2400", "Lohn- und Gehaltsvorschüsse"},
            {"2500", "Vorsteuer"},
            {"2700", "Kassa"},
            {"2800", "Bank"},
            {"2870", "Barverkehr mit Banken"},
            {"2880", "Forderungen Kreditkartenunternehmen"},
            {"2885", "Forderungen Debitkarten"},
            {"2890", "Schwebende Geldbewegungen"},

            // Klasse 3
            {"3120", "Bank"},
            {"3150", "Darlehen"},
            {"3180", "Verbindlichkeiten Kreditkartenunternehmen"},
            {"3190", "Verbindlichkeiten Debitkarten"},
            {"3300", "Lieferverbindlichkeiten"},
            {"3500", "Umsatzsteuer"},
            {"3520", "Umsatzsteuer-Zahllast"},
            {"3540", "Verbindlichkeiten Finanzamt"},
            {"3600", "Verbindlichkeiten Gesundheitskasse"},
            {"3610", "Verbindlichkeiten Gemeinde"},
            {"3800", "Sonstige Verbindlichkeiten"},
            {"3850", "Verbindlichkeiten gegen Mitarbeiter aus der Bezugsverrechnung"},

            // Klasse 4
            {"4000", "Handelswarenerlöse 20%"},
            {"4001", "Handelswarenerlöse 10%"},
            {"4002", "Handelswarenerlose 13%"},
            {"4100", "Fertigerzeugniserlöse"},
            {"4400", "Erlösberichtigungen"},
            {"4410", "Kundenskonti"},
            {"4500", "Bestandsveränderungen"},
            {"4810", "Mieterträge (Pachterträge)"},
            {"4820", "Provisionserträge"},
            {"4880", "Übrige betriebliche Erträge"},
            {"4890", "Mahnspesenvergütungen"},
            {"4900", "Eigenverbrauch 20%"},
            {"4901", "Eigenverbrauch 10%"},
            {"4902", "Eigenverbrauch 13%"},
            {"4903", "Eigenverbrauch 0%"},

            // Klasse 5
            {"5000", "Handelswareneinsatz 20%"},
            {"5001", "Handelswareneinsatz 10%"},
            {"5002", "Handelswareneinsatz 13%"},
            {"5100", "Rohstoffverbrauch (-einsatz)"},
            {"5300", "Hilfsstoffverbrauch"},
            {"5340", "Verpackungsmaterialverbrauch"},
            {"5400", "Betriebsstoffverbrauch"},
            {"5410", "Schmiermittelverbrauch"},
            {"5420", "Reparaturmaterialverbrauch"},
            {"5450", "Reinigungsmaterialverbrauch"},
            {"5500", "Verbrauch von Werkzeugen, Erzeugungshilfsmitteln"},
            {"5510", "Ersatzteileverbrauch"},
            {"5600", "Bezugskosten"},
            {"5880", "Lieferantenskonti auf Wareneinkauf (Materialaufwand)"},
            {"5890", "Umsatzbonus auf Wareneinkauf"},

            // Klasse 6
            {"6000", "Löhne"  },
            {"6100", "Lehrlingsentschädigungen Arbeiter"},
            {"6200", "Gehälter"},
            {"6300", "Lehrlingsentschädigungen Angestellte"},
            {"6440", "Aufwand für Betriebliche Vorsorgekassen Arbeiter"},
            {"6441", "Aufwand für Betriebliche Vorsorgekassen Angestellte"},
            {"6500", "Gesetzlicher Sozialaufwand Arbeiter"},
            {"6560", "Gesetzlicher Sozialaufwand Angestellte"},
            {"6600", "Dienstgeberbeitrag Arbeiter"},
            {"6610", "Zuschlag zum DB Arbeiter"},
            {"6620", "Kommunalsteuer Arbeiter"  },
            {"6630", "Wiener Dienstgeberabgabe Arbeiter"},
            {"6630", "Wiener Dienstgeberabgabe Arbeiter"},
            {"6660", "Dienstgeberbeitrag Angestellte"},
            {"6670", "Zuschlag zum DB Angestellte"},
            {"6680", "Kommunalsteuer Angestellte"},
            {"6690", "Wiener Dienstgeberabgabe Angestellte"},
            {"6700", "Freiwilliger Sozialaufwand"},

            // Klasse 7
            {"7100", "Grundsteuer"},
            {"7150", "Tourismusabgabe (Interessentenbeitragl"},
            {"7180", "Gebühren"},
            {"7190", "Sonstige Abgaben"},
            {"7200", "Instandhaltung durch Dritte"},
            {"7210", "Reinigung durch Dritte"},
            {"7220", "Entsorgungsaufwand"},
            {"7240", "Heizölverbrauch"},
            {"7250", "Treibstoffverbrauch"},
            {"7260", "Gasverbrauch"},
            {"7270", "Stromverbrauch"},
            {"7280", "Heizmaterialverbrauch (feste Brennstoffe)"},
            {"7300", "Ausgangsfrachten (Transporte durch Dritte)"},
            {"7310", "Paketgebühren 20%"},
            {"7311", "Paketgebühren 0%"},
            {"7320", "Pkw-und Kombi-Betriebsaufwand"},
            {"7321", "Motorbezogene Versicherungssteuer Pkw und Kombis"},
            {"7325", "Versicherungsaufwand Pkw und Kombis"},
            {"7326", "Parkgebühren, Straßenmaut Pkw und Kombis"},
            {"7330", "Lkw-Betriebsaufwand"},
            {"7331", "Motorbezogene Versicherungssteuer Lkw"},
            {"7332", "Kraftfahrzeugsteuer Lkw"},
            {"7335", "Versicherungsaufwand Lkw"},
            {"7336", "Parkgebühren, Straßenmaut Lkw"},
            {"7380", "Telefon- und Internetgebühren"},
            {"7385", "Portogebühren"},
            {"7400", "Mietaufwand (Pachtaufwand)"},
            {"7540", "Provisionen an Dritte (Nicht-Arbeitnehmer)"},
            {"7600", "Büromaterial (Büroaufwand, Bürobedarf)"},
            {"7610", "Kopien und sonstige Druckkosten"},
            {"7630", "Fachliteratur und Zeitungen"},
            {"7650", "Werbeaufwand"},
            {"7680", "Spenden und Trinkgelder"},
            {"7700", "Versicherungsaufwand"},
            {"7740", "Versicherungsbeiträge an die Sozialversicherungs- anstalt der gewerblichen Wirtschaft"},
            {"7750", "Rechts- und Beratungsaufwand"},
            {"7770", "Aus- und Fortbildung"},
            {"7780", "Kammerumlage"},
            {"7790", "Spesen des Geldverkehrs"},
            {"7791", "Sonstige Bankspesen"},
            {"7792", "Provisionen, Gebühren Kredit- und Debitkarten"},
            {"7811", "Konventionalstrafen"},
            {"7819", "Sonstige Schadensfälle"},
            {"7850", "Übrige betriebliche Aufwendungen"},
            {"7880", "Lieferantenskonti auf sonstige betriebliche Aufwendungen"},
            
            // Klasse 8
            {"8050", "Zinsenerträge aus Bankguthaben"},
            {"8051", "Zinsenerträge aus gewährten Darlehen"},
            {"8055", "Verzugszinsenerträge"},
            {"8056", "Sonstige Zinsenerträge 20%"},
            {"8057", "Sonstige Zinsenerträge 10%"},
            {"8310", "Zinsenaufwand für Bankkredite"},
            {"8311", "Sonstiger Aufwand für Bankkredite (z. B. Bereit- stellungsprovision, Überziehungsprovision)"},
            {"8315", "Zinsenaufwand für Darlehen"},
            {"8320", "Verzugszinsenaufwand"},
            {"8321", "Mahnspesen"},
            {"8325", "Zinsenaufwand für Lieferantenkredite 20%"},
            {"8326", "Zinsenaufwand für Lieferantenkredite 10%"},
            
            // Klasse 9
            {"9000", "Kapital"},
            {"9600", "Privat"},
            {"9610", "Privatsteuern"},
            {"9800", "Eröffnungsbilanzkonto (EBK)"},
            {"9850", "Schlussbilanzkonto (SBK)"},
            {"9890", "Gewinn- und Verlustkonto (GuV)"}
        };

        public int GetKontenklasse(Konto konto)
        {
            string nummer = Convert.ToString(konto.GetKontoNummer());

            return Convert.ToInt32(nummer[0]);
        }

        public string GetDescription(int kontennummer)
        {
            string result = unknownKontennummer;
            bool hasFound = false;

            for (int i = 0; i < kontenplan.GetLength(0) && !hasFound; i++)
            {
                int currentKontennummer = int.Parse(kontenplan[i, 0]);

                if (currentKontennummer == kontennummer)
                {
                    hasFound = true;
                    result = kontenplan[i, 1];
                }
            }

            if (!hasFound && (IsLieferantenKonto(kontennummer) || IsKundenKonto(kontennummer)))
            {
                result = $"Lieferantenkonto {kontennummer}";
            }

            return result;
        }

        public int SearchFor(string text)
        {
            // gibt KONTONUMMER, nicht index zurück        KONTONUMMER

            int result = -1;
            bool hasFound = false;

            for (int i = 0; i < kontenplan.GetLength(0) && !hasFound; i++)
            {
                string currentDescription = kontenplan[i, 1];

                if (currentDescription.ToUpper().Contains(text.ToUpper()))
                {
                    int currentKontennummer = int.Parse(kontenplan[i, 0]);
                    result = currentKontennummer;

                    hasFound = true;
                }
            }

            if (text.Contains(' ') && !hasFound && (text.ToLower().Contains("lieferkonto") || text.ToLower().Contains("lieferantenkonto") || text.ToLower().Contains("kundenkonto")))
            {
                string[] textParts = text.Split(' ');

                for (int i = 0; i < textParts.Length; i++)
                {
                    int kontennummer = 0;
                    bool parseSucces = int.TryParse(textParts[i], out kontennummer);

                    if (parseSucces)
                    {
                        result = kontennummer;
                    }
                }
            }

            return result;
        }

        public bool IsValidKontennummer(int number)
        {
            bool result = IsLieferantenKonto(number) || IsKundenKonto(number);

            for (int i = 0; i < kontenplan.GetLength(0) && !result; i++)
            {
                result = number == int.Parse(kontenplan[i, 0]);
            }

            return result;
        }

        public int Length()
        {
            // EXKLUSIV Kunden- und Lieferantenkonten
            return kontenplan.GetLength(0);
        }

        public int ErfolgskontenLength()
        {
            int result = 0;

            for (int i = 0; i < kontenplan.GetLength(0); i++)
            {
                if ('4' <= kontenplan[i, 0][0] && kontenplan[i, 0][0] <= '8')
                {
                    result++;
                }
            }

            return result;
        }

        public int[] GetErfolgsKonten()
        {
            int[] result = new int[ErfolgskontenLength()];
            int indexCounter = 0;

            for (int i = 0; i < kontenplan.GetLength(0); i++)
            {
                if ('4' <= kontenplan[i, 0][0] && kontenplan[i, 0][0] <= '8')
                {
                    result[indexCounter] = Convert.ToInt32(kontenplan[i, 0]);
                    indexCounter++;
                }
            }

            return result;
        }

        public int GetPosition(int kontonummer)
        {
            // gibt INDEX vom konto im konten-array zurück,     INDEX

            int result = -1;
            bool hasFoundPosition = false;

            for (int i = 0; i < kontenplan.GetLength(0) && !hasFoundPosition; i++)
            {
                int currentKontonummer = int.Parse(kontenplan[i, 0]);

                if (currentKontonummer == kontonummer)
                {
                    result = i;
                    hasFoundPosition = true;
                }
            }

            return result;
        }

        public int GetKontonummer(int position)
        {
            // funktioniert mit INDEX

            int result = -1;

            if (position >= 0 && position < kontenplan.GetLength(0))
            {
                result = int.Parse(kontenplan[position, 0]);
            }

            return result;
        }

        public void Print()
        {
            Console.WriteLine($"---------- Klasse 0 ----------");
            Console.WriteLine($"{kontenplan[0, 0]}: {kontenplan[0, 1]}");
            for (int i = 1; i < kontenplan.GetLength(0); i++)
            {
                if (kontenplan[i - 1, 0][0] != kontenplan[i, 0][0])
                {
                    Console.WriteLine($"---------- Klasse {kontenplan[i, 0][0]} ----------");
                }
                Console.WriteLine($"{kontenplan[i, 0]}: {kontenplan[i, 1]}");
            }
        }

        public int[] GetAufwandsKonten()
        {
            Kontenplan kontenplan = new();
            int aufwandskontenCurrIndex = 0;
            int aufwandskontenLength = 0;

            for (int i = 0; i < kontenplan.Length(); i++)
            {
                string currKonto = $"{Convert.ToString(kontenplan.GetKontonummer(i)),4}";

                if (('5' <= currKonto[0] && currKonto[0] <= '7') || (currKonto[0] == '8' && currKonto[1] != '0'))
                {
                    aufwandskontenLength++;
                }
            }

            int[] aufwandskonten = new int[aufwandskontenLength];

            for (int i = 0; i < kontenplan.Length(); i++)
            {
                string currKonto = $"{Convert.ToString(kontenplan.GetKontonummer(i)),4}";

                if (('5' <= currKonto[0] && currKonto[0] <= '7') || (currKonto[0] == '8' && currKonto[1] != '0'))
                {
                    aufwandskonten[aufwandskontenCurrIndex] = int.Parse(currKonto);
                    aufwandskontenCurrIndex++;
                }
            }

            return aufwandskonten;
        }

        public int[] GetErtragsKonten()
        {
            Kontenplan kontenplan = new();
            int ertragskontenCurrIndex = 0;
            int ertragskontenLength = 0;

            for (int i = 0; i < kontenplan.Length(); i++)
            {
                string currKonto = $"{Convert.ToString(kontenplan.GetKontonummer(i)),4}";

                if (currKonto[0] == '4' || (currKonto[0] == '8' && currKonto[1] == '0'))
                {
                    ertragskontenLength++;
                }
            }

            int[] ertragskonten = new int[ertragskontenLength];

            for (int i = 0; i < kontenplan.Length(); i++)
            {
                string currKonto = $"{Convert.ToString(kontenplan.GetKontonummer(i)),4}";

                if (currKonto[0] == '4' || (currKonto[0] == '8' && currKonto[1] == '0'))
                {
                    ertragskonten[ertragskontenCurrIndex] = int.Parse(currKonto);
                    ertragskontenCurrIndex++;
                }
            }

            return ertragskonten;
        }

        public bool IsErfolgskonto(int kontennummer)
        {
            bool result = false;
            string kontennummerToString = $"{Convert.ToString(kontennummer),4}";

            if ('4' <= kontennummerToString[0] && kontennummerToString[0] <= '8')
            {
                result = true;
            }

            return result;
        }

        public bool IsErloesKonto(int kontennummer)
        {
            string kontennummerStr = Convert.ToString(kontennummer);

            if (kontennummerStr[0] == '4' || (kontennummerStr[0] == '8' && kontennummerStr[1] == '0'))
            {
                return true;
            }

            return false;
        }

        public bool IsAufwandsKonto(int kontennummer)
        {
            string kontennummerStr = Convert.ToString(kontennummer);

            if (kontennummerStr[0] == '4' || (kontennummerStr[0] == '8' && kontennummerStr[1] == '0'))
            {
                return true;
            }

            return false;
        }

        public bool IsLieferantenKonto(int kontennummer)
        {
            return kontennummer >= LIEFERANTEN_KONTO_MIN && kontennummer <= LIEFERANTEN_KONTO_MAX;
        }

        public bool IsKundenKonto(int kontennummer)
        {
            return KUNDEN_KONTO_MIN <= kontennummer && kontennummer <= KUNDEN_KONTO_MAX;
        }

        public int GetKundenKontenAmount()
        {
            int result = 0;

            for (int i = KUNDEN_KONTO_MIN; i < KUNDEN_KONTO_MAX; i++)
            {
                result++;
            }

            return result;
        }

        public int GetLieferantenKontenAmount()
        {
            int result = 0;

            for (int i = LIEFERANTEN_KONTO_MIN; i < LIEFERANTEN_KONTO_MAX; i++)
            {
                result++;
            }

            return result;
        }
    }
}
