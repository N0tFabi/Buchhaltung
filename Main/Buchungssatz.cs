using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Main
{
    public enum Monat
    {
        January = 1,
        February = 2,
        March = 3,
        April = 4,
        May = 5,
        June = 6,
        July = 7,
        August = 8,
        September = 9,
        October = 10,
        November = 11,
        December = 12
    }

    [Serializable]
    public class Buchungssatz : Kontenplan
    {
        private List<int> sollKonten = new List<int>();
        private List<int> habenKonten = new List<int>();
        private List<decimal> sollWerte = new List<decimal>();
        private List<decimal> habenWerte = new List<decimal>();

        private string buchungssatzString = "";

        private string datum = "";

        public Buchungssatz(string asText)
        {
            
            buchungssatzString = asText;

            asText = asText.Replace('.', ',');

            string[] splitted = asText.Split('/');

            for (int i = 0; i < splitted.Length; i++)
            {
                splitted[i] = splitted[i].Trim();
            }

            string soll = splitted[0];
            string haben = splitted[1];

            string[] sollSplitted = soll.Split(' ');
            string[] habenSplitted = haben.Split(' ');

            int sollElements = sollSplitted.Length;
            int habenElements = habenSplitted.Length;

            if (sollElements == 3)
            {
                // Folgendes Schema: "<Kontonummer> <Betrag> EUR"
                if (!IsValidKontennummer(Convert.ToInt32(sollSplitted[0])))
                {
                    throw new Exception("Unbekannte Kontonummer");
                }

                if (Convert.ToDecimal(sollSplitted[1]) <= 0)
                {
                    throw new Exception("Ungueltiger Betrag");
                }

                sollKonten.Add(Convert.ToInt32(sollSplitted[0]));
                sollWerte.Add(Convert.ToDecimal(sollSplitted[1]));
            }
            else if (sollElements > 3 && sollElements % 3 == 0)
            {
                // Folgendes Schema: "<Kontonummer> <Betrag> EUR <Kontonummer> <Betrag> EUR ..."

                for (int i = 0; i < sollElements; i++)
                {
                    if (i % 3 == 0 && !IsValidKontennummer(Convert.ToInt32(sollSplitted[0])))
                    {
                        throw new Exception("Unbekannte Kontonummer");
                    }

                    if (i % 3 == 1 && Convert.ToDecimal(sollSplitted[i]) <= 0)
                    {
                        throw new Exception("Ungueltiger Betrag");
                    }

                    if (i % 3 == 0)
                    {
                        sollKonten.Add(Convert.ToInt32(sollSplitted[i]));
                        //sollKonten[i / 3] = Convert.ToInt32(sollSplitted[i]);
                    }
                    else if (i % 3 == 1)
                    {
                        sollWerte.Add(Convert.ToDecimal(sollSplitted[i]));
                        //sollWerte[i / 3] = Convert.ToDecimal(sollSplitted[i]);
                    }
                }
            }

            if (habenElements == 3)
            {
                // Folgendes Schema: "<Kontonummer> <Betrag> EUR"
                if (!IsValidKontennummer(Convert.ToInt32(habenSplitted[0])))
                {
                    throw new Exception("Unbekannte Kontonummer");
                }

                if (Convert.ToDecimal(habenSplitted[1]) <= 0)
                {
                    throw new Exception("Ungueltiger Betrag");
                }

                habenKonten.Add(Convert.ToInt32(habenSplitted[0]));
                habenWerte.Add(Convert.ToDecimal(habenSplitted[1]));
            }
            else if (habenElements > 3 && habenElements % 3 == 0)
            {
                // Folgendes Schema: "<Kontonummer> <Betrag> EUR <Kontonummer> <Betrag> EUR ..."
                for (int i = 0; i < habenElements; i++)
                {
                    if (i % 3 == 0 && !IsValidKontennummer(Convert.ToInt32(habenSplitted[0])))
                    {
                        throw new Exception("Unbekannte Kontonummer");
                    }

                    if (i % 3 == 1 && Convert.ToDecimal(sollSplitted[i % 3]) <= 0)
                    {
                        throw new Exception("Ungueltiger Betrag");
                    }

                    if (i % 3 == 0)
                    {
                        habenKonten.Add(Convert.ToInt32(habenSplitted[i]));
                        //habenKonten[i / 3] = Convert.ToInt32(habenSplitted[i]);
                    }
                    else if (i % 3 == 1)
                    {
                        habenWerte.Add(Convert.ToDecimal(habenSplitted[i]));
                    }
                }
            }

            decimal sollSum = 0, habenSum = 0;

            for (int i = 0; i < sollWerte.Count; i++)
            {
                sollSum += sollWerte[i];
            }

            for (int i = 0; i < habenWerte.Count; i++)
            {
                habenSum += habenWerte[i];
            }

            if (sollSum != habenSum)
            {
                throw new Exception("Betragssumme in Soll und Haben ungleich");
            }
        }

        public string GetDate()
        {
            return this.datum;
        }

        public void SetDatum(int year, Monat month, int day)
        {
            if (year > DateTime.Now.Year)
            {
                throw new Exception("Ungueltiges Jahr");
            }

            if (day > 31 || day <= 0)
            {
                throw new Exception("Ungueltiger Tag");
            }

            int i_month = (int)month;
            this.datum = $"{day}.{i_month}.{year}";
        }

        override
        public string ToString()
        {
            return this.buchungssatzString;
        }

        public new void Print()
        {
            int maxSollKontenLength = 0, maxHabenKontenLength = 0;
            int sollKontenCount = 0, habenKontenCount = 0;
            
            for (int i = 0; i < sollKonten.Count; i++)
            {
                if (sollKonten[i] != 0)
                {
                    sollKontenCount++;
                }

                if (habenKonten[i] != 0)
                {
                    habenKontenCount++;
                }
            }

            for (int i = 0; i < sollKontenCount; i++)
            {
                StringBuilder sb = new();

                sb.Append(sollKonten[i] + " ");
                sb.Append(sollWerte[i] + " EUR ");

                if (sb.ToString().Length >= maxSollKontenLength)
                {
                    maxSollKontenLength = sb.ToString().Length;
                }
            }

            for (int i = 0; i < habenKontenCount; i++)
            {
                StringBuilder sb = new();

                sb.Append(habenKonten[i] + " ");
                sb.Append(habenWerte[i] + " EUR ");

                if (sb.ToString().Length >= maxHabenKontenLength)
                {
                    maxHabenKontenLength = sb.ToString().Length;
                }
            }

            for (int i = 0; i < sollKonten.Count; i++)
            {
                string sollMsg = $"{sollKonten[i]} {sollWerte[i]:f2} EUR ";
                if (sollWerte[i] != 0)
                {
                    Console.Write(String.Format("{0," + maxSollKontenLength + "}", sollMsg));
                }
                else if (sollWerte[i] == 0 && habenWerte[i] != 0)
                {
                    for (int j = 0; j < maxSollKontenLength; j++)
                    {
                        Console.Write(" ");
                    }
                }

                if (sollWerte[i] != 0 || habenWerte[i] != 0)
                {
                    Console.Write("/ ");
                }

                string habenMsg = $"{habenKonten[i]} {habenWerte[i]:f2} EUR ";
                if (habenWerte[i] != 0)
                {
                    Console.Write(String.Format("{0," + maxHabenKontenLength + "}", habenMsg));
                }

                if (sollWerte[i] != 0 || habenWerte[i] != 0)
                {
                    Console.WriteLine();
                }
            }
        }

        public int[] GetSollKonten()
        {
            return this.sollKonten.Where(x => x != 0).ToArray();
        }

        public int[] GetHabenKonten()
        {
            return this.habenKonten.Where(x => x != 0).ToArray();
        }

        public decimal[] GetSollWerte()
        {
            return this.sollWerte.Where(x => x != 0).ToArray();
        }

        public decimal[] GetHabenWerte()
        {
            return this.habenWerte.Where(x => x != 0).ToArray();
        }

        public decimal GetHabenWertFromKonto(int konto)
        {
            decimal result = -1;

            for (int i = 0; i < sollKonten.Count; i++)
            {
                if (sollKonten[i] == 0)
                {
                    continue;
                }

                if (sollKonten[i] == konto)
                {
                    result = sollWerte[i];
                    return result;
                }
            }

            return result;
        }

        public decimal GetSollWertFromKonto(int konto)
        {
            decimal result = -1;

            for (int i = 0; i < habenKonten.Count; i++)
            {
                if (habenKonten[i] == 0)
                {
                    continue;
                }

                if (habenKonten[i] == konto)
                {
                    result = habenWerte[i];
                    return result;
                }
            }

            return result;
        }

        public decimal GetBetrag()
        {
            decimal result = 0;

            decimal[] sollWerte = GetSollWerte();

            for (int i = 0; i < sollWerte.Length; i++)
            {
                result += sollWerte[i];
            }

            return result;
        }
    }
}
