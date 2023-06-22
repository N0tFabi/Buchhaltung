namespace Main
{
    [Serializable]
    public class Konto : Kontenplan
    {
        private List<string> sollDatum = new List<string>();
        private List<decimal> sollWerte = new List<decimal>();
        private List<List<int>> sollGegenkonten = new List<List<int>>();

        private List<string> habenDatum = new List<string>();
        private List<decimal> habenWerte = new List<decimal>();
        private List<List<int>> habenGegenkonten = new List<List<int>>();

        private string description = "";
        private int kontoNummer = 0;

        public List<List<int>> GetSollGegenkonten()
        {
            return this.sollGegenkonten;
        }

        public List<List<int>> GetHabenGegenkonten()
        {
            return this.habenGegenkonten;
        }

        public int GetEintraege()
        {
            return sollWerte.Count + habenWerte.Count;
        }

        public int GetKontoNummer()
        {
            return this.kontoNummer;
        }

        public string[] GetSollDatum()
        {
            return this.sollDatum.ToArray();
        }

        public string[] GetHabenDatum()
        {
            return this.habenDatum.ToArray();
        }

        public void SetDescription(string value)
        {
            this.description = value;

            
            kontoNummer = SearchFor(this.description);
        }

        public void SetKontoNummer(int value)
        {
            this.kontoNummer = value;

            string description = GetDescription(value);

            if (description == unknownKontennummer)
            {
                if (IsLieferantenKonto(value))
                {
                    description = $"Lieferantenkonto {value}";
                }
                else if (IsKundenKonto(value))
                {
                    description = $"Kundenkonto {value}";
                }
                else
                {
                    throw new Exception("Unbekannte Kontonummer");
                }
            }

            this.description = description;
        }

        public string GetDescription()
        {
            return this.description;
        }

        public void AddSollBetrag(decimal value, string datum, int[] gegenkonten)
        {
            sollWerte.Add(value);
            sollDatum.Add(datum);
            sollGegenkonten.Add(gegenkonten.ToList());
        }

        public void AddHabenBetrag(decimal value, string datum, int[] gegenkonten)
        {
            habenWerte.Add(value);
            habenDatum.Add(datum);
            habenGegenkonten.Add(gegenkonten.ToList());
        }

        public decimal[] GetSollWerte()
        {
            return this.sollWerte.ToArray();
        }

        public decimal[] GetHabenWerte()
        {
            return this.habenWerte.ToArray();
        }

        public decimal CalculateSaldo()
        {
            // Wenn Saldo Positiv, dann Saldo im SOLL
            // Wenn Saldo Negativ, dann Saldo im HABEN

            decimal result = 0;

            decimal sollSum = 0, habenSum = 0;

            for (int i = 0; i < sollWerte.Count; i++)
            {
                sollSum += sollWerte[i];
            }

            for (int i = 0; i < habenWerte.Count; i++)
            {
                habenSum += habenWerte[i];
            }

            if (sollSum > habenSum)
            {
                result = habenSum - sollSum;
            }
            else if (sollSum < habenSum)
            {
                result = habenSum - sollSum;
            }
            
            return result;
        }

        public int GetKontenklasse()
        {
            return Convert.ToInt32(Convert.ToString(kontoNummer)[0]) - '0';
        }

        public string[] GetSollMessage()
        {
            string[] sollMessage = new string[1_000];

            for (int i = 0; i < sollGegenkonten.Count; i++)
            {
                sollMessage[i] = $"Soll: {sollDatum[i]}  Gegenkonto: ";

                for (int j = 0; j < sollGegenkonten[i].Count; j++)
                {
                    if (sollGegenkonten[i][j] == 0)
                    {
                        continue;
                    }

                    if (j > 0)
                    {
                        sollMessage[i] += ", ";
                    }

                    sollMessage[i] += $"{sollGegenkonten[i][j]}";
                }

                sollMessage[i] += $"  Betrag: {sollWerte[i]} EUR";
            }

            return sollMessage.Where(x => !String.IsNullOrEmpty(x)).ToArray();
        }

        public string[] GetHabenMessage()
        {
            string[] habenMessage = new string[1_000];

            for (int i = 0; i < habenGegenkonten.Count; i++)
            {
                habenMessage[i] = $"Haben: {habenDatum[i]}  Gegenkonto: ";

                for (int j = 0; j < habenGegenkonten[i].Count; j++)
                {
                    if (habenGegenkonten[i][j] == 0)
                    {
                        continue;
                    }

                    if (j > 0)
                    {
                        habenMessage[i] += ", ";
                    }

                    habenMessage[i] += $"{habenGegenkonten[i][j]}";
                }

                habenMessage[i] += $"  Betrag: {habenWerte[i]} EUR";
            }

            return habenMessage.Where(x => !String.IsNullOrEmpty(x)).ToArray();
        }

        public new void Print()
        {
            Console.WriteLine($"\n == Konto {kontoNummer} ({GetDescription(kontoNummer)}) == ");

            int maxSollLength = 0, maxHabenLength = 0;

            string[] sollMessages = GetSollMessage();
            string[] habenMessages = GetHabenMessage();

            for (int i = 0; i < sollMessages.Length; i++)
            {
                if (sollMessages[i].Length > maxSollLength)
                {
                    maxSollLength = sollMessages[i].Length + 1;
                }
            }

            for (int i = 0; i < habenMessages.Length; i++)
            {
                if (habenMessages[i].Length > maxHabenLength)
                {
                    maxHabenLength = habenMessages[i].Length;
                }
            }

            string sollMessagePlaceHolder = "";

            for (int i = 0; i < maxSollLength; i++)
            {
                sollMessagePlaceHolder += ' ';
            }

            for (int i = 0; i < sollMessages.Length || i < habenMessages.Length; i++)
            {
                try
                {
                    if (kontoNummer == SearchFor("Schlussbilanzkonto"))
                    {
                        Console.Write(String.Format("{0," + maxSollLength + "}", sollMessages[i].Replace("Gegenkonto: ", "Gegenkonto Saldo: ")));
                    }
                    else
                    {
                        Console.Write(String.Format("{0," + maxSollLength + "}", sollMessages[i]));
                    }
                }
                catch
                { 
                    Console.Write(sollMessagePlaceHolder); 
                }

                Console.Write("  |  ");

                bool hasPrintedHaben = false;

                try
                {
                    if (kontoNummer == SearchFor("Schlussbilanzkonto"))
                    {
                        Console.WriteLine(String.Format("{0," + maxHabenLength + "}", habenMessages[i].Replace("Gegenkonto: ", "Gegenkonto Saldo: ")));
                    }
                    else
                    {
                        Console.WriteLine(String.Format("{0," + maxHabenLength + "}", habenMessages[i]));
                    }
                    hasPrintedHaben = true;
                }
                catch
                { 
                    hasPrintedHaben = false;
                }

                if (!hasPrintedHaben)
                {
                    Console.WriteLine();
                }
            }
        }

        public void Clear()
        {
            this.sollDatum = null!;
            this.sollWerte = null!;
            this.sollGegenkonten = null!;

            this.sollDatum = new List<string>();
            this.sollWerte = new List<decimal>();
            this.sollGegenkonten = new List<List<int>>();


            this.habenDatum = null!;
            this.habenWerte = null!;
            this.habenGegenkonten = null!;

            this.habenDatum = new List<string>();
            this.habenWerte = new List<decimal>();
            this.habenGegenkonten = new List<List<int>>();
        }

        public override string ToString()
        {
            return $"{this.kontoNummer} | {this.description}";
        }
    }
}
