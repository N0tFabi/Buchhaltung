using System.Text;

namespace Main
{
    public class AutoBuchungssatz : Kontenplan
    {
        public Buchungssatz WarenVerkauf(int kundenKonto, double betrag, int steuersatz)
        {
            if (Convert.ToString(kundenKonto)[0] != '2' || Convert.ToString(kundenKonto)[1] != '0')
            {
                throw new Exception("Ungueltige Kontonummer des Kunden");
            }

            int erloesKonto = 0;

            if (steuersatz == 20)
            {
                erloesKonto = 4000;
            }
            else if (steuersatz == 13)
            {
                erloesKonto = 4002;
            }
            else if (steuersatz == 10)
            {
                erloesKonto = 4001;
            }
            else
            {
                throw new Exception("Unbekannter Steuersatz");
            }

            double kundenKontoBetrag = betrag;
            double erloesKontoBetrag = (100 * betrag) / (100 + steuersatz);
            double ustKontoBetrag = kundenKontoBetrag - erloesKontoBetrag;

            
            int ustKontonummer = SearchFor("Umsatzsteuer");

            Buchungssatz buchungssatz = new($"{kundenKonto} {kundenKontoBetrag} EUR / {erloesKonto} {erloesKontoBetrag} EUR {ustKontonummer} {ustKontoBetrag} EUR");

            return buchungssatz;
        }

        public Buchungssatz WarenEinkauf (int lieferantenKonto, double betrag, int steuersatz)
        {
            if (Convert.ToString(lieferantenKonto)[0] != '3' || Convert.ToString(lieferantenKonto)[1] != '3')
            {
                throw new Exception("Ungueltige Kontonummer des Lieferanten");
            }

            int aufwandsKonto = 0;

            if (steuersatz == 20)
            {
                aufwandsKonto = 5000;
            }
            else if (steuersatz == 13)
            {
                aufwandsKonto = 5002;
            }
            else if (steuersatz == 10)
            {
                aufwandsKonto = 5001;
            }
            else
            {
                throw new Exception("Unbekannter Steuersatz");
            }

            double lieferantenKontoBetrag = betrag;
            double aufwandsKontoBetrag = (100 * betrag) / (100 + steuersatz);
            double vstKontoBetrag = lieferantenKontoBetrag - aufwandsKontoBetrag;

            
            int vstKonto = SearchFor("Vorsteuer");

            Buchungssatz buchungssatz = new($"{aufwandsKonto} {aufwandsKontoBetrag} EUR {vstKonto} {vstKontoBetrag} EUR / {lieferantenKonto} {lieferantenKontoBetrag} EUR");

            buchungssatz.Print();
            return buchungssatz;
        }

        public Buchungssatz KundenSkonto(int kundenKonto, decimal betrag, int steuersatz, int skontoSatz, int eingangsKonto)
        {
            if (Convert.ToString(kundenKonto)[0] != '2' || Convert.ToString(kundenKonto)[1] != '0')
            {
                throw new Exception("Ungueltige Kontonummer des Kunden");
            }

            
            int kundenSkontiKonto = SearchFor("Kundenskonti");
            int ustKonto = SearchFor("Umsatzsteuer");

            decimal eingangsKontoBetrag = betrag * ((100 - skontoSatz) / 100.0m);
            decimal kundenSkontiBetrag = (betrag - eingangsKontoBetrag) / (1 + steuersatz / 100.00m);
            decimal ustBetrag = betrag - eingangsKontoBetrag - kundenSkontiBetrag;

            eingangsKontoBetrag = Math.Round(eingangsKontoBetrag, 2);
            kundenSkontiBetrag = Math.Round(kundenSkontiBetrag, 2);
            ustBetrag = Math.Round(ustBetrag, 2);

            if (eingangsKontoBetrag + kundenSkontiBetrag + ustBetrag != betrag)
            {
                decimal diff = betrag - (eingangsKontoBetrag + kundenSkontiBetrag + ustBetrag);

                eingangsKontoBetrag += diff;
            }

            Buchungssatz buchungssatz = new($"{eingangsKonto} {eingangsKontoBetrag} EUR {kundenSkontiKonto} {kundenSkontiBetrag} EUR {ustKonto} {ustBetrag} EUR / {kundenKonto} {betrag} EUR");

            return buchungssatz;
        }
    
        public Buchungssatz KundenSkonto(Buchungssatz verkauf, int skontoSatz, int eingangsKonto)
        {
            

            int[] sollKonten = verkauf.GetSollKonten();
            int[] habenKonten = verkauf.GetHabenKonten();

            decimal[] sollWerte = verkauf.GetSollWerte();
            decimal[] habenWerte = verkauf.GetHabenWerte();

            decimal ustKontoBetrag = 0;
            decimal erloesKontoBetrag = 0;

            decimal betrag = verkauf.GetBetrag();

            for (int i = 0; i < Math.Max(sollKonten.Length, habenKonten.Length); i++)
            {
                if (habenKonten[i] == SearchFor("Umsatzsteuer"))
                {
                    ustKontoBetrag = habenWerte[i];
                }

                if (IsErfolgskonto(habenKonten[i]))
                {
                    erloesKontoBetrag = habenWerte[i];
                }
            }

            decimal steuersatz = (ustKontoBetrag / erloesKontoBetrag) * 100;

            decimal eingangsKontoBetrag = betrag * ((100 - skontoSatz) / 100.0m);
            decimal ustBetrag = (betrag - eingangsKontoBetrag) - (betrag - eingangsKontoBetrag) / (1 + steuersatz / 100.0m); ;
            decimal kundenSkontiBetrag = betrag - eingangsKontoBetrag - ustBetrag;


            int ustKonto = SearchFor("Umsatzsteuer");
            int kundenSkontiKonto = SearchFor("Kundenskonti");
            int kundenKonto = sollKonten[0];

            Buchungssatz buchungssatz = new($"{eingangsKonto} {eingangsKontoBetrag} EUR {kundenSkontiKonto} {kundenSkontiBetrag} EUR {ustKonto} {ustBetrag} EUR / {kundenKonto} {betrag} EUR");
            
            return buchungssatz;
        }

        public Buchungssatz LieferantenSkonto(int lieferantenKonto, decimal betrag, int steuersatz, int skontoSatz, int ausgangsKonto)
        {
            if (Convert.ToString(lieferantenKonto)[0] != '3' || Convert.ToString(lieferantenKonto)[1] != '3')
            {
                throw new Exception("Ungueltige Kontonummer des Kunden");
            }

            

            int lieferSkontiKonto = SearchFor("Lieferantenskonti");
            int vstKonto = SearchFor("Vorsteuer");

            decimal ausgangsKontoBetrag = betrag * ((100 - skontoSatz) / 100.0m);
            decimal lieferSkontiBetrag = (betrag - ausgangsKontoBetrag) / (1 + steuersatz / 100.00m);
            decimal vstBetrag = betrag - ausgangsKontoBetrag - lieferSkontiBetrag;

            ausgangsKontoBetrag = Math.Round(ausgangsKontoBetrag, 2);
            lieferSkontiBetrag = Math.Round(lieferSkontiBetrag, 2);
            vstBetrag = Math.Round(vstBetrag, 2);

            if (ausgangsKontoBetrag + lieferSkontiBetrag + vstBetrag != betrag)
            {
                decimal diff = betrag - (ausgangsKontoBetrag + lieferSkontiBetrag + vstBetrag);

                ausgangsKontoBetrag += diff;
            }

            Buchungssatz buchungssatz = new($"{lieferantenKonto} {betrag} EUR / {ausgangsKonto} {ausgangsKontoBetrag} EUR {lieferSkontiKonto} {lieferSkontiBetrag} EUR {vstKonto} {vstBetrag} EUR");

            return buchungssatz;
        }
    
        public Buchungssatz LieferantenSkonto(Buchungssatz kauf, int skontoSatz, int ausgangsKonto)
        {
            

            int[] sollKonten = kauf.GetSollKonten();
            int[] habenKonten = kauf.GetHabenKonten();

            decimal[] sollWerte = kauf.GetSollWerte();
            decimal[] habenWerte = kauf.GetHabenWerte();

            decimal vstKontoBetrag = 0;
            decimal aufwandsKontoBetrag = 0;

            decimal betrag = kauf.GetBetrag();

            for (int i = 0; i < Math.Max(sollKonten.Length, habenKonten.Length); i++)
            {
                try
                {
                    if (sollKonten[i] == SearchFor("Vorsteuer"))
                    {
                        vstKontoBetrag = sollWerte[i];
                    }
                } catch { }

                try
                {
                    if (IsErfolgskonto(sollKonten[i]))
                    {
                        aufwandsKontoBetrag = sollWerte[i];
                    }
                } catch { }
                
                
            }

            decimal steuersatz = (vstKontoBetrag / aufwandsKontoBetrag) * 100;

            decimal ausgangsKontoBetrag = betrag * ((100 - skontoSatz) / 100m);
            decimal lieferantenSkontiBetrag = (betrag - ausgangsKontoBetrag) / (1 + (steuersatz / 100m));
            decimal vstBetrag = betrag - ausgangsKontoBetrag - lieferantenSkontiBetrag;

            int vstKonto = SearchFor("Vorsteuer");
            int lieferantenSkontiKonto = SearchFor("Lieferantenskonti");
            int lieferantenKonto = habenKonten[0];

            Buchungssatz buchungssatz = new($"{lieferantenKonto} {betrag} EUR / {ausgangsKonto} {ausgangsKontoBetrag} EUR {lieferantenSkontiKonto} {lieferantenSkontiBetrag} EUR {vstKonto} {vstBetrag} EUR");

            return buchungssatz;
        }

        public Buchungssatz Storno(Buchungssatz buchungssatz)
        {
            decimal[] sollWerte = buchungssatz.GetSollWerte();
            decimal[] habenWerte = buchungssatz.GetHabenWerte();

            int[] sollKonten = buchungssatz.GetSollKonten();
            int[] habenKonten = buchungssatz.GetHabenKonten();

            StringBuilder sb = new StringBuilder();

            for (int i = 0; i < habenWerte.Length; i++)
            {
                sb.Append($"{habenKonten[i]} {habenWerte[i]} EUR ");
            }

            sb.Append("/ ");

            for (int i = 0; i < sollWerte.Length; i++)
            {
                sb.Append($"{sollKonten[i]} {sollWerte[i]} EUR ");
            }

            Buchungssatz result = new(sb.ToString());

            return result;
        }
    }
}
