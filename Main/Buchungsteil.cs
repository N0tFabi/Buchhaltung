using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Main
{
    public class Buchungsteil
    {
        private readonly int konto;
        private readonly decimal betrag;

        public Buchungsteil(int konto, decimal betrag)
        {
            this.konto = konto;
            this.betrag = betrag;
        }
    }
}
