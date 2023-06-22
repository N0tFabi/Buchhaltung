using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Main
{
    public class Entry
    {
        public decimal Wert { get; set; }
        public string? Datum { get; set; }
        public List<int>? Gegenkonten { get; set; }
        public bool IsSoll { get; set; }

        public override string ToString()
        {
            return $"{Datum}: {Wert} EUR; {(IsSoll ? "Soll" : "Haben")}";
        }
    }
}
