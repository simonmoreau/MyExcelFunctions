using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MyExcelFunctions
{
    class TexteEnLettre
    {
        readonly Dictionary<long, String> dic = new Dictionary<long, string>();
        readonly Dictionary<long, String> logs = new Dictionary<long, string>();

        void init()
        {
            dic.Add(0, "zero");
            dic.Add(1, "un");
            dic.Add(2, "deux");
            dic.Add(3, "trois");
            dic.Add(4, "quatre");
            dic.Add(5, "cinq");
            dic.Add(6, "six");
            dic.Add(7, "sept");
            dic.Add(8, "huit");
            dic.Add(9, "neuf");
            dic.Add(10, "dix");
            dic.Add(11, "onze");
            dic.Add(12, "douze");
            dic.Add(13, "treize");
            dic.Add(14, "quatorze");
            dic.Add(15, "quinze");
            dic.Add(16, "seize");
            dic.Add(20, "vingt");
            dic.Add(30, "trente");
            dic.Add(40, "quarante");
            dic.Add(50, "cinquante");
            dic.Add(60, "soixante");
            dic.Add(70, "soixante-dix");
            dic.Add(80, "quatre-vingts");

            logs.Add(2, "cent"); // 10 puissance 2
            logs.Add(3, "mille"); // 10 puissance 3
            logs.Add(6, "million"); // 10 puissance 6
            logs.Add(9, "milliard"); // 10 puissance 9
        }

        public TexteEnLettre()
        {
            init();
        }

        /// <summary>
        /// Conversion de nombre entier positif en lettres (FR)
        /// </summary>
        /// <param name="nb">Le nombre (Ex:325123)</param>
        /// <returns></returns>
        public string IntToFr(long nb)
        {
            if (dic.ContainsKey(nb)) return dic[nb];

            int log = nb.ToString().Length - 1; // equivalent a log a base 10 : log = (int)Math.Floor(Math.Log10(nb));

            for (int i = log; i > 1; i--)
            {
                if (logs.ContainsKey(i))
                {
                    int master = (int)Math.Floor(Math.Pow(10, i)); // master = 1000
                    long coeff = nb / master;                       // coeff  = 3
                    long reste = nb - (coeff * master);             // reste  = 25123
                    return (coeff > 1 ? IntToFr(coeff) : "") + " " + logs[i] + (i==2 && coeff > 1 && reste == 0 ? "s" : "") + (reste > 0 ? " " + IntToFr(reste) : "");
                }
            }

            if (dic.ContainsKey(nb)) return dic[nb];

            long r = nb % 10;
            long d = nb / 10;
            long D = d * 10;
            string conjonction = (r == 1 && d != 8 && d != 9) ? " et " : "-";

            // soixante-dix
            if (d == 7)
            {
                return dic[60] + conjonction + IntToFr(nb - 60);
            }

            // quatre-vingt
            if (d == 8)
            {
                return dic[80].Trim('s') + conjonction + IntToFr(r);
            }

            // quatre-vingt dix
            if (d == 9)
            {
                return dic[80].Trim('s') + conjonction + IntToFr(nb - 80);
            }

            // autres
            if (dic.ContainsKey(D))
            {
                return dic[D] + conjonction + IntToFr(r);
            }

            return "";
        }
    }
}