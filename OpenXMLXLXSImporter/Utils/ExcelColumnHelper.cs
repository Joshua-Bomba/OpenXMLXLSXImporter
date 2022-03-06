using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenXMLXLSXImporter.Utils
{
    public static class ExcelColumnHelper
    {
        public const uint BASE = 26U;
        public static uint GetColumnStringAsIndex(string s)
        {
            char[] chars = s.ToArray();
            uint[] charValuesAtZero = new uint[chars.Length];
            for (uint i = 0; i < chars.Length; i++)
            {
                if (chars[i] < 128 && Char.IsLetter(chars[i]))
                {
                    charValuesAtZero[i] = (uint)(Char.ToUpperInvariant(chars[i]) - 65);
                }
                else
                {
                    throw new ArgumentException("unexpected character exception");
                }
            }
            
            uint sum = 0;

            for (uint i = 0; i < charValuesAtZero.Length;i++)
            {
                sum += (uint)(charValuesAtZero[i] * Math.Floor(Math.Pow(BASE, i)));
            }

            uint log = (uint)(charValuesAtZero.Length - 1);
            sum = sum + Normalize(log);

            return sum;
        }

        public static uint Normalize(uint log)
            => (uint)(Math.Pow(BASE, log) - 1) * BASE / (BASE - 1U);

        public static string GetColumnIndexAsString(uint index)
        {
            char[] results;
            //hmm looks like were going to need to use... Math!!!

            //hmm looks like excel is not base 26 as I initally though looks like we are going to have to do extra maths
            //and by do maths I mean copy pasta

            //Grabs the log26 of index + some additional processing for the Edges ie AZ (copy pasta magic)
            uint log = (uint)Math.Floor(Math.Log(((index * (BASE - 1U)) / BASE) + 1U, BASE));
            uint digits = log + 1U;
            results = new char[digits];

            uint normalizedIndex = index - Normalize(log);

            for (uint count = 0; normalizedIndex + digits > 0; digits--)
            {
                uint digit = normalizedIndex % BASE;
                char c = (char)(digit + 65);
                results[count++] = c;
                
                normalizedIndex = normalizedIndex / BASE;
            }
            return new string(results);
            }
    }
}
