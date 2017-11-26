using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;

namespace Rubberduck.UI
{
    public static class HotkeyDisplayConverter
    {
        private static readonly List<Tuple<Keys, string>> _keys = new List<Tuple<Keys, string>>
        {
            Tuple.Create(Keys.D0, "0"),
            Tuple.Create(Keys.D1, "1"),
            Tuple.Create(Keys.D2, "2"),
            Tuple.Create(Keys.D3, "3"),
            Tuple.Create(Keys.D4, "4"),
            Tuple.Create(Keys.D5, "5"),
            Tuple.Create(Keys.D6, "6"),
            Tuple.Create(Keys.D7, "7"),
            Tuple.Create(Keys.D8, "8"),
            Tuple.Create(Keys.D9, "9")
        };

        public static string Convert(Keys value)
        {
            var tuple = _keys.SingleOrDefault(k => k.Item1 == value);
            return tuple == null ? value.ToString() : tuple.Item2;
        }

        public static string Convert(string value)
        {
            var tuple = _keys.SingleOrDefault(k => k.Item1.ToString() == value);
            return tuple == null ? value : tuple.Item2;
        }

        public static string ConvertBack(string value)
        {
            var tuple = _keys.SingleOrDefault(k => k.Item2 == value);
            return tuple == null ? value : tuple.Item1.ToString();
        }
    }
}
