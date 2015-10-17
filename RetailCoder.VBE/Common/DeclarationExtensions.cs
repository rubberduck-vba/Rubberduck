using System.Windows.Media.Imaging;
using Rubberduck.Parsing.Symbols;

namespace Rubberduck.Common
{
    public static class DeclarationExtensions
    {
        private static readonly DeclarationIconCache Cache = new DeclarationIconCache();

        public static BitmapImage BitmapImage(this Declaration declaration)
        {
            return Cache[declaration];
        }
    }
}
