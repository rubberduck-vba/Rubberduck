using System.Globalization;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Resources;

namespace Rubberduck.CodeAnalysis.Inspections.Extensions
{
    internal static class DeclarationTypeExtensions
    {
        public static string ToLocalizedString(this DeclarationType type)
        {
            return RubberduckUI.ResourceManager.GetString("DeclarationType_" + type, CultureInfo.CurrentUICulture);
        }
    }
}