using System.Globalization;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Resources;

namespace Rubberduck.Inspections.Inspections.Extensions
{
    public static class DeclarationTypeExtensions
    {
        public static string ToLocalizedString(this DeclarationType type)
        {
            return RubberduckUI.ResourceManager.GetString("DeclarationType_" + type, CultureInfo.CurrentUICulture);
        }
    }
}