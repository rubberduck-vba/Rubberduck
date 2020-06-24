using System.Globalization;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Resources;

namespace Rubberduck.CodeAnalysis.Inspections.Extensions
{
    public static class DeclarationTypeExtensions
    {
        //ToDo: Move this to resources. (This will require moving resource lookups to Core.)
        public static string ToLocalizedString(this DeclarationType type)
        {
            return RubberduckUI.ResourceManager.GetString("DeclarationType_" + type, CultureInfo.CurrentUICulture);
        }
    }
}