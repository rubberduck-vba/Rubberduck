using Rubberduck.Resources.Inspections;

namespace Rubberduck.CodeAnalysis.Inspections.Concrete.ThunderCode
{
    internal static class ThunderCodeFormatExtension
    {
        public static string ThunderCodeFormat(this string inspectionBase, params object[] args)
        {
            return string.Format(InspectionResults.ThunderCode_Base, string.Format(inspectionBase, args));
        }
    }
}
