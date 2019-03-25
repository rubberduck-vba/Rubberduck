using Rubberduck.Resources.Inspections;

namespace Rubberduck.Inspections.Inspections.Concrete.ThunderCode
{
    public static class ThunderCodeFormatExtension
    {
        public static string ThunderCodeFormat(this string inspectionBase, params object[] args)
        {
            return string.Format(InspectionResults.ThunderCode_Base, string.Format(inspectionBase, args));
        }
    }
}
