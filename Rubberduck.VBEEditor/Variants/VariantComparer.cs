using System.Globalization;

namespace Rubberduck.VBEditor.Variants
{
    public static class VariantComparer
    {
        public static VariantComparisonResults Compare(object x, object y)
        {
            return Compare(x, y, VariantComparisonFlags.NORM_IGNORECASE);
        }

        public static VariantComparisonResults Compare(object x, object y, VariantComparisonFlags flags)
        {
            return Compare(x, y, CultureInfo.InvariantCulture.LCID, flags);
        }

        public static VariantComparisonResults Compare(object x, object y, int lcid, VariantComparisonFlags flags)
        {
            object dy;
            try
            {
                dy = VariantConverter.ChangeType(y, x.GetType());
            }
            catch
            {
                dy = y;
            }
            return (VariantComparisonResults)VariantNativeMethods.VarCmp(ref x, ref dy, lcid, (uint)flags);
        }
    }
}
