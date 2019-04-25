namespace Rubberduck.Parsing.Symbols
{
    public enum Accessibility
    {
        Private = 1,
        Friend = 2,
        Implicit = 3,
        Public = 4,
        Global = 5,
        Static = 6
    }

    public static class AccessibilityExtensions
    {
        /// <summary>
        /// Gets the string/token representation of an accessibility specifier.
        /// </summary>
        /// <remarks>Implicit accessibility being unspecified, yields an empty string.</remarks>
        public static string TokenString(this Accessibility access)
        {
            return access == Accessibility.Implicit ? string.Empty : access.ToString();
        }
    }

}
