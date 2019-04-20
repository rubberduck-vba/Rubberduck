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
        public static string CodeString(this Accessibility access)
        {
            return access == Accessibility.Implicit ? string.Empty : access.ToString();
        }
    }

}
