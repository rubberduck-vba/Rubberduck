namespace Rubberduck.Parsing.Inspections.Abstract
{
    public enum ParsePass
    {
        AttributesPass,
        CodePanePass,
    }

    public interface IParseTreeInspection : IInspection
    {
        ParsePass Pass { get; }
        IInspectionListener Listener { get; }
    }
}
