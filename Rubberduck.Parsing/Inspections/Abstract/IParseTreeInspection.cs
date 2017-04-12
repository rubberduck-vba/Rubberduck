namespace Rubberduck.Parsing.Inspections.Abstract
{
    public interface IParseTreeInspection : IInspection
    {
        IInspectionListener Listener { get; }
    }
}
