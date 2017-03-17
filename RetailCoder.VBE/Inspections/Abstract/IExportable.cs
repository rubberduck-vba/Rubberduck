namespace Rubberduck.Inspections.Abstract
{
    public interface IExportable
    {
        object[] ToArray();
        string ToClipboardString();
    }
}