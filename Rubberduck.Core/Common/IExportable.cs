namespace Rubberduck.Common
{
    public interface IExportable
    {
        object[] ToArray();
        string ToClipboardString();
    }
}