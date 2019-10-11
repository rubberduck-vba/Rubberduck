namespace Rubberduck.Common
{
    /// <summary>
    /// Allows exporting to a variety of formats
    /// </summary>
    public interface IExportable
    {
        /// <summary>
        /// Exports the object as an array of members
        /// </summary>
        /// <returns></returns>
        object[] ToArray();

        /// <summary>
        /// Exports the object as a formatted string
        /// </summary>
        /// <returns></returns>
        string ToClipboardString();
    }
}