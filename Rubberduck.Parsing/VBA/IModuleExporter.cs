using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.Parsing.VBA
{
    public interface IModuleExporter
    {
        string ExportPath { get; }
        bool TempFile { get; }

        /// <summary>
        /// Exports the specified component and returns the path to the created file.
        /// </summary>
        /// <param name="component">The module to export.</param>
        /// <param name="tempFile">True if a unique temp file name should be generated.</param>
        /// <returns>Returns a string containing the path and filename of the created file.</returns>
        string Export(IVBComponent component, bool tempFile = false);
    }
}
