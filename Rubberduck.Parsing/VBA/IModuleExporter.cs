using Microsoft.Vbe.Interop;

namespace Rubberduck.Parsing.VBA
{
    public interface IModuleExporter
    {
        string ExportPath { get; }

        /// <summary>
        /// Exports the specified component and returns the path to the created file.
        /// </summary>
        /// <param name="component">The module to export.</param>
        /// <returns>Returns a string containing the path and filename of the created file.</returns>
        string Export(VBComponent component);
    }
}
