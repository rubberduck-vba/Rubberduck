using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Runtime.InteropServices;

namespace Rubberduck.SourceControl.Interop
{
    [ComVisible(true)]
    [Guid("E68A88BB-E15C-40D8-8D18-CAF7637312B5")]
    [ProgId("Rubberduck.FileStatusEntries")]
    [ClassInterface(ClassInterfaceType.None)]
    [Description("Collection of IFileEntries representing the status of the repository files.")]
    public class FileStatusEntries : IEnumerable
    {
        private IEnumerable<IFileStatusEntry> entries;
        public FileStatusEntries(IEnumerable<IFileStatusEntry> entries)
        {
            this.entries = entries;
        }

        [DispId(-4)]
        public IEnumerator GetEnumerator()
        {
            return entries.GetEnumerator();
        }
    }
}
