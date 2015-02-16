using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Collections;
using System.Runtime.InteropServices;
using System.ComponentModel;
using Rubberduck.SourceControl;

namespace Rubberduck.Interop
{
    [ComVisible(true)]
    [Guid("E68A88BB-E15C-40D8-8D18-CAF7637312B5")]
    [ProgId("Rubberduck.FileStatusEntries")]
    [ClassInterface(ClassInterfaceType.None)]
    [System.ComponentModel.Description("Collection of IFileEntries representing the status of the repository files.")]
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
