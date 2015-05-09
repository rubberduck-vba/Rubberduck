using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Runtime.InteropServices;

namespace Rubberduck.SourceControl.Interop
{
    [ComVisible(true)]
    [Guid("423A3B28-376B-4F96-A2E0-96E354965048")]
    [ProgId("Rubberduck.Branches")]
    [ClassInterface(ClassInterfaceType.None)]
    [Description("Collection of string representation of branches in a repository.")]
    public class Branches : IEnumerable
    {
        private IEnumerable<string> branches;
        internal Branches(IEnumerable<string> branches)
        {
            this.branches = branches;
        }

        [DispId(-4)]
        public IEnumerator GetEnumerator()
        {
            return branches.GetEnumerator();
        }
    }
}
