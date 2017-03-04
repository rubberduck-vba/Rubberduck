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
        private readonly IEnumerable<IBranch> _branches;
        internal Branches(IEnumerable<IBranch> branches)
        {
            _branches = branches;
        }

        [DispId(-4)]
        public IEnumerator GetEnumerator()
        {
            return _branches.GetEnumerator();
        }
    }
}
