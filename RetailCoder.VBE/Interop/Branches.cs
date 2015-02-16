using System;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Collections;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.ComponentModel;

namespace Rubberduck.Interop
{
    [ComVisible(true)]
    [Guid("423A3B28-376B-4F96-A2E0-96E354965048")]
    [ProgId("Rubberduck.Branches")]
    [ClassInterface(ClassInterfaceType.None)]
    [System.ComponentModel.Description("Collection of string representation of branches in a repository.")]
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
