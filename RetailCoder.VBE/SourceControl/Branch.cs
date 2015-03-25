using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Rubberduck.SourceControl
{
    //todo: expose to com
    public interface IBranch
    {
        string Name { get; }
        string CanonicalName { get; }
        bool IsRemote { get; }
        bool IsCurrentHead { get; }
    }

    public class Branch : IBranch
    {
        public string Name { get; private set; }
        public string CanonicalName { get; private set; }
        public bool IsRemote { get; private set; }
        public bool IsCurrentHead { get; private set; }

        public Branch(LibGit2Sharp.Branch branch)
            : this(branch.Name, branch.CanonicalName, branch.IsRemote, branch.IsCurrentRepositoryHead)
        { }

        public Branch(string name, string friendlyName, bool isRemote, bool isCurrentHead)
        {
            this.Name = name;
            this.CanonicalName = friendlyName;
            this.IsRemote = isRemote;
            this.IsCurrentHead = isCurrentHead;
        }
    }
}
