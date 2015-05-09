﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.ComponentModel;

namespace Rubberduck.SourceControl
{
    [ComVisible(true)]
    [Guid("B64F14D4-D083-4B41-BE99-4736C1D24B56")]
    public interface IBranch
    {
        [DispId(0)]
        string Name { get; }

        [ComVisible(false)]
        [DispId(3)]
        string CanonicalName { get; }

        [ComVisible(false)]
        [DispId(2)]
        bool IsRemote { get; }

        [DispId(3)]
        bool IsCurrentHead { get; }
    }

    [ComVisible(true)]
    [Guid("6154532B-8880-40E9-B41E-2419C30B9F9A")]
    [ProgId("Rubberduck.Branch")]
    [ClassInterface(ClassInterfaceType.None)]
    public class Branch : IBranch
    {
        public string Name { get; private set; }
        public string CanonicalName { get; private set; }
        public bool IsRemote { get; private set; }
        public bool IsCurrentHead { get; private set; }

        public Branch(LibGit2Sharp.Branch branch)
            : this(branch.Name, branch.CanonicalName, branch.IsRemote, branch.IsCurrentRepositoryHead)
        { }

        public Branch(string name, string canonicalName, bool isRemote, bool isCurrentHead)
        {
            this.Name = name;
            this.CanonicalName = canonicalName;
            this.IsRemote = isRemote;
            this.IsCurrentHead = isCurrentHead;
        }
    }
}
