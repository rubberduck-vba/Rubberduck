using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Rubberduck.AddRemoveReferences
{
    public interface IReferenceInfo
    {
        Guid Guid { get; }
        string Name { get; }
        string FullPath { get; }
        int Major { get; }
        int Minor { get; }
        int? Priority { get; }
    }
}
