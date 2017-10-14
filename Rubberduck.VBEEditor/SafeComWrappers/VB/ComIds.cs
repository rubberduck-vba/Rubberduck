using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Rubberduck.VBEditor.SafeComWrappers.VB.Abstract;
using Rubberduck.VBEditor.SafeComWrappers.VB.Enums;

namespace Rubberduck.VBEditor.SafeComWrappers.VB
{
    internal static class ComIds
    {
        private static readonly IReadOnlyDictionary<VBType, IComIds> _dictionary = new Dictionary<VBType, IComIds>
        {
            {VBType.VBA, new VBA.ComIds()},
            //{VBType.VB6, new VB6.ComIds()}
        };

        internal static IReadOnlyDictionary<VBType, IComIds> For => _dictionary;
    }
}
