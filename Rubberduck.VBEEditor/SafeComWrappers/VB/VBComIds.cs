using System.Collections.Generic;
using Rubberduck.VBEditor.SafeComWrappers.VB.Abstract;
using Rubberduck.VBEditor.SafeComWrappers.VB.Enums;

namespace Rubberduck.VBEditor.SafeComWrappers.VB
{
    internal static class VBComIds
    {
        private static readonly IReadOnlyDictionary<VBType, IVBComIds> _dictionary = new Dictionary<VBType, IVBComIds>
        {
            {VBType.VBA, new VBA.VBComIds()},
            //{VBType.VB6, new VB6.VBComIds()}
        };

        internal static IReadOnlyDictionary<VBType, IVBComIds> For => _dictionary;
    }
}
