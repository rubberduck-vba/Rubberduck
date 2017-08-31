using System;
using System.Runtime.InteropServices;

using Rubberduck.RibbonDispatcher.AbstractCOM;

namespace Rubberduck.RibbonDispatcher.ControlDecorators {
    /// <summary>The interface for controls that can be sized.</summary>
    [CLSCompliant(true)]
    public interface ISizeableMixin {
        /// <summary>TODO</summary>
        [DispId(DispIds.Size)]
        RdControlSize Size { get; set; }
    }
}
