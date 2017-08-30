using System;
using System.Runtime.InteropServices;
using Microsoft.Office.Core;

namespace Rubberduck.RibbonDispatcher.Abstract {
    /// <summary>COM-compatible alias for the OFFICE enumeration <see cref="RibbonControlSize"/>.</summary>
    /// <remarks>
    /// This enumeration exists because COM Interop claims that (though it should)
    /// it doesn't know about the OFFICE enumeration {RibbonControlSize}.
    /// </remarks>
    [Serializable]
    [ComVisible(true)]
    [Guid("844043EF-6B1B-471D-BC67-BD2CB3A8E7E4")]
    [CLSCompliant(true)]
    public enum MyRibbonControlSize {
        /// <summary>TODO</summary>
        Regular = RibbonControlSize.RibbonControlSizeRegular,
        /// <summary>TODO</summary>
        Large   = RibbonControlSize.RibbonControlSizeLarge
    }
}
