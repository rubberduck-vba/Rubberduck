////////////////////////////////////////////////////////////////////////////////////////////////////
//                                Copyright (c) 2017 Pieter Geerkens                              //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.Runtime.InteropServices;
using Office = Microsoft.Office.Core;

namespace Rubberduck.RibbonDispatcher.AbstractCOM {
    /// <summary>COM-visible alias for the OFFICE enumeration <see cref="Office.RibbonControlSize"/>.</summary>
    /// <remarks>
    /// This enumeration exists because COM Interop claims that (though it should)
    /// it doesn't know about the OFFICE enumeration {RibbonControlSize}.
    /// </remarks>
    [Serializable]
    [ComVisible(true)]
    [CLSCompliant(true)]
    [Guid(RubberduckGuid.RdControlSize)]
    public enum RdControlSize {
        /// <summary>TODO</summary>
        [DispId(1)]
        rdRegular = Office.RibbonControlSize.RibbonControlSizeRegular,
        /// <summary>TODO</summary>
        [DispId(2)]
        rdLarge = Office.RibbonControlSize.RibbonControlSizeLarge
    }
}
