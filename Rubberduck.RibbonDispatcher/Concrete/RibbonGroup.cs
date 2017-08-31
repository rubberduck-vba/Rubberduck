////////////////////////////////////////////////////////////////////////////////////////////////////
//                                Copyright (c) 2017 Pieter Geerkens                              //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.Diagnostics.CodeAnalysis;
using System.Resources;
using System.Runtime.InteropServices;

using Rubberduck.RibbonDispatcher.AbstractCOM;

namespace Rubberduck.RibbonDispatcher.Concrete {
    /// <summary>TODO</summary>
    [SuppressMessage("Microsoft.Interoperability", "CA1409:ComVisibleTypesShouldBeCreatable",
       Justification = "Publc, Non-Creatable class.")]
    [Serializable]
    [CLSCompliant(true)]
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(IRibbonGroup))]
    [Guid(RubberduckGuid.RibbonGroup)]
    public class RibbonGroup : RibbonCommon, IRibbonGroup
    {
        internal RibbonGroup(string id, ResourceManager mgr, bool visible, bool enabled, RdControlSize size)
            : base(id, mgr, visible, enabled, size) {; }
    }
}
