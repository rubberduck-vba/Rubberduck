////////////////////////////////////////////////////////////////////////////////////////////////////
//                                Copyright (c) 2017 Pieter Geerkens                              //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.Resources;
using System.Runtime.InteropServices;

using Rubberduck.RibbonDispatcher.Abstract;
using Rubberduck.RibbonDispatcher.AbstractCOM;

namespace Rubberduck.RibbonDispatcher.Concrete {
    /// <summary>TODO</summary>
    [Serializable]
    [CLSCompliant(true)]
    [ClassInterface(ClassInterfaceType.None)]
    public class RibbonGroup : RibbonCommon, IRibbonGroup
    {
        internal RibbonGroup(string id, ResourceManager mgr, bool visible, bool enabled, RdControlSize size)
            : base(id, mgr, visible, enabled, size) {; }
    }
}
