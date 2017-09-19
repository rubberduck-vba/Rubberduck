////////////////////////////////////////////////////////////////////////////////////////////////////
//                                Copyright (c) 2017 Pieter Geerkens                              //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.Runtime.InteropServices;

using Microsoft.Office.Core;

using Rubberduck.RibbonDispatcher.AbstractCOM;

namespace Rubberduck.RibbonDispatcher {

    /// <summary>TODO</summary>
    [Serializable]
    [ComVisible(true)]
    [CLSCompliant(true)]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(IMain))]
    [Guid(RubberduckGuid.Main)]
    [ProgId(RubberduckProgId.RibbonDispatcherProgId)]
    public class Main : IMain {
        /// <summary>TODO</summary>
        public Main() { }
        /// <inheritdoc/>
        public IRibbonViewModel NewRibbonViewModel(IRibbonUI ribbonUI, IResourceManager resourceManager)
            => new RibbonViewModel(ribbonUI, resourceManager);
    }
}
