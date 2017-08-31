////////////////////////////////////////////////////////////////////////////////////////////////////
//                                Copyright (c) 2017 Pieter Geerkens                              //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.Resources;
using System.Runtime.InteropServices;

using Microsoft.Office.Core;

using Rubberduck.RibbonDispatcher.AbstractCOM;

namespace Rubberduck.RibbonDispatcher {

    /// <summary>TODO</summary>
    [Serializable]
    [ComVisible(true)]
   // [Guid("DF52F97E-8828-4585-834B-33DDFB5B9B82")]
    [CLSCompliant(true)]
    [ClassInterface(ClassInterfaceType.None)]
    public class Main : IMain {
        /// <summary>TODO</summary>
        public Main() { }
        /// <inheritdoc/>
        public IAbstractDispatcher NewRibbonViewModel(IRibbonUI ribbonUI, ResourceManager resourceManager)
            => new RibbonViewModel(ribbonUI, resourceManager);
    }
}
