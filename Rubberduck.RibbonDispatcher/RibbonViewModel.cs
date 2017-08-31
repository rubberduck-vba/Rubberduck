////////////////////////////////////////////////////////////////////////////////////////////////////
//                                Copyright (c) 2017 Pieter Geerkens                              //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.Resources;

using Microsoft.Office.Core;

using Rubberduck.RibbonDispatcher.AbstractCOM;
using Rubberduck.RibbonDispatcher.Concrete;

namespace Rubberduck.RibbonDispatcher {
    /// <summary>TODO</summary>
    [Serializable]
    //[ComVisible(true)]
    //[Guid("A3444644-112E-4971-8D84-A7C177DF0A42")]
    [CLSCompliant(true)]
    internal class RibbonViewModel : AbstractRibbonDispatcher, IAbstractDispatcher {
        /// <summary>TODO</summary>
        public RibbonViewModel(IRibbonUI ribbonUI, ResourceManager resourceManager) : base() {
            InitializeRibbonFactory(ribbonUI, resourceManager);
        }

        public new IRibbonFactory RibbonFactory => base.RibbonFactory;
    }
}
