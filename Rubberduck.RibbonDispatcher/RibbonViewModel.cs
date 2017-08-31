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
    [CLSCompliant(true)]
    internal class RibbonViewModel : AbstractDispatcher, IAbstractDispatcher {
        /// <summary>TODO</summary>
        public RibbonViewModel(IRibbonUI ribbonUI, ResourceManager resourceManager) : base() {
            InitializeRibbonFactory(ribbonUI, resourceManager);
        }

        public new IRibbonFactory RibbonFactory => base.RibbonFactory;
    }
}
