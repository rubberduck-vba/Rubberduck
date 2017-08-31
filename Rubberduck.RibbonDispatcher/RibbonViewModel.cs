using System;
using System.Diagnostics.CodeAnalysis;
using System.Resources;
using System.Runtime.InteropServices;

using Microsoft.Office.Core;

using Rubberduck.RibbonDispatcher.Concrete;
using Rubberduck.RibbonDispatcher.AbstractCOM;

namespace Rubberduck.RibbonDispatcher {
    /// <summary>Implementation of (all) the callbacks for the Fluent Ribbon; for COM clients.</summary>
    /// <remarks>
    /// DOT NET clients are expected to find it more convenient to inherit their View 
    /// Model class from {AbstractDispatcher} than to compose against an instance of 
    /// {RibbonViewModel}. COM clients will most likely find the reverse true. 
    /// </remarks>
    [SuppressMessage("Microsoft.Interoperability", "CA1405:ComVisibleTypeBaseTypesShouldBeComVisible")]
    [SuppressMessage("Microsoft.Interoperability", "CA1409:ComVisibleTypesShouldBeCreatable",
        Justification ="Public Non-Creatable class for COM.")]
    [Serializable]
    [ComVisible(true)]
    [CLSCompliant(true)]
    [ComDefaultInterface(typeof(IRibbonViewModel))]
    [Guid(RubberduckGuid.RibbonViewModel)]
    public sealed class RibbonViewModel : AbstractDispatcher, IRibbonViewModel {
        /// <summary>TODO</summary>
        internal RibbonViewModel(IRibbonUI RibbonUI, IResourceManager ResourceManager) : base() 
            => InitializeRibbonFactory(RibbonUI, ResourceManager);

        /// <summary>TODO</summary>
     //   public void OnRibbonLoad(IRibbonUI RibbonUI) {;}
    }
}
