using System;
using System.Resources;
using System.Runtime.InteropServices;

using Rubberduck.RibbonDispatcher.Abstract;
using Rubberduck.RibbonDispatcher.EventHandlers;

namespace Rubberduck.RibbonDispatcher.Concrete {
    /// <summary>Returns a new Ribbon DropDownViewModel instance.</summary>
    [Serializable]
    [CLSCompliant(true)]
    [ClassInterface(ClassInterfaceType.None)]
    [ComSourceInterfaces(typeof(ISelectionMadeEvents))]
    public class RibbonDropDown : RibbonCommon, IRibbonDropDown {
        internal RibbonDropDown(string id, ResourceManager mgr, bool visible, bool enabled, MyRibbonControlSize size)
            : base(id, mgr, visible, enabled, size){
        }

        /// <summary>TODO</summary>
        public event SelectionMadeEventHandler SelectionMade;

        /// <summary>TODO</summary>
        public string SelectedItemId { get; set; }

        /// <summary>TODO</summary>
        public void OnAction(string itemId) => SelectionMade?.Invoke(this, new SelectionMadeEventArgs(itemId));

        /// <summary>TODO</summary>
        public IRibbonCommon AsRibbonControl => this;
    }
}
