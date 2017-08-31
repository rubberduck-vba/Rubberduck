////////////////////////////////////////////////////////////////////////////////////////////////////
//                                Copyright (c) 2017 Pieter Geerkens                              //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Linq;
using System.Runtime.InteropServices;

using Rubberduck.RibbonDispatcher.ControlMixins;
using Rubberduck.RibbonDispatcher.AbstractCOM;

namespace Rubberduck.RibbonDispatcher.Concrete {
    /// <summary>Returns a new Ribbon DropDownViewModel instance.</summary>
    [SuppressMessage("Microsoft.Interoperability", "CA1409:ComVisibleTypesShouldBeCreatable",
      Justification = "Publc, Non-Creatable, class with exported Events.")]
    [Serializable]
    [CLSCompliant(true)]
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.None)]
    [ComSourceInterfaces(typeof(ISelectionMadeEvents))]
    [ComDefaultInterface(typeof(IRibbonDropDown))]
    [Guid(RubberduckGuid.RibbonDropDown)]
    public class RibbonDropDown : RibbonCommon, IRibbonDropDown, ISelectableMixin {
        internal RibbonDropDown(string itemId, IResourceManager mgr, bool visible, bool enabled)
            : base(itemId, mgr, visible, enabled) {}

        /// <summary>TODO</summary>
        public event SelectedEventHandler SelectionMade;

        private int                     _selectedItemIndex;
        private IList<ISelectableItem>  _items  = new List<ISelectableItem>();

        /// <summary>TODO</summary>
        public string   SelectedItemId {
            get => _items[_selectedItemIndex].Id;
            set { _selectedItemIndex = _items.IndexOf(_items.FirstOrDefault(t => t.Id==value));
                  OnActionDropDown(value, _selectedItemIndex);
                }
        }
        /// <summary>TODO</summary>
        public int      SelectedItemIndex {
            get => _selectedItemIndex;
            set => OnActionDropDown(SelectedItemId, value);
        }
        /// <summary>Call back for OnAction events from the drop-down ribbon elements.</summary>
        public void OnActionDropDown(string SelectedId, int SelectedIndex) {
            _selectedItemIndex = SelectedIndex;
            SelectionMade?.Invoke(SelectedId, SelectedIndex);
            OnChanged();
        }

        /// <summary>Returns this RibbonDropDown with a new {SelectableItem} in its list.</summary>
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification = "Matches COM usage.")]
        public IRibbonDropDown AddItem(ISelectableItem SelectableItem) {
            _items.Add(SelectableItem);
            return this;
        }

        /// <summary>TODO</summary>
        public ISelectableItem this[int ItemIndex] => _items[ItemIndex];
        /// <summary>TODO</summary>
        public ISelectableItem this[string ItemId] => ( from i in _items where i.Id == ItemId select i ).FirstOrDefault();

        /// <summary>Call back for ItemCount events from the drop-down ribbon elements.</summary>
        public int      ItemCount                => _items?.Count ?? -1;
        /// <summary>Call back for GetItemID events from the drop-down ribbon elements.</summary>
        public string   ItemId(int Index)        => _items[Index].Id;
        /// <summary>Call back for GetItemLabel events from the drop-down ribbon elements.</summary>
        public string   ItemLabel(int Index)     => _items[Index].Label;
        /// <summary>Call back for GetItemScreenTip events from the drop-down ribbon elements.</summary>
        public string   ItemScreenTip(int Index) => _items[Index].ScreenTip;
        /// <summary>Call back for GetItemSuperTip events from the drop-down ribbon elements.</summary>
        public string   ItemSuperTip(int Index)  => _items[Index].SuperTip;
        /// <summary>Call back for GetItemLabel events from the drop-down ribbon elements.</summary>
        public object   ItemImage(int Index)     => "MacroSecurity";
        /// <summary>Call back for GetItemScreenTip events from the drop-down ribbon elements.</summary>
        public bool     ItemShowImage(int Index) => _items[Index].ShowImage;
        /// <summary>Call back for GetItemSuperTip events from the drop-down ribbon elements.</summary>
        public bool     ItemShowLabel(int Index) => _items[Index].ShowImage;
    }
}
