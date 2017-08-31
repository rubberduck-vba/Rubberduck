////////////////////////////////////////////////////////////////////////////////////////////////////
//                                Copyright (c) 2017 Pieter Geerkens                              //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Linq;
using System.Runtime.InteropServices;
using stdole;

using Rubberduck.RibbonDispatcher.ControlMixins;
using Rubberduck.RibbonDispatcher.AbstractCOM;

namespace Rubberduck.RibbonDispatcher.Concrete {
    /// <summary>Returns a new Ribbon DropDownViewModel instance.</summary>
    [SuppressMessage("Microsoft.Interoperability", "CA1409:ComVisibleTypesShouldBeCreatable",
      Justification = "Publc, Non-Creatable class with exported Events.")]
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
        [DispId(DispIds.SelectedItemId)]
        public string   SelectedItemId {
            get => _items[_selectedItemIndex].Id;
            set { _selectedItemIndex = _items.IndexOf(_items.FirstOrDefault(t => t.Id==value));
                  OnActionDropDown(value, _selectedItemIndex);
                }
        }
        /// <summary>TODO</summary>
        [DispId(DispIds.SelectedItemIndex)]
        public int      SelectedItemIndex {
            get => _selectedItemIndex;
            set => OnActionDropDown(SelectedItemId, value);
        }
        /// <summary>Call back for OnAction events from the drop-down ribbon elements.</summary>
        [DispId(DispIds.OnActionDropDown)]
        public void OnActionDropDown(string selectedId, int selectedIndex) {
            _selectedItemIndex = selectedIndex;
            SelectionMade?.Invoke(selectedId, selectedIndex);
            OnChanged();
        }

        /// <summary>Returns this RibbonDropDown with a new {SelectableItem} in its list.</summary>
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification = "Matches COM usage.")]
        public IRibbonDropDown AddItem(ISelectableItem selectableItem) {
            _items.Add(selectableItem);
            return this;
        }

        /// <summary>TODO</summary>
        public ISelectableItem this[int itemIndex] => _items[itemIndex];
        /// <summary>TODO</summary>
        public ISelectableItem this[string itemId] => ( from i in _items where i.Id == itemId select i ).FirstOrDefault();

        /// <summary>Call back for ItemCount events from the drop-down ribbon elements.</summary>
        [DispId(DispIds.ItemCount)]
        public int      ItemCount                => _items?.Count ?? -1;
        /// <summary>Call back for GetItemID events from the drop-down ribbon elements.</summary>
        [DispId(DispIds.ItemId)]
        public string   ItemId(int index)        => _items[index].Id;
        /// <summary>Call back for GetItemLabel events from the drop-down ribbon elements.</summary>
        [DispId(DispIds.ItemLabel)]
        public string   ItemLabel(int index)     => _items[index].Label;
        /// <summary>Call back for GetItemScreenTip events from the drop-down ribbon elements.</summary>
        [DispId(DispIds.ItemScreenTip)]
        public string   ItemScreenTip(int index) => _items[index].ScreenTip;
        /// <summary>Call back for GetItemSuperTip events from the drop-down ribbon elements.</summary>
        [DispId(DispIds.ItemSuperTip)]
        public string   ItemSuperTip(int index)  => _items[index].SuperTip;
        /// <summary>Call back for GetItemLabel events from the drop-down ribbon elements.</summary>
        [DispId(DispIds.ItemImage)]
        public object   ItemImage(int index)     => "MacroSecurity";
        /// <summary>Call back for GetItemScreenTip events from the drop-down ribbon elements.</summary>
        [DispId(DispIds.ItemShowImage)]
        public bool     ItemShowImage(int index) => _items[index].ShowImage;
        /// <summary>Call back for GetItemSuperTip events from the drop-down ribbon elements.</summary>
        [DispId(DispIds.ItemShowLabel)]
        public bool     ItemShowLabel(int index) => _items[index].ShowImage;
    }
}
