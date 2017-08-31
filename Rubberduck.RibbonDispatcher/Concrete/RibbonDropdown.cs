////////////////////////////////////////////////////////////////////////////////////////////////////
//                                Copyright (c) 2017 Pieter Geerkens                              //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.Collections.Generic;
using System.Linq;
using System.Resources;
using System.Runtime.InteropServices;

using Rubberduck.RibbonDispatcher.Abstract;
using Rubberduck.RibbonDispatcher.AbstractCOM;
using Rubberduck.RibbonDispatcher.EventHandlers;

namespace Rubberduck.RibbonDispatcher.Concrete {
    /// <summary>Returns a new Ribbon DropDownViewModel instance.</summary>
    [Serializable]
    [CLSCompliant(true)]
    [ClassInterface(ClassInterfaceType.None)]
    [ComSourceInterfaces(typeof(ISelectionMadeEvents))]
    [ComDefaultInterface(typeof(IRibbonDropDown))]
    [Guid(RubberduckGuid.RibbonDropDown)]
    public class RibbonDropDown : RibbonCommon, IRibbonDropDown {
        internal RibbonDropDown(string id, ResourceManager mgr, bool visible, bool enabled, RdControlSize size,
                SelectionMadeEventHandler onSelectionMade, ISelectableItem[] items = null)
            : base(id, mgr, visible, enabled, size){
            if(onSelectionMade != null) SelectionMade += onSelectionMade;
            _items = items?.ToList()?.AsReadOnly();
        }

        /// <summary>TODO</summary>
        public event SelectionMadeEventHandler SelectionMade;

        private int                             _selectedItemIndex;
        private IReadOnlyList<ISelectableItem>  _items;

        /// <summary>TODO</summary>
        [DispId(DispIds.SelectedItemId)]
        public string      SelectedItemId {
            get { return _items[_selectedItemIndex].Id; }
            set { _selectedItemIndex = _items.Where((t,i) => t.Id == value).Select((t,i)=>i).FirstOrDefault();
                  OnChanged();
                }
        }
        /// <summary>TODO</summary>
        [DispId(DispIds.SelectedItemIndex)]
        public int         SelectedItemIndex {
            get { return _selectedItemIndex; }
            set { _selectedItemIndex = value; OnChanged(); }
        }
        /// <summary>Call back for OnAction events from the drop-down ribbon elements.</summary>
        [DispId(DispIds.OnActionDropDown)]
        public void OnActionDropDown(string selectedId, int selectedIndex) {
            _selectedItemIndex = selectedIndex;
            SelectionMade?.Invoke(this, new SelectionMadeEventArgs(selectedId, selectedIndex));
            OnChanged();
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
        public object   ItemImage(int index)     => "MacroSecurity"; // _items[index].Label;
        /// <summary>Call back for GetItemScreenTip events from the drop-down ribbon elements.</summary>
        [DispId(DispIds.ItemShowImage)]
        public bool     ItemShowImage(int index) => _items[index].ShowImage;
        /// <summary>Call back for GetItemSuperTip events from the drop-down ribbon elements.</summary>
        [DispId(DispIds.ItemShowLabel)]
        public bool     ItemShowLabel(int index) => _items[index].ShowImage;
    }
}
