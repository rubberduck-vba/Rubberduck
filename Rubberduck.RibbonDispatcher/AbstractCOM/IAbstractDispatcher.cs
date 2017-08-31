using System;
using System.Runtime.InteropServices;

using Microsoft.Office.Core;

namespace Rubberduck.RibbonDispatcher.AbstractCOM {
    /// <summary>TODO</summary>
    [ComVisible(true)]
    [CLSCompliant(true)]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [Guid(RubberduckGuid.IRibbonViewModel)]
    public interface IRibbonViewModel {
        /// <summary>TODO</summary>
        [DispId(DispIds.RibbonFactory)]
        IRibbonFactory RibbonFactory { get; }

        /// <summary>TODO</summary>
        [DispId(DispIds.LoadImage)]
        object LoadImage(string imageId);

        /// <summary>Call back for GetDescription events from ribbon elements.</summary>
        [DispId(DispIds.Description)]
        string        GetDescription(IRibbonControl Control);
        /// <summary>Call back for GetEnabled events from ribbon elements.</summary>
        [DispId(DispIds.IsEnabled)]
        bool          GetEnabled    (IRibbonControl Control);
        /// <summary>Call back for GetKeyTip events from ribbon elements.</summary>
        [DispId(DispIds.KeyTip)]
        string        GetKeyTip     (IRibbonControl Control);
        /// <summary>Call back for GetLabel events from ribbon elements.</summary>
        [DispId(DispIds.Label)]
        string        GetLabel      (IRibbonControl Control);
        /// <summary>Call back for GetScreenTip events from ribbon elements.</summary>
        [DispId(DispIds.ScreenTip)]
        string        GetScreenTip  (IRibbonControl Control);
        /// <summary>Call back for GetSize events from ribbon elements.</summary>
        [DispId(DispIds.Size)]
        RdControlSize GetSize       (IRibbonControl Control);
        /// <summary>Call back for GetSuperTip events from ribbon elements.</summary>
        [DispId(DispIds.SuperTip)]
        string        GetSuperTip   (IRibbonControl Control);
        /// <summary>Call back for GetVisible events from ribbon elements.</summary>
        [DispId(DispIds.IsVisible)]
        bool          GetVisible    (IRibbonControl Control);

        /// <summary>Call back for GetImage events from ribbon elements.</summary>
        [DispId(DispIds.Image)]
        object        GetImage      (IRibbonControl Control);
        /// <summary>Call back for GetShowImage events from ribbon elements.</summary>
        [DispId(DispIds.ShowImage)]
        bool          GetShowImage  (IRibbonControl Control);
        /// <summary>Call back for GetShowLabe l events from ribbon elements.</summary>
        [DispId(DispIds.ShowLabel)]
        bool          GetShowLabel  (IRibbonControl Control);

        /// <summary>Call back for GetPressed events from the checkBox and toggleButton ribbon elements.</summary>
        [DispId(DispIds.IsPressed)]
        bool          GetPressed    (IRibbonControl Control);
        /// <summary>Call back for OnAction events from the checkBox and toggleButton ribbon elements.</summary>
        [DispId(DispIds.OnActionToggle)]
        void   OnActionToggle(IRibbonControl Control, bool Pressed);

        /// <summary>Call back for OnAction events from the button ribbon elements.</summary>
        [DispId(DispIds.OnAction)]
        void   OnAction(IRibbonControl Control);

        /// <summary>TODO</summary>
        [DispId(DispIds.SelectedItemId)]
        string GetSelectedItemId(IRibbonControl Control);
        /// <summary>TODO</summary>
        [DispId(DispIds.SelectedItemIndex)]
        int    GetSelectedItemIndex(IRibbonControl Control);
        /// <summary>Call back for OnAction events from the drop-down ribbon elements.</summary>
        [DispId(DispIds.OnActionDropDown)]
        void   OnActionDropDown(IRibbonControl Control, string SelectedId, int SelectedIndex);

        /// <summary>Call back for ItemCount events from the drop-down ribbon elements.</summary>
        [DispId(DispIds.ItemCount)]
        int    GetItemCount(IRibbonControl Control);
        /// <summary>Call back for GetItemID events from the drop-down ribbon elements.</summary>
        [DispId(DispIds.ItemId)]
        string GetItemId(IRibbonControl Control, int Index);
        /// <summary>Call back for GetItemLabel events from the drop-down ribbon elements.</summary>
        [DispId(DispIds.ItemLabel)]
        string GetItemLabel(IRibbonControl Control, int Index);
        /// <summary>Call back for GetItemScreenTip events from the drop-down ribbon elements.</summary>
        [DispId(DispIds.ItemScreenTip)]
        string GetItemScreenTip(IRibbonControl Control, int Index);
        /// <summary>Call back for GetItemSuperTip events from the drop-down ribbon elements.</summary>
        [DispId(DispIds.ItemSuperTip)]
        string GetItemSuperTip(IRibbonControl Control, int Index);

        /// <summary>Call back for GetItemImage events from the drop-down ribbon elements.</summary>
        [DispId(DispIds.ItemImage)]
        object GetItemImage(IRibbonControl Control, int Index);
        /// <summary>Call back for GetItemShowImage events from the drop-down ribbon elements.</summary>
        [DispId(DispIds.ItemShowImage)]
        bool   GetItemShowImage(IRibbonControl Control, int Index);
        /// <summary>Call back for GetItemShowLabel events from the drop-down ribbon elements.</summary>
        [DispId(DispIds.ItemShowLabel)]
        bool   GetItemShowLabel(IRibbonControl Control, int Index);
    }
}
