////////////////////////////////////////////////////////////////////////////////////////////////////
//                                Copyright (c) 2017 Pieter Geerkens                              //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.Runtime.InteropServices;
using Microsoft.Office.Core;

using Dispids = Rubberduck.RibbonDispatcher.DispIds;

namespace Rubberduck.RibbonDispatcher.AbstractCOM {
    /// <summary>TODO</summary>
    [ComVisible(true)]
    [CLSCompliant(true)]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [Guid(RubberduckGuid.IAbstractDispatcher)]
    public interface IAbstractDispatcher {
        /// <summary>TODO</summary>
        [DispId(Dispids.RibbonFactory)]
        IRibbonFactory RibbonFactory { get; }

        /// <summary>Call back for GetDescription events from ribbon elements.</summary>
        [DispId(Dispids.Description)]
        string        GetDescription(IRibbonControl control);
        /// <summary>Call back for GetEnabled events from ribbon elements.</summary>
        [DispId(Dispids.IsEnabled)]
        bool          GetEnabled    (IRibbonControl control);
        /// <summary>Call back for GetKeyTip events from ribbon elements.</summary>
        [DispId(Dispids.KeyTip)]
        string        GetKeyTip     (IRibbonControl control);
        /// <summary>Call back for GetLabel events from ribbon elements.</summary>
        [DispId(Dispids.Label)]
        string        GetLabel      (IRibbonControl control);
        /// <summary>Call back for GetScreenTip events from ribbon elements.</summary>
        [DispId(Dispids.ScreenTip)]
        string        GetScreenTip  (IRibbonControl control);
        /// <summary>Call back for GetSize events from ribbon elements.</summary>
        [DispId(Dispids.Size)]
        RdControlSize GetSize       (IRibbonControl control);
        /// <summary>Call back for GetSuperTip events from ribbon elements.</summary>
        [DispId(Dispids.SuperTip)]
        string        GetSuperTip   (IRibbonControl control);
        /// <summary>Call back for GetVisible events from ribbon elements.</summary>
        [DispId(Dispids.IsVisible)]
        bool          GetVisible    (IRibbonControl control);

        /// <summary>Call back for GetImage events from ribbon elements.</summary>
        [DispId(Dispids.Image)]
        object        GetImage      (IRibbonControl control);
        /// <summary>Call back for GetShowImage events from ribbon elements.</summary>
        [DispId(Dispids.ShowImage)]
        bool          GetShowImage  (IRibbonControl control);
        /// <summary>Call back for GetShowLabe l events from ribbon elements.</summary>
        [DispId(Dispids.ShowLabel)]
        bool          GetShowLabel  (IRibbonControl control);

        /// <summary>Call back for GetPressed events from the checkBox and toggleButton ribbon elements.</summary>
        [DispId(Dispids.IsPressed)]
        bool          GetPressed    (IRibbonControl control);
        /// <summary>Call back for OnAction events from the checkBox and toggleButton ribbon elements.</summary>
        [DispId(Dispids.OnActionToggle)]
        void OnActionToggle(IRibbonControl control, bool pressed);

        /// <summary>Call back for OnAction events from the button ribbon elements.</summary>
        [DispId(Dispids.OnAction)]
        void OnAction(IRibbonControl control);
    }
}
