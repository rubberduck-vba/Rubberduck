////////////////////////////////////////////////////////////////////////////////////////////////////
//                                Copyright (c) 2017 Pieter Geerkens                              //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.Runtime.InteropServices;
using Microsoft.Office.Core;

namespace Rubberduck.RibbonDispatcher.AbstractCOM {
    /// <summary>TODO</summary>
    [ComVisible(true)]
    [Guid("5AC9C64D-D476-42E3-8700-EB19C1863519")]
    [CLSCompliant(true)]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IAbstractDispatcher {
        /// <summary>TODO</summary>
        IRibbonFactory RibbonFactory { get; }

        /// <summary>Call back for GetDescription events from ribbon elements.</summary>
        string        GetDescription(IRibbonControl control);
        /// <summary>Call back for GetEnabled events from ribbon elements.</summary>
        bool          GetEnabled    (IRibbonControl control);
        /// <summary>Call back for GetKeyTip events from ribbon elements.</summary>
        string        GetKeyTip     (IRibbonControl control);
        /// <summary>Call back for GetLabel events from ribbon elements.</summary>
        string        GetLabel      (IRibbonControl control);
        /// <summary>Call back for GetScreenTip events from ribbon elements.</summary>
        string        GetScreenTip  (IRibbonControl control);
        /// <summary>Call back for GetSize events from ribbon elements.</summary>
        RdControlSize GetSize       (IRibbonControl control);
        /// <summary>Call back for GetSuperTip events from ribbon elements.</summary>
        string        GetSuperTip   (IRibbonControl control);
        /// <summary>Call back for GetVisible events from ribbon elements.</summary>
        bool          GetVisible    (IRibbonControl control);

        /// <summary>Call back for GetImage events from ribbon elements.</summary>
        object        GetImage      (IRibbonControl control);
        /// <summary>Call back for GetShowImage events from ribbon elements.</summary>
        bool          GetShowImage  (IRibbonControl control);
        /// <summary>Call back for GetShowLabe l events from ribbon elements.</summary>
        bool          GetShowLabel  (IRibbonControl control);

        /// <summary>Call back for GetPressed events from the checkBox and toggleButton ribbon elements.</summary>
        bool          GetPressed    (IRibbonControl control);
        /// <summary>Call back for OnAction events from the checkBox and toggleButton ribbon elements.</summary>
        void OnActionToggle(IRibbonControl control, bool pressed);

        /// <summary>Call back for OnAction events from the button ribbon elements.</summary>
        void OnAction(IRibbonControl control);
    }
}
