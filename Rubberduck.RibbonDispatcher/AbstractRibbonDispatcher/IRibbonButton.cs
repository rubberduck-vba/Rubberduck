﻿using System;
using System.Runtime.InteropServices;

namespace Rubberduck.RibbonDispatcher.Abstract {
    [ComVisible(true)][CLSCompliant(true)]
    public interface IRibbonButton {
        event EventHandler Clicked;

        bool ShowLabel { get; set; }
        bool ShowImage { get; set; }

        void OnAction();

        IRibbonCommon AsRibbonControl { get; }
    }
}
