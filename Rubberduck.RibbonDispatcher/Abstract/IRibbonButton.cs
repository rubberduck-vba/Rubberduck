using System;
using System.Runtime.InteropServices;

namespace Rubberduck.RibbonDispatcher.Abstract
{
    [ComVisible(true)]
    [CLSCompliant(true)]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IRibbonButton : IRibbonCommon {
        event EventHandler Clicked;

        void OnAction();
    }
}
