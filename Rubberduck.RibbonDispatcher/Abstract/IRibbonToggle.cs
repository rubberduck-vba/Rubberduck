using System;
using System.Runtime.InteropServices;

namespace Rubberduck.RibbonDispatcher.Abstract {

    /// <summary>TODO</summary>
    [ComVisible(true)]
    [Guid("D03E9DE1-F37D-40D6-89D6-A6B76A608D97")]
    [CLSCompliant(true)]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IRibbonToggle : IRibbonCheckBox {
    }

    /// <summary>TODO</summary>
    [ComVisible(true)]
    [Guid("42D56042-3FE9-4F1F-AD49-3ED0EE6CC987")]
    [CLSCompliant(true)]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IRibbonImageable {
        /// <summary>TODO</summary>
        bool ShowImage { get; }
        /// <summary>TODO</summary>
        bool ShowLabel { get; }
    }

    /// <summary>TODO</summary>
    [ComVisible(true)]
    [Guid("4BD0C027-BD10-4942-B9FA-96A29AB07FE8")]
    [CLSCompliant(true)]
    [InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    public interface IToggledEvents {
        /// <summary>TODO</summary>
        void Toggled(object sender, IToggledEventArgs e);
    }

    /// <summary>TODO</summary>
    [ComVisible(true)]
    [Guid("09B49B8B-145A-435D-BE62-17B605D3931A")]
    [CLSCompliant(true)]
    [InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    public interface IClickedEvents {
        /// <summary>TODO</summary>
        void Clicked(object sender, EventArgs e);
    }

    /// <summary>TODO</summary>
    [ComVisible(true)]
    [Guid("3AD5B841-BA7F-4CFA-9A60-8124B802BF46")]
    [CLSCompliant(true)]
    [InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    public interface ISelectionMadeEvents {
        /// <summary>TODO</summary>
        void SelectionMade(object sender, ISelectionMadeEventArgs e);
    }
}
