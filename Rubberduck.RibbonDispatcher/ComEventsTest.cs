using System;
using System.Diagnostics.CodeAnalysis;
using System.Runtime.InteropServices;
namespace EventSource {
    /// <summary>TODO</summary>
    public delegate void ClickEventHandler(int x, int y);
    /// <summary>TODO</summary>
    public delegate void ResizeEventHandler();
    /// <summary>TODO</summary>
    public delegate void PulseEventHandler();

    /// <summary>Step 1: Defines an event sink interface (ButtonEvents) to be     
    /// implemented by the COM sink.</summary>
    [CLSCompliant(true), ComVisible(true)]
    [InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    [Guid("12345679-0001-1010-1010-010101010101")]
    public interface IButtonEvents {
        /// <summary>TODO</summary>
        void Click(int x, int y);
        /// <summary>TODO</summary>
        void Resize();
        /// <summary>TODO</summary>
        void Pulse();
    }

    /// <summary>TODO</summary>
    [CLSCompliant(true), ComVisible(true)]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [Guid("12345679-0002-1010-1010-010101010101")]
    public interface IRibbonButton {
        /// <summary>TODO</summary>
        void CauseClickEvent(int x, int y);
        /// <summary>TODO</summary>
        void CauseResizeEvent();
        /// <summary>TODO</summary>
        void CausePulse();
    }
    /// <summary>Step 2: Connects the event sink interface to a class 
    /// by passing the namespace and event sink interface
    /// ("EventSource.ButtonEvents, EventSrc").</summary>
    [CLSCompliant(true), ComVisible(true), Serializable]
    [ClassInterface(ClassInterfaceType.None)]
    [ComSourceInterfaces(typeof(IButtonEvents))]
    [SuppressMessage("Microsoft.Interoperability", "CA1409:ComVisibleTypesShouldBeCreatable",
        Justification = "Publc, Non-Creatable class with exported Events.")]
    [ComDefaultInterface(typeof(IRibbonButton))]
    [Guid("12345679-0003-1010-1010-010101010101")]
    public class Button : IRibbonButton {
        /// <summary>TODO</summary>
        public event ClickEventHandler Click;
        /// <summary>TODO</summary>
        public event ResizeEventHandler Resize;
        /// <summary>TODO</summary>
        public event PulseEventHandler Pulse;

        internal Button() { }
        /// <summary>TODO</summary>
        public void CauseClickEvent(int x, int y) => Click(x, y);
        /// <summary>TODO</summary>
        public void CauseResizeEvent() => Resize();
        /// <summary>TODO</summary>
        public void CausePulse() => Pulse();
    }

    /// <summary>TODO</summary>
    [CLSCompliant(true), ComVisible(true)]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [Guid("12345679-0004-1010-1010-010101010101")]
    public interface IMaster {
        /// <summary>TODO</summary>
        IRibbonButton NewButton();
    }
    /// <summary>TODO</summary>
    [CLSCompliant(true), ComVisible(true), Serializable]
    [ClassInterface(ClassInterfaceType.None)]
    [ComSourceInterfaces(typeof(IButtonEvents))]
    [ComDefaultInterface(typeof(IMaster))]
    [Guid("12345679-0005-1010-1010-010101010101")]
    public class Master : IMaster {
        /// <summary>TODO</summary>
        public IRibbonButton NewButton() => new Button();
    }
}
