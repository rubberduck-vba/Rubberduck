////////////////////////////////////////////////////////////////////////////////////////////////////
//                                Copyright (c) 2017 Pieter Geerkens                              //
////////////////////////////////////////////////////////////////////////////////////////////////////
namespace Rubberduck.RibbonDispatcher.Abstract {
    /// <summary>The total interface (required to be) exposed externally by RibbonButton objects.</summary>
    public interface IToggleItem {
        /// <summary>TODO</summary>
        bool IsPressed { get; }

        /// <summary>TODO</summary>
        void OnAction(bool isPressed);
    }
}
