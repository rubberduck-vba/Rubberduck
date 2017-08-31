////////////////////////////////////////////////////////////////////////////////////////////////////
//                                Copyright (c) 2017 Pieter Geerkens                              //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.Diagnostics.CodeAnalysis;
using System.Runtime.InteropServices;

using Rubberduck.RibbonDispatcher.AbstractCOM;
using Rubberduck.RibbonDispatcher.Concrete;
using Rubberduck.RibbonDispatcher.EventHandlers;

namespace Rubberduck.RibbonDispatcher {
    using stdole;
    using static RdControlSize;
    /// <summary>TODO</summary>
    [ComVisible(true)]
    [Guid("ACF7C6E9-8314-484F-B81C-9B926E0731AC")]
    [CLSCompliant(true)]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IRibbonFactory {
        /// <summary>Returns a new Ribbon Group ViewModel instance.</summary>
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification="Matches COM usage.")]
        RibbonGroup NewRibbonGroup(
            string              id,
            bool                visible = true,
            bool                enabled = true,
            RdControlSize       size    = rdLarge
        );

        /// <summary>Returns a new Ribbon ActionButton ViewModel instance.</summary>
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification="Matches COM usage.")]
        RibbonButton NewRibbonButton(
            string              id,
            bool                visible         = true,
            bool                enabled         = true,
            RdControlSize       size            = rdLarge,
            IPictureDisp        image           = null,
            bool                showImage       = false,
            bool                showLabel       = true,
            EventHandler        onClickedAction = null
        );
        /// <summary>Returns a new Ribbon ActionButton ViewModel instance.</summary>
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification="Matches COM usage.")]
        RibbonButton NewRibbonButtonMso(
            string              id,
            bool                visible         = true,
            bool                enabled         = true,
            RdControlSize       size            = rdLarge,
            string              imageMso        = "Unknown",
            bool                showImage       = false,
            bool                showLabel       = true,
            EventHandler        onClickedAction = null
        );

        /// <summary>Returns a new Ribbon ToggleButton ViewModel instance.</summary>
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification="Matches COM usage.")]
        RibbonToggleButton NewRibbonToggle(
            string              id,
            bool                visible         = true,
            bool                enabled         = true,
            RdControlSize       size            = rdLarge,
            IPictureDisp        image           = null,
            bool                showImage       = false,
            bool                showLabel       = true,
            ToggledEventHandler onToggledAction = null
        );
        /// <summary>Returns a new Ribbon ToggleButton ViewModel instance.</summary>
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification="Matches COM usage.")]
        RibbonToggleButton NewRibbonToggleMso(
            string              id,
            bool                visible         = true,
            bool                enabled         = true,
            RdControlSize       size            = rdLarge,
            string              imageMso        = "Unknown",
            bool                showImage       = false,
            bool                showLabel       = true,
            ToggledEventHandler onToggledAction = null
        );

        /// <summary>Returns a new Ribbon CheckBox ViewModel instance.</summary>
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification="Matches COM usage.")]
        RibbonCheckBox NewRibbonCheckBox(
            string              id,
            bool                visible         = true,
            bool                enabled         = true,
            RdControlSize       size            = rdLarge,
            ToggledEventHandler onToggledAction = null
        );

        /// <summary>Returns a new Ribbon DropDownViewModel instance.</summary>
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification="Matches COM usage.")]
        RibbonDropDown NewRibbonDropDown(
            string              id,
            bool                visible = true,
            bool                enabled = true,
            RdControlSize       size    = rdLarge
        );

        /// <summary>TODO</summary>
        void Invalidate();
        /// <summary>TODO</summary>
        void InvalidateControl(string controlId);
        /// <summary>TODO</summary>
        void InvalidateControlMso(string controlId);
        /// <summary>TODO</summary>
        void ActivateTab(string controlId);
    }
}
