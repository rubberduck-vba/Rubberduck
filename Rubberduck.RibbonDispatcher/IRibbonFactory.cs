////////////////////////////////////////////////////////////////////////////////////////////////////
//                                Copyright (c) 2017 Pieter Geerkens                              //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.Diagnostics.CodeAnalysis;
using System.Runtime.InteropServices;
using stdole;

using Rubberduck.RibbonDispatcher.AbstractCOM;
using Rubberduck.RibbonDispatcher.Concrete;
using Rubberduck.RibbonDispatcher.EventHandlers;

namespace Rubberduck.RibbonDispatcher {
    using static RdControlSize;

    /// <summary>TODO</summary>
    [ComVisible(true)]
    [CLSCompliant(true)]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [Guid(RubberduckGuid.IRibbonFactory)]
    public interface IRibbonFactory {
        /// <summary>TODO</summary>
        [DispId(1)]
        void Invalidate();
        /// <summary>TODO</summary>
        [DispId(2)]
        void InvalidateControl(string controlId);
        /// <summary>TODO</summary>
        [DispId(3)]
        void InvalidateControlMso(string controlId);
        /// <summary>TODO</summary>
        [DispId(4)]
        void ActivateTab(string controlId);

        /// <summary>Returns a new Ribbon Group ViewModel instance.</summary>
        [DispId(11)]
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification="Matches COM usage.")]
        RibbonGroup NewRibbonGroup(
            string              id,
            bool                visible = true,
            bool                enabled = true
        );

        /// <summary>Returns a new Ribbon ActionButton ViewModel instance that uses a custom image (or none).</summary>
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification="Matches COM usage.")]
        [DispId(12)]
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
        /// <summary>Returns a new Ribbon ActionButton ViewModel instance that uses an Office built-in image.</summary>
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification="Matches COM usage.")]
        [DispId(13)]
        RibbonButton NewRibbonButtonMso(
            string              id,
            bool                visible         = true,
            bool                enabled         = true,
            RdControlSize       size            = rdLarge,
            string              imageMso        = "MacroSecurity",  // This one get's peope's attention ;-)
            bool                showImage       = false,
            bool                showLabel       = true,
            EventHandler        onClickedAction = null
        );

        /// <summary>Returns a new Ribbon ToggleButton ViewModel instance that uses a custom image (or none).</summary>
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification="Matches COM usage.")]
        [DispId(14)]
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
        /// <summary>Returns a new Ribbon ToggleButton ViewModel instance that uses an Office built-in image.</summary>
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification="Matches COM usage.")]
        [DispId(15)]
        RibbonToggleButton NewRibbonToggleMso(
            string              id,
            bool                visible         = true,
            bool                enabled         = true,
            RdControlSize       size            = rdLarge,
            string              imageMso        = "MacroSecurity",  // This one get's peope's attention ;-)
            bool                showImage       = false,
            bool                showLabel       = true,
            ToggledEventHandler onToggledAction = null
        );

        /// <summary>Returns a new Ribbon CheckBox ViewModel instance.</summary>
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification="Matches COM usage.")]
        [DispId(16)]
        RibbonCheckBox NewRibbonCheckBox(
            string              id,
            bool                visible         = true,
            bool                enabled         = true,
            ToggledEventHandler onToggledAction = null
        );

        /// <summary>Returns a new Ribbon DropDownViewModel instance.</summary>
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification="Matches COM usage.")]
        [DispId(17)]
        RibbonDropDown NewRibbonDropDown(
            string              id,
            bool                visible = true,
            bool                enabled = true,
            SelectionMadeEventHandler onSelectionMade = null,
            ISelectableItem[]   items   = null
        );

        /// <summary>Returns a new Ribbon DropDownViewModel instance.</summary>
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification = "Matches COM usage.")]
        [DispId(18)]
        SelectableItem NewSelectableItem(string id, IPictureDisp image = null);

        /// <summary>Returns a new Ribbon DropDownViewModel instance.</summary>
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification = "Matches COM usage.")]
        [DispId(19)]
        SelectableItem NewSelectableItemMso(string id, string imageMso = "MacroSecurity");
    }
}
