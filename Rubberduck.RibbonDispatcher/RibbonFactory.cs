using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics.CodeAnalysis;
using System.Resources;
using stdole;

using Microsoft.Office.Core;

using Rubberduck.RibbonDispatcher.ControlDecorators;
using Rubberduck.RibbonDispatcher.AbstractCOM;
using Rubberduck.RibbonDispatcher.Concrete;
using Rubberduck.RibbonDispatcher.EventHandlers;

namespace Rubberduck.RibbonDispatcher {
    using System.Runtime.InteropServices;
    using static RdControlSize;

    /// <summary>TODO</summary>
    /// <remarks>
    /// The {SuppressMessage} attributes are left in the source here, instead of being 'fired and
    /// forgotten' to the Global Suppresion file, as commentary on a practice often seen as a C#
    /// anti-pattern. Although non-standard C# practice, these "optional parameters with default 
    /// values" usages are (believed to be) the only means of implementing functionality equivalent
    /// to "overrides" in a COM-compatible way.
    /// </remarks>
    [Serializable]
    [CLSCompliant(true)]
    [ComDefaultInterface(typeof(IRibbonFactory))]
    public class RibbonFactory : IRibbonFactory {
        internal RibbonFactory(IRibbonUI ribbonUI, ResourceManager resourceManager) {
            _ribbonUI        = ribbonUI;
            _resourceManager = resourceManager;
            _controls        = new Dictionary<string, IRibbonCommon>();
            _sizeables       = new Dictionary<string, ISizeableDecorator>();
            _actionables     = new Dictionary<string, IActionableDecorator>();
            _toggleables     = new Dictionary<string, IToggleableDecorator>();
            _selectables     = new Dictionary<string, ISelectableDecorator>();
            _imageables      = new Dictionary<string, IImageableDecorator>();
        }

        private readonly IRibbonUI                                  _ribbonUI;
        private readonly ResourceManager                            _resourceManager;
        private readonly IDictionary<string, IRibbonCommon>         _controls;
        private readonly IDictionary<string, ISizeableDecorator>    _sizeables;
        private readonly IDictionary<string, IActionableDecorator>  _actionables;
        private readonly IDictionary<string, ISelectableDecorator>  _selectables;
        private readonly IDictionary<string, IImageableDecorator>   _imageables;
        private readonly IDictionary<string, IToggleableDecorator>  _toggleables;

        /// <summary>Returns a readonly collection of all Ribbon Controls in this Ribbon ViewModel.</summary>
        public IReadOnlyDictionary<string, IRibbonCommon>        Controls    => new ReadOnlyDictionary<string, IRibbonCommon>(_controls);
        /// <summary>Returns a readonly collection of all Ribbon (Action) Buttons in this Ribbon ViewModel.</summary>
        public IReadOnlyDictionary<string, ISizeableDecorator>   Sizeables   => new ReadOnlyDictionary<string, ISizeableDecorator>(_sizeables);
        /// <summary>Returns a readonly collection of all Ribbon (Action) Buttons in this Ribbon ViewModel.</summary>
        public IReadOnlyDictionary<string, IActionableDecorator> Actionables => new ReadOnlyDictionary<string, IActionableDecorator>(_actionables);
        /// <summary>Returns a readonly collection of all Ribbon DropDowns in this Ribbon ViewModel.</summary>
        public IReadOnlyDictionary<string, ISelectableDecorator> Selectables => new ReadOnlyDictionary<string, ISelectableDecorator>(_selectables);
        /// <summary>Returns a readonly collection of all Ribbon Imageable Controls in this Ribbon ViewModel.</summary>
        public IReadOnlyDictionary<string, IImageableDecorator>  Imageables  => new ReadOnlyDictionary<string, IImageableDecorator>(_imageables);
        /// <summary>Returns a readonly collection of all Ribbon Toggle Buttons in this Ribbon ViewModel.</summary>
        public IReadOnlyDictionary<string, IToggleableDecorator> Toggleables => new ReadOnlyDictionary<string, IToggleableDecorator>(_toggleables);

        private void PropertyChanged(object sender, IControlChangedEventArgs e) => _ribbonUI.InvalidateControl(e.ControlId);

        private T Add<T>(T ctrl) where T:RibbonCommon {
            _controls.Add(ctrl.Id, ctrl);

            _actionables.AddNotNull(ctrl.Id, ctrl as IActionableDecorator);
            _sizeables.AddNotNull(ctrl.Id, ctrl as ISizeableDecorator);
            _selectables.AddNotNull(ctrl.Id, ctrl as ISelectableDecorator);
            _imageables.AddNotNull(ctrl.Id, ctrl as IImageableDecorator);
            _toggleables.AddNotNull(ctrl.Id, ctrl as IToggleableDecorator);

            ctrl.Changed += PropertyChanged;
            return ctrl;
        }

        /// <summary>TODO</summary>
        public void Invalidate()                            => _ribbonUI.Invalidate();
        /// <summary>TODO</summary>
        public void InvalidateControl(string controlId)     => _ribbonUI.InvalidateControl(controlId);
        /// <summary>TODO</summary>
        public void InvalidateControlMso(string controlId)  => _ribbonUI.InvalidateControlMso(controlId);
        /// <summary>TODO</summary>
        public void ActivateTab(string controlId)           => _ribbonUI.ActivateTab(controlId);

        /// <summary>Returns a new Ribbon Group ViewModel instance.</summary>
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification="Matches COM usage.")]
        public RibbonGroup NewRibbonGroup(
            string              itemId,
            bool                visible         = true,
            bool                enabled         = true
        ) => Add(new RibbonGroup(itemId, _resourceManager, visible, enabled));

        /// <summary>Returns a new Ribbon ActionButton ViewModel instance.</summary>
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification="Matches COM usage.")]
        public RibbonButton NewRibbonButton(
            string              itemId,
            bool                visible         = true,
            bool                enabled         = true,
            RdControlSize       size            = rdLarge,
            IPictureDisp        image           = null,
            bool                showImage       = true,
            bool                showLabel       = true,
            EventHandler        onClickedAction = null
        ) => Add(new RibbonButton(itemId, _resourceManager, visible, enabled, size, image, showImage, showLabel, onClickedAction));
        /// <summary>Returns a new Ribbon ActionButton ViewModel instance.</summary>
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification="Matches COM usage.")]
        public RibbonButton NewRibbonButtonMso(
            string              itemId,
            bool                visible         = true,
            bool                enabled         = true,
            RdControlSize       size            = rdLarge,
            string              imageMso        = "Unknown",
            bool                showImage       = true,
            bool                showLabel       = true,
            EventHandler        onClickedAction = null
        ) => Add(new RibbonButton(itemId, _resourceManager, visible, enabled, size, imageMso, showImage, showLabel, onClickedAction));

        /// <summary>Returns a new Ribbon ToggleButton ViewModel instance.</summary>
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification="Matches COM usage.")]
        public RibbonToggleButton NewRibbonToggle(
            string              itemId,
            bool                visible         = true,
            bool                enabled         = true,
            RdControlSize       size            = rdLarge,
            IPictureDisp        image           = null,
            bool                showImage       = true,
            bool                showLabel       = true,
            ToggledEventHandler onToggledAction = null
        ) => Add(new RibbonToggleButton(itemId, _resourceManager, visible, enabled, size, image, showImage, showLabel, onToggledAction));
        /// <summary>Returns a new Ribbon ToggleButton ViewModel instance.</summary>
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification="Matches COM usage.")]
        public RibbonToggleButton NewRibbonToggleMso(
            string              itemId,
            bool                visible         = true,
            bool                enabled         = true,
            RdControlSize       size            = rdLarge,
            string              imageMso        = "Unknown",
            bool                showImage       = true,
            bool                showLabel       = true,
            ToggledEventHandler onToggledAction = null
        ) => Add(new RibbonToggleButton(itemId, _resourceManager, visible, enabled, size, imageMso, showImage, showLabel, onToggledAction));

        /// <summary>Returns a new Ribbon CheckBox ViewModel instance.</summary>
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification="Matches COM usage.")]
        public RibbonCheckBox NewRibbonCheckBox(
            string              itemId,
            bool                visible         = true,
            bool                enabled         = true,
            ToggledEventHandler onToggledAction = null
        ) => Add(new RibbonCheckBox(itemId, _resourceManager, visible, enabled, onToggledAction));

        /// <summary>Returns a new Ribbon DropDownViewModel instance.</summary>
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification="Matches COM usage.")]
        public RibbonDropDown NewRibbonDropDown(
            string               itemId,
            bool                 visible         = true,
            bool                 enabled         = true,
            SelectedEventHandler onSelectionMade = null,
            ISelectableItem[]    items           = null
        ) => Add(new RibbonDropDown(itemId, _resourceManager, visible, enabled, onSelectionMade, items));

        /// <inheritdoc/>
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification = "Matches COM usage.")]
        public SelectableItem NewSelectableItem(string itemId, IPictureDisp image = null)
            => new SelectableItem(itemId, _resourceManager, image);

        /// <inheritdoc/>
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification = "Matches COM usage.")]
        public SelectableItem NewSelectableItemMso(string itemId, string imageMso = "MacroSecurity")
            => new SelectableItem(itemId, _resourceManager, imageMso);
    }
}
