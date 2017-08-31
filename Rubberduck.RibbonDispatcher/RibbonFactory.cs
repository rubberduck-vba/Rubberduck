using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics.CodeAnalysis;
using System.Resources;
using System.Runtime.InteropServices;

using Microsoft.Office.Core;

using Rubberduck.RibbonDispatcher.Abstract;
using Rubberduck.RibbonDispatcher.Concrete;
using Rubberduck.RibbonDispatcher.EventHandlers;

namespace Rubberduck.RibbonDispatcher {
    using stdole;
    using static MyRibbonControlSize;

    /// <summary>TODO</summary>
    /// <remarks>
    /// The {SuppressMessage} attributes are left in the source here, instead of being 'fired and
    /// forgotten' to the Global Suppresion file, as commentary on a practice often seen as a C#
    /// anti-pattern. Although non-standard C# practice, these "optional parameters with default 
    /// values" usages are (believed to be) the only means of implementing functionality equivalent
    /// to "overrides" in a COM-compatible way.
    /// </remarks>
    [ComVisible(false)]
    [CLSCompliant(true)]
    public class RibbonFactory {
        internal RibbonFactory(IRibbonUI ribbonUI, ResourceManager resourceManager) {
            _ribbonUI        = ribbonUI;
            _controls        = new Dictionary<string, IRibbonCommon>();
            _buttons         = new Dictionary<string, IActionItem>();
            _toggles         = new Dictionary<string, IToggleItem>();
            _dropDowns       = new Dictionary<string, IRibbonDropDown>();
            _imageables      = new Dictionary<string, IImageableItem>();
            _resourceManager = resourceManager;
        }

        private readonly IRibbonUI                              _ribbonUI;
        private readonly ResourceManager                        _resourceManager;
        private readonly IDictionary<string, IRibbonCommon>     _controls;
        private readonly IDictionary<string, IActionItem>       _buttons;
        private readonly IDictionary<string, IRibbonDropDown>   _dropDowns;
        private readonly IDictionary<string, IImageableItem>    _imageables;
        private readonly IDictionary<string, IToggleItem>       _toggles;

        /// <summary>Returns a readonly collection of all Ribbon Controls in this Ribbon ViewModel.</summary>
        public IReadOnlyDictionary<string, IRibbonCommon>    Controls   => new ReadOnlyDictionary<string, IRibbonCommon>(_controls);
        /// <summary>Returns a readonly collection of all Ribbon (Action) Buttons in this Ribbon ViewModel.</summary>
        public IReadOnlyDictionary<string, IActionItem>      Buttons    => new ReadOnlyDictionary<string, IActionItem>(_buttons);
        /// <summary>Returns a readonly collection of all Ribbon DropDowns in this Ribbon ViewModel.</summary>
        public IReadOnlyDictionary<string, IRibbonDropDown>  DropDowns  => new ReadOnlyDictionary<string, IRibbonDropDown>(_dropDowns);
        /// <summary>Returns a readonly collection of all Ribbon Imageable Controls in this Ribbon ViewModel.</summary>
        public IReadOnlyDictionary<string, IImageableItem>   Imageables => new ReadOnlyDictionary<string, IImageableItem>(_imageables);
        /// <summary>Returns a readonly collection of all Ribbon Toggle Buttons in this Ribbon ViewModel.</summary>
        public IReadOnlyDictionary<string, IToggleItem>      Toggles    => new ReadOnlyDictionary<string, IToggleItem>(_toggles);


        private void PropertyChanged(object sender, IControlChangedEventArgs e) => _ribbonUI.InvalidateControl(e.ControlId);

        private T Add<T>(T ctrl) where T:RibbonCommon {
            _controls.Add(ctrl.Id, ctrl);
            var button    = ctrl as IRibbonButton;    if (button    != null) _buttons   .Add(ctrl.Id, button);
            var dropDown  = ctrl as IRibbonDropDown;  if (dropDown  != null) _dropDowns .Add(ctrl.Id, dropDown);
            var imageable = ctrl as IImageableItem; if (imageable != null) _imageables.Add(ctrl.Id, imageable);
            var toggle    = ctrl as IToggleItem;      if (toggle    != null) _toggles   .Add(ctrl.Id, toggle);

            ctrl.Changed += PropertyChanged;
            return ctrl;
        }

        /// <summary>Returns a new Ribbon Group ViewModel instance.</summary>
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification="Matches COM usage.")]
        public IRibbonCommon NewRibbonGroup(
            string              id,
            bool                visible = true,
            bool                enabled = true,
            MyRibbonControlSize size    = Large
        ) => Add(new RibbonGroup(id, _resourceManager, visible, enabled, size));

        /// <summary>Returns a new Ribbon ActionButton ViewModel instance.</summary>
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification="Matches COM usage.")]
        public IRibbonButton NewRibbonButton(
            string              id,
            bool                visible         = true,
            bool                enabled         = true,
            MyRibbonControlSize size            = Large,
            string              imageMso        = "Unknown",
            bool                showImage       = false,
            bool                showLabel       = true,
            EventHandler        onClickedAction = null
        ) => Add(new RibbonButton(id, _resourceManager, visible, enabled, size, imageMso, showImage, showLabel, onClickedAction));
        /// <summary>Returns a new Ribbon ActionButton ViewModel instance.</summary>
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification="Matches COM usage.")]
        public IRibbonButton NewRibbonButton(
            string              id,
            bool                visible         = true,
            bool                enabled         = true,
            MyRibbonControlSize size            = Large,
            IPictureDisp        image           = null,
            bool                showImage       = false,
            bool                showLabel       = true,
            EventHandler        onClickedAction = null
        ) => Add(new RibbonButton(id, _resourceManager, visible, enabled, size, image, showImage, showLabel, onClickedAction));

        /// <summary>Returns a new Ribbon ToggleButton ViewModel instance.</summary>
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification="Matches COM usage.")]
        public IRibbonToggle NewRibbonToggle(
            string              id,
            bool                visible         = true,
            bool                enabled         = true,
            MyRibbonControlSize size            = Large,
            string              imageMso        = "Unknown",
            bool                showImage       = false,
            bool                showLabel       = true,
            ToggledEventHandler onToggledAction = null
        ) => Add(new RibbonToggle(id, _resourceManager, visible, enabled, size, imageMso, showImage, showLabel, onToggledAction));
        /// <summary>Returns a new Ribbon ToggleButton ViewModel instance.</summary>
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification="Matches COM usage.")]
        public IRibbonToggle NewRibbonToggle(
            string              id,
            bool                visible         = true,
            bool                enabled         = true,
            MyRibbonControlSize size            = Large,
            IPictureDisp        image           = null,
            bool                showImage       = false,
            bool                showLabel       = true,
            ToggledEventHandler onToggledAction = null
        ) => Add(new RibbonToggle(id, _resourceManager, visible, enabled, size, image, showImage, showLabel, onToggledAction));

        /// <summary>Returns a new Ribbon CheckBox ViewModel instance.</summary>
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification="Matches COM usage.")]
        public IRibbonCheckBox NewRibbonCheckBox(
            string              id,
            bool                visible         = true,
            bool                enabled         = true,
            MyRibbonControlSize size            = Large,
            ToggledEventHandler onToggledAction = null
        ) => Add(new RibbonCheckBox(id, _resourceManager, visible, enabled, size, onToggledAction));

        /// <summary>Returns a new Ribbon DropDownViewModel instance.</summary>
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification="Matches COM usage.")]
        public IRibbonDropDown NewRibbonDropDown(
            string              id,
            bool                visible = true,
            bool                enabled = true,
            MyRibbonControlSize size    = Large
        ) => Add(new RibbonDropDown(id, _resourceManager, visible, enabled, size));

        /// <summary>TODO</summary>
        public void Invalidate()                            => _ribbonUI.Invalidate();
        /// <summary>TODO</summary>
        public void InvalidateControl(string controlId)     => _ribbonUI.InvalidateControl(controlId);
        /// <summary>TODO</summary>
        public void InvalidateControlMso(string controlId)  => _ribbonUI.InvalidateControlMso(controlId);
        /// <summary>TODO</summary>
        public void ActivateTab(string controlId)           => _ribbonUI.ActivateTab(controlId);
   }

    #region WorkInProgress for VBA Excel Add-Ins
    /// <summary>TODO</summary>
    public interface IMain {
        /// <summary>TODO</summary>
        IRibbonViewModel NewRibbonViewModel(IRibbonUI ribbonUI, ResourceManager resourceManager);
    }

    /// <summary>TODO</summary>
    public class Main : IMain {
        /// <summary>TODO</summary>
        public Main() { }
        /// <inheritdoc/>
        public IRibbonViewModel NewRibbonViewModel(IRibbonUI ribbonUI, ResourceManager resourceManager)
            => new RibbonViewModel(ribbonUI, resourceManager);
    }

    /// <summary>TODO</summary>
    [SuppressMessage("Microsoft.Design", "CA1040:AvoidEmptyInterfaces")]
    public interface IRibbonViewModel {

    }

    /// <summary>TODO</summary>
    public class RibbonViewModel : AbstractRibbonDispatcher, IRibbonViewModel {
        /// <summary>TODO</summary>
        public RibbonViewModel(IRibbonUI ribbonUI, ResourceManager resourceManager) : base() {
            InitializeRibbonFactory(ribbonUI, resourceManager);
        }
    }
    #endregion
}
