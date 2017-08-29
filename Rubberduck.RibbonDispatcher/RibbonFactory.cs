using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics.CodeAnalysis;
using System.Runtime.InteropServices;

using Microsoft.Office.Core;

using Rubberduck.RibbonDispatcher.Abstract;
using Rubberduck.RibbonDispatcher.Concrete;

namespace Rubberduck.RibbonDispatcher
{
    using static RibbonControlSize;

    using ControlSize         = RibbonControlSize;
    using LanguageStrings     = IRibbonTextLanguageControl;
    using ClickedEventHandler = EventHandler<ClickedEventArgs>;

    /// <summary>TODO</summary>
    /// <remarks>
    /// The {SuppressMessage} attributes are left in the source here, instead of being 'fired and
    /// forgotten' to the Global Suppresion file. This explicitly points out that although "optional
    /// parameters with default values" is often regarded ad an anti-pattern for DOT NET development,
    /// COM does not support overrides. This is (believed) to be the only means of implementing
    /// the equivalent functionality in a COM-compatible way.
    /// </remarks>
    [ComVisible(false)]
    [CLSCompliant(true)]
    public class RibbonFactory {
        internal RibbonFactory(IRibbonUI ribbonUI) {
            _ribbonUI  = ribbonUI;
            _controls  = new Dictionary<string,IRibbonCommon>();
            _buttons   = new Dictionary<string,IRibbonButton>();
            _toggles   = new Dictionary<string,IRibbonToggle>();
            _dropDowns = new Dictionary<string,IRibbonDropDown>();
        }

        private readonly IRibbonUI _ribbonUI;

        private void PropertyChanged(object sender, ChangedControlEventArgs e) => _ribbonUI.InvalidateControl(e.ControlId);

        private T Add<T>(T ctrl) where T:IRibbonCommon {
            _controls.Add(ctrl.Id, ctrl);
            var button   = ctrl as IRibbonButton;   if (button   != null) _buttons  .Add(ctrl.Id, button);
            var toggle   = ctrl as IRibbonToggle;   if (toggle   != null) _toggles  .Add(ctrl.Id, toggle);
            var dropDown = ctrl as IRibbonDropDown; if (dropDown != null) _dropDowns.Add(ctrl.Id, dropDown);

            ctrl.Changed += PropertyChanged;
            return ctrl;
        }

        public IReadOnlyDictionary<string,IRibbonCommon>     Controls  => new ReadOnlyDictionary<string,IRibbonCommon>(_controls);
        private IDictionary<string,IRibbonCommon>           _controls  { get; }
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification="Matches COM usage.")]
        public IRibbonCommon NewRibbonGroup(
            string          id,
            LanguageStrings strings = null, 
            bool            visible = true,
            bool            enabled = true,
            ControlSize     size    = RibbonControlSizeLarge
        ) => Add(new RibbonGroup(id, strings, visible, enabled, size));

        public IReadOnlyDictionary<string,IRibbonButton>     Buttons   => new ReadOnlyDictionary<string,IRibbonButton>(_buttons);
        private IDictionary<string,IRibbonButton>           _buttons   { get; }
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification="Matches COM usage.")]
        public IRibbonButton NewRibbonButton(
            string          id,
            LanguageStrings strings   = null, 
            bool            visible   = true,
            bool            enabled   = true,
            ControlSize     size      = RibbonControlSizeLarge,
            bool            showImage = false,
            bool            showLabel = true,
            EventHandler    onClickedAction = null
        ) => Add(new RibbonButton(id, strings, visible, enabled, size, showImage, showLabel, onClickedAction));

        public IReadOnlyDictionary<string,IRibbonToggle>     Toggles   => new ReadOnlyDictionary<string,IRibbonToggle>(_toggles);
        private IDictionary<string,IRibbonToggle>           _toggles   { get; }
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification="Matches COM usage.")]
        public IRibbonToggle NewRibbonToggle(
            string          id,
            LanguageStrings strings   = null, 
            bool            visible   = true,
            bool            enabled   = true,
            ControlSize     size      = RibbonControlSizeLarge,
            bool            showImage = false,
            bool            showLabel = true,
            ClickedEventHandler onClickedAction = null
        ) => Add(new RibbonToggle(id, strings, visible, enabled, size, showImage, showLabel, onClickedAction));

        public IReadOnlyDictionary<string,IRibbonDropDown>   DropDowns => new ReadOnlyDictionary<string,IRibbonDropDown>(_dropDowns);
        private IDictionary<string,IRibbonDropDown>         _dropDowns { get; }
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification="Matches COM usage.")]
        public IRibbonDropDown NewRibbonDropDown(
            string          id,
            LanguageStrings strings = null, 
            bool            visible = true,
            bool            enabled = true,
            ControlSize     size    = RibbonControlSizeLarge
        ) => Add(new RibbonDropDown(id, strings, visible, enabled, size));

        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification="Matches COM usage.")]
        public static LanguageStrings NewLanguageControlRibbonText(
            string label,
            string screenTip      = null,
            string superTip       = null,
            string keyTip         = null,
            string alternateLabel = null,
            string description    = null
        ) => new RibbonTextLanguageControl(label, screenTip, superTip, keyTip, alternateLabel, description);

    }
}
