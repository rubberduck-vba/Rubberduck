using System.Collections.Generic;
using System.Collections.Immutable;
using Microsoft.Office.Core;

namespace RubberDuck.RibbonDispatcher {
    using System;
    using static RibbonControlSize;
    using System.Runtime.InteropServices;

    [ComVisible(true)][CLSCompliant(true)]
    public class RibbonFactory {
        internal RibbonFactory(IRibbonUI ribbonUI) {
            RibbonUI = ribbonUI;

            _Controls     = new Dictionary<string,IRibbonCommon>();
            _Buttons      = new Dictionary<string,IRibbonButton>();
            _Toggles      = new Dictionary<string,IRibbonToggle>();
            _DropDowns    = new Dictionary<string,IRibbonDropDown>();
        }

        private IRibbonUI    RibbonUI    { get; }

        private void PropertyChanged(object sender, ChangedControlEventArgs e) => RibbonUI.InvalidateControl(e.ControlId);

        private IRibbonCommon Add(IRibbonCommon ctrl) {
            _Controls.Add(ctrl.Id, ctrl);
            if (ctrl is IRibbonButton     button) _Buttons.Add(ctrl.Id, button);
            if (ctrl is IRibbonToggle     toggle) _Toggles.Add(ctrl.Id, toggle);
            if (ctrl is IRibbonDropDown dropDown) _DropDowns.Add(ctrl.Id, dropDown);

            ctrl.Changed += PropertyChanged;
            return ctrl;
        }

        #region IRibbonCommon - all controls including Groups & Tabs
        public IReadOnlyDictionary<string,IRibbonCommon>     Controls  => _Controls.ToImmutableDictionary();       
        private IDictionary<string,IRibbonCommon>           _Controls  { get; }
        public IRibbonCommon NewRibbonCommon(string id)
                => NewRibbonCommon(id, true, true, RibbonControlSizeLarge);
        public IRibbonCommon NewRibbonCommon(string id, bool visible)
                => NewRibbonCommon(id, visible, true, RibbonControlSizeLarge);
        public IRibbonCommon NewRibbonCommon(string id, bool visible, bool enabled)
                => NewRibbonCommon(id, visible, enabled,  RibbonControlSizeLarge);
        public IRibbonCommon NewRibbonCommon(string id, bool visible, bool enabled, RibbonControlSize size)
            => Add(new RibbonCommon(id, visible, enabled, size));
        #endregion

        #region IRibbonButton - standard (action) buttons
        public IReadOnlyDictionary<string,IRibbonButton>     Buttons   => _Buttons.ToImmutableDictionary();
        private IDictionary<string,IRibbonButton>           _Buttons   { get; }
        public RibbonButton NewRibbonButton(string id)
                => NewRibbonButton(id, true, true, RibbonControlSizeLarge);
        public RibbonButton NewRibbonButton(string id, bool visible)
                => NewRibbonButton(id, visible, true, RibbonControlSizeLarge);
        public RibbonButton NewRibbonButton(string id, bool visible, bool enabled)
                => NewRibbonButton(id, visible, enabled,  RibbonControlSizeLarge);
        public RibbonButton NewRibbonButton(string id, bool visible, bool enabled, RibbonControlSize size) {
            var ctrl = new RibbonButton(id, visible, enabled, size);
            Add(ctrl);
            return ctrl;
        }
        #endregion

        #region IRibbonToggle - checkBoxes & toggleButtons
        public IReadOnlyDictionary<string,IRibbonToggle>     Toggles   => _Toggles.ToImmutableDictionary();
        private IDictionary<string,IRibbonToggle>           _Toggles   { get; }
        public RibbonToggle NewRibbonToggle(string id)
                => NewRibbonToggle(id, true, true, RibbonControlSizeLarge);
        public RibbonToggle NewRibbonToggle(string id, bool visible)
                => NewRibbonToggle(id, visible, true, RibbonControlSizeLarge);
        public RibbonToggle NewRibbonToggle(string id, bool visible, bool enabled)
                => NewRibbonToggle(id, visible, enabled,  RibbonControlSizeLarge);
        public RibbonToggle NewRibbonToggle(string id, bool visible, bool enabled, RibbonControlSize size) {
            var ctrl = new RibbonToggle(id, visible, enabled, size);
            Add(ctrl);
            return ctrl;
        }
        #endregion

        #region IRibbonDropdown - comboBoxes & dropDowns
        public IReadOnlyDictionary<string,IRibbonDropDown>   DropDowns => _DropDowns.ToImmutableDictionary();
        private IDictionary<string,IRibbonDropDown>         _DropDowns { get; }
        public RibbonDropDown NewRibbonDropDown(string id)
                => NewRibbonDropDown(id, true, true, RibbonControlSizeLarge);
        public RibbonDropDown NewRibbonDropDown(string id, bool visible)
                => NewRibbonDropDown(id, visible, true, RibbonControlSizeLarge);
        public RibbonDropDown NewRibbonDropDown(string id, bool visible, bool enabled)
                => NewRibbonDropDown(id, visible, enabled,  RibbonControlSizeLarge);
        public RibbonDropDown NewRibbonDropDown(string id, bool visible, bool enabled, RibbonControlSize size) {
            var ctrl = new RibbonDropDown(id, visible, enabled, size);
            Add(ctrl);
            return ctrl;
        }
        #endregion

        #region Control Strings by Language
        public IRibbonTextLanguageControl NewLanguageControlRibbonText(
            string label,
            string screenTip,
            string superTip,
            string keyTip
        ) => new RibbonTextLanguageControl(label, screenTip, superTip, keyTip, null, null);

        public IRibbonTextLanguageControl NewLanguageControlRibbonText(
            string label,
            string screenTip,
            string superTip,
            string keyTip,
            string alternateLabel
        ) => new RibbonTextLanguageControl(label, screenTip, superTip, keyTip, alternateLabel, null);

        public IRibbonTextLanguageControl NewLanguageControlRibbonText(
            string label,
            string screenTip,
            string superTip,
            string keyTip,
            string alternateLabel,
            string description
        ) => new RibbonTextLanguageControl(label, screenTip, superTip, keyTip, alternateLabel, description);
        #endregion
    }
}
