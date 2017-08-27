using System.Collections.Generic;
using System.Collections.Immutable;
using Microsoft.Office.Core;
using System.Windows.Forms;

namespace RubberDuck.RibbonSupport {
    using Office = Microsoft.Office.Core;

    using System;
    using static RibbonControlSize;
    using System.Runtime.InteropServices;

    [ComVisible(true)][CLSCompliant(true)]
    public class RibbonFactory {
        public RibbonFactory(Office.IRibbonUI ribbonUI) {
            RibbonUI = ribbonUI;

            _Controls     = new Dictionary<string,IRibbonCommon>();
            _Buttons      = new Dictionary<string,IRibbonButton>();
            _Toggles      = new Dictionary<string,IRibbonToggle>();
            _Dropdowns    = new Dictionary<string,IRibbonDropdown>();
        }

        public IRibbonUI    RibbonUI    { get; }

        private void PropertyChanged(object sender, EventArgs e) {
            if(sender is IRibbonCommon ctrl) RibbonUI.InvalidateControl(ctrl.ID);
        }

        private void Add(IRibbonCommon ctrl) {
            _Controls.Add(ctrl.ID, ctrl);
            if (ctrl is IRibbonButton)   _Buttons  .Add(ctrl.ID, ctrl as IRibbonButton);
            if (ctrl is IRibbonToggle)   _Toggles  .Add(ctrl.ID, ctrl as IRibbonToggle);
            if (ctrl is IRibbonDropdown) _Dropdowns.Add(ctrl.ID, ctrl as IRibbonDropdown);

            ctrl.Changed += PropertyChanged;
        }

        #region IRibbonCommon - all controls including Groups & Tabs
        public IReadOnlyDictionary<string,IRibbonCommon>     Controls  => _Controls.ToImmutableDictionary();       
        private IDictionary<string,IRibbonCommon>           _Controls  { get; }
        public RibbonCommon NewRibbonCommon(string id)
                => NewRibbonCommon(id, true, true, RibbonControlSizeLarge);
        public RibbonCommon NewRibbonCommon(string id, bool visible)
                => NewRibbonCommon(id, visible, true, RibbonControlSizeLarge);
        public RibbonCommon NewRibbonCommon(string id, bool visible, bool enabled)
                => NewRibbonCommon(id, visible, enabled,  RibbonControlSizeLarge);
        public RibbonCommon NewRibbonCommon(string id, bool visible, bool enabled, RibbonControlSize size) {
            var ctrl = new RibbonCommon(id, visible, enabled, size);
            Add(ctrl);
            return ctrl;
        }
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
        public IReadOnlyDictionary<string,IRibbonDropdown>   Dropdowns => _Dropdowns.ToImmutableDictionary();
        private IDictionary<string,IRibbonDropdown>         _Dropdowns { get; }
        public RibbonDropdown NewRibbonDropdown(string id)
                => NewRibbonDropdown(id, true, true, RibbonControlSizeLarge);
        public RibbonDropdown NewRibbonDropdown(string id, bool visible)
                => NewRibbonDropdown(id, visible, true, RibbonControlSizeLarge);
        public RibbonDropdown NewRibbonDropdown(string id, bool visible, bool enabled)
                => NewRibbonDropdown(id, visible, enabled,  RibbonControlSizeLarge);
        public RibbonDropdown NewRibbonDropdown(string id, bool visible, bool enabled, RibbonControlSize size) {
            var ctrl = new RibbonDropdown(id, visible, enabled, size);
            Add(ctrl);
            return ctrl;
        }
        #endregion
   }
}
