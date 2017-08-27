using System.Collections.Generic;
using System.Collections.Immutable;
using Microsoft.Office.Core;

namespace RubberDuck.RibbonDispatcher {
    using System;
    using static RibbonControlSize;
    using System.Runtime.InteropServices;

    using ControlSize         = RibbonControlSize;
    using LanguageStrings     = IRibbonTextLanguageControl;

    [ComVisible(true)][CLSCompliant(true)]
    public class RibbonFactory {
        internal RibbonFactory(IRibbonUI ribbonUI) {
            RibbonUI   = ribbonUI;

            _Controls  = new Dictionary<string,IRibbonCommon>();
            _Buttons   = new Dictionary<string,IRibbonButton>();
            _Toggles   = new Dictionary<string,IRibbonToggle>();
            _DropDowns = new Dictionary<string,IRibbonDropDown>();
        }

        private IRibbonUI    RibbonUI    { get; }

        private void PropertyChanged(object sender, ChangedControlEventArgs e) => RibbonUI.InvalidateControl(e.ControlId);

        private T Add<T>(T ctrl) where T:IRibbonCommon {
            _Controls.Add(ctrl.Id, ctrl);
            if (ctrl is IRibbonButton     button) _Buttons  .Add(ctrl.Id, button);
            if (ctrl is IRibbonToggle     toggle) _Toggles  .Add(ctrl.Id, toggle);
            if (ctrl is IRibbonDropDown dropDown) _DropDowns.Add(ctrl.Id, dropDown);

            ctrl.Changed += PropertyChanged;
            return ctrl;
        }

        public IReadOnlyDictionary<string,IRibbonCommon>     Controls  => _Controls.ToImmutableDictionary();       
        private IDictionary<string,IRibbonCommon>           _Controls  { get; }
        public IRibbonCommon NewRibbonCommon(
            string          id,
            LanguageStrings strings=null, 
            bool            visible=true,
            bool            enabled=true,
            ControlSize     size   =RibbonControlSizeLarge
        ) => Add(new RibbonCommon(id, strings, visible, enabled, size));

        public IReadOnlyDictionary<string,IRibbonButton>     Buttons   => _Buttons.ToImmutableDictionary();
        private IDictionary<string,IRibbonButton>           _Buttons   { get; }
        public IRibbonButton NewRibbonButton(
            string          id,
            LanguageStrings strings=null, 
            bool            visible=true,
            bool            enabled=true,
            ControlSize     size   =RibbonControlSizeLarge
        ) => Add(new RibbonButton(id, strings, visible, enabled, size));

        public IReadOnlyDictionary<string,IRibbonToggle>     Toggles   => _Toggles.ToImmutableDictionary();
        private IDictionary<string,IRibbonToggle>           _Toggles   { get; }
        public IRibbonToggle NewRibbonToggle(
            string          id,
            LanguageStrings strings=null, 
            bool            visible=true,
            bool            enabled=true,
            ControlSize     size   =RibbonControlSizeLarge
        ) => Add(new RibbonToggle(id, strings, visible, enabled, size));

        public IReadOnlyDictionary<string,IRibbonDropDown>   DropDowns => _DropDowns.ToImmutableDictionary();
        private IDictionary<string,IRibbonDropDown>         _DropDowns { get; }
        public IRibbonDropDown NewRibbonDropDown(
            string          id,
            LanguageStrings strings=null, 
            bool            visible=true,
            bool            enabled=true,
            ControlSize     size   =RibbonControlSizeLarge
        ) => Add(new RibbonDropDown(id, strings, visible, enabled, size));

        public LanguageStrings NewLanguageControlRibbonText(
            string label,
            string screenTip      = null,
            string superTip       = null,
            string keyTip         = null,
            string alternateLabel = null,
            string description    = null
        ) => new RibbonTextLanguageControl(label, screenTip, superTip, keyTip, alternateLabel, description);

    }
}
