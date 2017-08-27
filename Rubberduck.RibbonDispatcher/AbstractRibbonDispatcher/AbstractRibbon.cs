using System;
using System.Runtime.InteropServices;
using stdole;
using Microsoft.Office.Core;

namespace Rubberduck.RibbonDispatcher.Abstract {
    using System.Collections.Generic;
    using static RibbonControlSize;
    using static System.Globalization.CultureInfo;

    /// <summary>(All) the callbacks for the Fluent Ribbon.</summary>
    /// <remarks>
    /// The callback names are chosen to be identical to the corresponding xml tag in
    /// the Ribbon schema, except for:
    ///  - PascalCase instead of camelCase; and
    ///  - In some instances, a disambiguating usage suffix such as OnActionToggle(,)
    ///    instead of a plain OnAction(,).
    /// </remarks>
    [ComVisible(true)][CLSCompliant(true)]
    public abstract class AbstractRibbon {
        protected AbstractRibbon() : base() {;}
        
        public virtual void OnRibbonLoad(IRibbonUI ribbonUI) {
            RibbonFactory         = new RibbonFactory(ribbonUI);
        }
        protected RibbonFactory     RibbonFactory    { get; private set; }    // A property because might be exposed in future

        public IRibbonCommon     Controls(string controlId)              => GetValueOrNull(RibbonFactory.Controls, controlId);
        public IRibbonButton     Buttons(string controlId)               => GetValueOrNull(RibbonFactory.Buttons,  controlId);
        public IRibbonToggle     Toggles(string controlId)               => GetValueOrNull(RibbonFactory.Toggles,  controlId);
        public IRibbonDropDown   DropDowns(string controlId)             => GetValueOrNull(RibbonFactory.DropDowns,controlId);

        // All controls (almost) including Groups
        public string            GetDescription(IRibbonControl control)  => Controls(control?.Id)?.Description??Unknown(control);
        public string            GetKeyTip(IRibbonControl control)       => Controls(control?.Id)?.KeyTip     ??"??";
        public string            GetLabel(IRibbonControl control)        => Controls(control?.Id)?.Label      ??Unknown(control);
        public string            GetScreenTip(IRibbonControl control)    => Controls(control?.Id)?.ScreenTip  ??Unknown(control);
        public string            GetSuperTip(IRibbonControl control)     => Controls(control?.Id)?.SuperTip   ??Unknown(control);

        public bool              GetEnabled(IRibbonControl control)      => Controls(control?.Id)?.Enabled    ??false;
        public IPictureDisp      GetImage(IRibbonControl control)        => Controls(control?.Id)?.Image;
        public RibbonControlSize GetSize(IRibbonControl control)         => Controls(control?.Id)?.Size       ??RibbonControlSizeLarge;
        public bool              GetVisible(IRibbonControl control)      => Controls(control?.Id)?.Visible    ??true;

        // Button Controls
        public bool              GetShowImage(IRibbonControl control)    => Buttons(control?.Id)?.ShowImage   ??false;
        public bool              GetShowLabel(IRibbonControl control)    => Buttons(control?.Id)?.ShowLabel   ??true;
        public void              OnAction(IRibbonControl control)        => Buttons(control?.Id)?.OnAction();

        // Toggle Controls: checkBoxes & toggleButtons
        public bool GetPressed(IRibbonControl control)                   => Toggles(control?.Id)?.IsPressed   ??false;
        public void OnActionToggle(IRibbonControl control, bool pressed) => Toggles(control?.Id)?.OnAction(pressed);

        private TValue GetValueOrNull<TValue>(IReadOnlyDictionary<string,TValue> dictionary, string key)
        {
            TValue ctrl;
            dictionary.TryGetValue(key, out ctrl);
            return ctrl;
        }

        private string Unknown(IRibbonControl control) {
            return string.Format(InvariantCulture, $"Unknown control '{control?.Id??""}'");
        }
    }
}
