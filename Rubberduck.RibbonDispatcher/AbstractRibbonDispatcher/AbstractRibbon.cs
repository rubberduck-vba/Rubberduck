using System;
using System.Runtime.InteropServices;
using stdole;
using Microsoft.Office.Core;

namespace RubberDuck.RibbonDispatcher {
    using static RibbonControlSize;
    using static System.Globalization.CultureInfo;

    [ComVisible(true)][CLSCompliant(true)]
    public abstract class AbstractRibbon {
        protected AbstractRibbon() : base() {;}
        
        public virtual void OnRibbonLoad(IRibbonUI ribbonUI) {
            RibbonFactory         = new RibbonFactory(ribbonUI);
        }
        protected RibbonFactory     RibbonFactory    { get; set; }    // A property because might be exposed in future

        #region Ribbon Callbacks
        public IRibbonCommon     Controls(string controlId)              => RibbonFactory.Controls .TryGetValue(controlId,out var ctrl) ? ctrl : null;
        public IRibbonButton     Buttons(string controlId)               => RibbonFactory.Buttons  .TryGetValue(controlId,out var ctrl) ? ctrl : null;
        public IRibbonToggle     Toggles(string controlId)               => RibbonFactory.Toggles  .TryGetValue(controlId,out var ctrl) ? ctrl : null;
        public IRibbonDropDown   DropDowns(string controlId)             => RibbonFactory.DropDowns.TryGetValue(controlId,out var ctrl) ? ctrl : null;

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
        #endregion

        private string Unknown(IRibbonControl control) {
            return string.Format(InvariantCulture, $"Unknown control '{control?.Id??""}'");
        }
    }
}
