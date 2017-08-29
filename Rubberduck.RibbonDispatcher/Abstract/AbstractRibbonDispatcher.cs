using System;
using System.Runtime.InteropServices;

using stdole;
using Microsoft.Office.Core;

namespace Rubberduck.RibbonDispatcher.Abstract {
    using static RibbonControlSize;
    using static System.Globalization.CultureInfo;

    /// <summary>(All) the callbacks for the Fluent Ribbon.</summary>
    /// <remarks>
    /// The callback names are chosen to be identical to the corresponding xml tag in
    /// the Ribbon schema, except for:
    ///  - PascalCase instead of camelCase; and
    ///  - In some instances, a disambiguating usage suffix such as OnActionToggle(,)
    ///    instead of a plain OnAction(,).
    ///    
    /// Whenever possible the Dispatcher will return default values acceptable to OFFICE
    /// even if the Control.Id supplied to a callback is unknown. These defaults are
    /// chosen to maximize visibility for the unknown control, but disable its functionality.
    /// This is believed to support the principle of 'least surprise', given the OFFICE 
    /// Ribbon's propensity to fail, silently and/or fatally, at the slightest provocation.
    /// </remarks>
    [ComVisible(true)]
    [CLSCompliant(true)]
    public abstract class AbstractRibbonDispatcher {      
        protected void           InitializeRibbonFactory(IRibbonUI ribbonUI) => RibbonFactory = new RibbonFactory(ribbonUI);
        protected RibbonFactory  RibbonFactory  { get; private set; }

        public IRibbonCommon     Controls       (IRibbonControl control) => RibbonFactory.Controls.GetOrDefault(control?.Id);
        public IRibbonButton     Buttons        (IRibbonControl control) => RibbonFactory.Buttons.GetOrDefault(control?.Id);
        public IRibbonToggle     Toggles        (IRibbonControl control) => RibbonFactory.Toggles.GetOrDefault(control?.Id);
        public IRibbonDropDown   DropDowns      (IRibbonControl control) => RibbonFactory.DropDowns.GetOrDefault(control?.Id);

        public string            GetDescription (IRibbonControl control) => Controls(control)?.Description ?? Unknown(control);
        public string            GetKeyTip      (IRibbonControl control) => Controls(control)?.KeyTip      ?? "??";
        public string            GetLabel       (IRibbonControl control) => Controls(control)?.Label       ?? Unknown(control);
        public string            GetScreenTip   (IRibbonControl control) => Controls(control)?.ScreenTip   ?? Unknown(control);
        public string            GetSuperTip    (IRibbonControl control) => Controls(control)?.SuperTip    ?? Unknown(control);

        public bool              GetEnabled     (IRibbonControl control) => Controls(control)?.Enabled     ?? false;
        public IPictureDisp      GetImage       (IRibbonControl control) => Controls(control)?.Image;
        public bool              GetShowImage   (IRibbonControl control) => Controls(control)?.ShowImage   ?? false;
        public bool              GetShowLabel   (IRibbonControl control) => Controls(control)?.ShowLabel   ?? true;
        public RibbonControlSize GetSize        (IRibbonControl control) => Controls(control)?.Size        ?? RibbonControlSizeLarge;
        public bool              GetVisible     (IRibbonControl control) => Controls(control)?.Visible     ?? true;

        public void OnAction(IRibbonControl control)                     => Buttons(control)?.OnAction();

        public bool GetPressed(IRibbonControl control)                   => Toggles(control)?.IsPressed    ?? false;
        public void OnActionToggle(IRibbonControl control, bool pressed) => Toggles(control)?.OnAction(pressed);

        private static string Unknown(IRibbonControl control) 
            => string.Format(InvariantCulture, $"Unknown control '{control?.Id??""}'");
    }
}
