using System;
using System.Runtime.InteropServices;
using Microsoft.Office.Core;

namespace RubberDuck.RibbonDispatcher {
    [ComVisible(true)][CLSCompliant(true)]
    public class RibbonText {
        public RibbonText() { ; }

        public IRibbonTextLanguageControl NewLanguageControlRibbonText(
            string label,
            string screenTip,
            string superTip,
            string keyTip
        ) => new RibbonTextLanguageControl(label, screenTip, superTip, keyTip, null,null);

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
    }
}
