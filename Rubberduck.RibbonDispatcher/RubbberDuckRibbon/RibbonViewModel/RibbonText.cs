using Microsoft.Office.Core;
using System;
using System.Runtime.InteropServices;

namespace RubberDuck.Ribbon {
    [ComVisible(true)][CLSCompliant(true)]
    public static class RibbonText {
        public static IRibbonTextLanguageControl NewLanguageControlRibbonText(
            string label,
            string screentip,
            string supertip,
            string keytip
        ) => new RibbonTextLanguageControl(label,screentip,supertip,keytip,null,null);
        public static IRibbonTextLanguageControl NewLanguageControlRibbonText(
            string label,
            string screentip,
            string supertip,
            string keytip,
            string alternateLabel
        ) => new RibbonTextLanguageControl(label,screentip,supertip,keytip,alternateLabel,null);
        public static IRibbonTextLanguageControl NewLanguageControlRibbonText(
            string label,
            string screentip,
            string supertip,
            string keytip,
            string alternateLabel,
            string description
        ) => new RibbonTextLanguageControl(label,screentip,supertip,keytip,alternateLabel,description);
    }
}
