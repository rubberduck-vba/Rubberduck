using System.Diagnostics.CodeAnalysis;

using System.Windows.Forms;
using Microsoft.Office.Core;

namespace RubberDuck.Ribbon {
    using static RibbonText;
    internal class StandardButtonsModel {
        [SuppressMessage("Microsoft.Globalization", "CA1303:Do not pass literals as localized parameters", MessageId = "System.Windows.Forms.MessageBox.Show(System.String,System.String,System.Windows.Forms.MessageBoxButtons)")]
        [SuppressMessage("Microsoft.Globalization", "CA1300:SpecifyMessageBoxOptions")]
        public StandardButtonsModel(RibbonFactory factory) {
            _factory = factory;

            Group   = _factory.NewRibbonCommon("StandardButtonsGroup");
            Group.SetText(NewLanguageControlRibbonText("Standard Buttons", null, null, null));

            Button1 = _factory.NewRibbonButton("MyButton1");
            Button1.Clicked += (s,e) => MessageBox.Show("Standard Button #1 pressed.", (s as IRibbonCommon)?.ID??"", MessageBoxButtons.OK);
            Button1.AsRibbonControl.SetText(NewLanguageControlRibbonText("Std. Button 1", null, null, null));

            Button2 = _factory.NewRibbonButton("MyButton2");
            Button2.Clicked += (s,e) => MessageBox.Show("Standard Button #2 pressed.", (s as IRibbonCommon)?.ID??"", MessageBoxButtons.OK);
            Button2.AsRibbonControl.SetText(NewLanguageControlRibbonText("Std. Button 2", null, null, null));
        }
        RibbonFactory _factory;

        public IRibbonCommon Group   { get; }
        public IRibbonButton Button1 { get; }
        public IRibbonButton Button2 { get; }
    }
}
