using System.Diagnostics.CodeAnalysis;

using System.Windows.Forms;
using Microsoft.Office.Core;

namespace RubberDuck.RibbonDispatcher {
    using System;

    internal class StandardButtonsModel {
        [SuppressMessage("Microsoft.Globalization", "CA1300:SpecifyMessageBoxOptions")]
        [SuppressMessage("Microsoft.Globalization", "CA1303:Do not pass literals as localized parameters", MessageId = "System.Windows.Forms.MessageBox.Show(System.String,System.String,System.Windows.Forms.MessageBoxButtons)")]
        public StandardButtonsModel(RibbonFactory factory,  EventHandler<ClickedEventArgs> showAdvancedAction) {
            Group   = factory.NewRibbonCommon("StandardButtonsGroup");
            Group.SetText(factory.NewLanguageControlRibbonText("Standard Buttons", null, null, null));

            Button1 = factory.NewRibbonButton("MyButton1");
            Button1.Clicked += (s,e) => MessageBox.Show("Standard Button #1 pressed.", (s as IRibbonCommon)?.Id??"", MessageBoxButtons.OK);
            Button1.AsRibbonControl.SetText(factory.NewLanguageControlRibbonText("Std. Button 1", null, null, null));

            Button2 = factory.NewRibbonButton("MyButton2");
            Button2.Clicked += (s,e) => MessageBox.Show("Standard Button #2 pressed.", (s as IRibbonCommon)?.Id??"", MessageBoxButtons.OK);
            Button2.AsRibbonControl.SetText(factory.NewLanguageControlRibbonText("Std. Button 2", null, null, null));

            Toggle1 = factory.NewRibbonToggle("ShowAdvancedToggle");
            Toggle1.Clicked += showAdvancedAction;
            Toggle1.AsRibbonControl.SetText(factory.NewLanguageControlRibbonText("Show Custom", "Toggle Custom Buttons", "Toggles display of the Custom Buttons group.", null, "Hide Custom"));
        }

        public IRibbonCommon Group   { get; }
        public IRibbonButton Button1 { get; }
        public IRibbonButton Button2 { get; }
        public IRibbonToggle Toggle1 { get; }
    }
}
