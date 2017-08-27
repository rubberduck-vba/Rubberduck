using System.Diagnostics.CodeAnalysis;

using System.Windows.Forms;
using Rubberduck.RibbonDispatcher.Abstract;

namespace Rubberduck.RibbonDispatcher.Excel2013 {
    using System;

    internal class StandardButtonsViewModel {
        [SuppressMessage("Microsoft.Globalization", "CA1300:SpecifyMessageBoxOptions")]
        [SuppressMessage("Microsoft.Globalization", "CA1303:Do not pass literals as localized parameters", MessageId = "System.Windows.Forms.MessageBox.Show(System.String,System.String,System.Windows.Forms.MessageBoxButtons)")]
        public StandardButtonsViewModel(RibbonFactory factory,  EventHandler<ClickedEventArgs> showAdvancedAction) {
            Group   = factory.NewRibbonGroup("StandardButtonsGroup",
                    factory.NewLanguageControlRibbonText("Standard Buttons", null, null, null));

            Button1 = factory.NewRibbonButton("MyButton1",
                    factory.NewLanguageControlRibbonText("Std. Button 1", null, null, null));
            Button1.Clicked += (s,e) => MessageBox.Show("Standard Button #1 pressed.", (s as IRibbonCommon)?.Id??"", MessageBoxButtons.OK);

            Button2 = factory.NewRibbonButton("MyButton2",
                    factory.NewLanguageControlRibbonText("Std. Button 2", null, null, null));
            Button2.Clicked += (s,e) => MessageBox.Show("Standard Button #2 pressed.", (s as IRibbonCommon)?.Id??"", MessageBoxButtons.OK);

            Toggle1 = factory.NewRibbonToggle("ShowAdvancedToggle",
                factory.NewLanguageControlRibbonText("Show Custom", "Toggle Custom Buttons", "Toggles display of the Custom Buttons group.", null, "Hide Custom"));
            Toggle1.Clicked += showAdvancedAction;
        }

        public IRibbonCommon Group   { get; }
        public IRibbonButton Button1 { get; }
        public IRibbonButton Button2 { get; }
        public IRibbonToggle Toggle1 { get; }
    }
}
