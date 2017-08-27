using System.Diagnostics.CodeAnalysis;

using System.Windows.Forms;
using Microsoft.Office.Core;

namespace RubberDuck.Ribbon {
    using static RibbonText;
    internal class CustomButtonsModel {
        [SuppressMessage("Microsoft.Globalization", "CA1303:Do not pass literals as localized parameters", MessageId = "System.Windows.Forms.MessageBox.Show(System.String,System.String,System.Windows.Forms.MessageBoxButtons)")]
        [SuppressMessage("Microsoft.Globalization", "CA1300:SpecifyMessageBoxOptions")]
        public CustomButtonsModel(RibbonFactory factory) {
            _factory = factory;

            Group   = _factory.NewRibbonCommon("CustomButtonsGroup");
            Group.SetText(NewLanguageControlRibbonText("Custom Apps", null, null, null));

            Button1 = _factory.NewRibbonButton("AppLaunchButton1");
            Button1.AsRibbonControl.SetText(NewLanguageControlRibbonText("Button #1", null, null, null));
            Button1.Clicked += (s,e) => MessageBox.Show("Pressed Custom Button #1 pressed.", (s as IRibbonCommon)?.ID??"", MessageBoxButtons.OK);

            Button2 = _factory.NewRibbonButton("AppLaunchButton2");
            Button2.AsRibbonControl.SetText(NewLanguageControlRibbonText("Button #2", null, null, null));
            Button2.Clicked += (s,e) => MessageBox.Show("Custom Button #2 pressed.", (s as IRibbonCommon)?.ID??"", MessageBoxButtons.OK);
 
            Button3 = _factory.NewRibbonButton("AppLaunchButton3");
            Button3.AsRibbonControl.SetText(NewLanguageControlRibbonText("Button #3", null, null, null));
            Button3.Clicked += (s,e) => MessageBox.Show("Custom Button #3 pressed.", (s as IRibbonCommon)?.ID??"", MessageBoxButtons.OK);
        }
        RibbonFactory _factory;

        IRibbonCommon Group   { get; }
        IRibbonButton Button1 { get; }
        IRibbonButton Button2 { get; }
        IRibbonButton Button3 { get; }
    }
}
