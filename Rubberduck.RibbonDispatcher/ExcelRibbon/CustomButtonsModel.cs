using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;

using System.Windows.Forms;
using Microsoft.Office.Core;

namespace RubberDuck.RibbonDispatcher {
    using static RibbonControlSize;

    internal class CustomButtonsModel {
        [SuppressMessage("Microsoft.Globalization", "CA1300:SpecifyMessageBoxOptions")]
        [SuppressMessage("Microsoft.Globalization", "CA1303:Do not pass literals as localized parameters", MessageId = "System.Windows.Forms.MessageBox.Show(System.String,System.String,System.Windows.Forms.MessageBoxButtons)")]
        public CustomButtonsModel(RibbonFactory factory) {
            Group   = factory.NewRibbonCommon("CustomButtonsGroup",factory.NewLanguageControlRibbonText("Custom Apps", null, null, null),false);

            Button1 = factory.NewRibbonButton("AppLaunchButton1",factory.NewLanguageControlRibbonText("Button #1", null, null, null));
            Button1.Clicked += (s,e) => MessageBox.Show("Pressed Custom Button #1 pressed.", (s as IRibbonCommon)?.Id??"", MessageBoxButtons.OK);

            Button2 = factory.NewRibbonButton("AppLaunchButton2",factory.NewLanguageControlRibbonText("Button #2", null, null, null));
            Button2.Clicked += (s,e) => MessageBox.Show("Custom Button #2 pressed.", (s as IRibbonCommon)?.Id??"", MessageBoxButtons.OK);
 
            Button3 = factory.NewRibbonButton("AppLaunchButton3",factory.NewLanguageControlRibbonText("Button #3", null, null, null));
            Button3.Clicked += (s,e) => MessageBox.Show("Custom Button #3 pressed.", (s as IRibbonCommon)?.Id??"", MessageBoxButtons.OK);
 
            SizeToggle = factory.NewRibbonToggle("SizeToggle",factory.NewLanguageControlRibbonText("Show Small", "Toggles Button Size", 
                                "Toggles between large and small buttons for this ribbon group.", null, "Show Large"));
            SizeToggle.Clicked += (s,e) => SetIsLarge( ! e.IsPressed);
        }

        public void SetVisible(bool value) { Group.Visible = value; }
        public void SetIsLarge(bool value) { 
            _isLarge = value;
            foreach (IRibbonCommon b in Buttons) b.Size = _isLarge ? RibbonControlSizeLarge : RibbonControlSizeRegular;
            Group.OnChanged();
        } bool _isLarge;

        private IList<IRibbonButton> Buttons => new List<IRibbonButton>() {Button1, Button2, Button3 };
        public IRibbonCommon Group      { get; }
        public IRibbonButton Button1    { get; }
        public IRibbonButton Button2    { get; }
        public IRibbonButton Button3    { get; }
        public IRibbonToggle SizeToggle { get; }
    }
}
