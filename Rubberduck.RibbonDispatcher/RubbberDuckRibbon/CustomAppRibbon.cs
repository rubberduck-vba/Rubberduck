using System;
using System.Diagnostics.CodeAnalysis;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using stdole;
using Microsoft.Office.Core;

namespace RubberDuck.Ribbon {
    using static RibbonControlSize;

    [ComVisible(true)][CLSCompliant(true)]
    public class CustomAppRibbon : IRibbonExtensibility {

        public CustomAppRibbon() : base() {;}

        #region IRibbonExtensibility Members

        public string GetCustomUI(string RibbonID) => GetResourceText("RubberDuck.RibbonSupport.CustomAppRibbon.xml");
        #endregion

        // vvv Much of this really should be in an abstract base class, but Excel croaks over that for an unknown reason. vvv
        #region Ribbon Callbacks
        [SuppressMessage("Microsoft.Naming", "CA1707:IdentifiersShouldNotContainUnderscores")]
        [CLSCompliant(false)]
        public void Ribbon_Load(IRibbonUI ribbonUI) {
            RibbonFactory           = new RibbonFactory(ribbonUI);
            _standardButtonsGroup   = new StandardButtonsModel(RibbonFactory);
            _customButtonsGroup     = new CustomButtonsModel(RibbonFactory);
        }

        public RibbonFactory     RibbonFactory    { get; private set; }

        [SuppressMessage("Microsoft.Performance", "CA1823:AvoidUnusedPrivateFields")]
        StandardButtonsModel     _standardButtonsGroup;
        [SuppressMessage("Microsoft.Performance", "CA1823:AvoidUnusedPrivateFields")]
        CustomButtonsModel       _customButtonsGroup;

        public IRibbonCommon     Controls(string controlId)   => RibbonFactory.Controls[controlId];
        public IRibbonButton     Buttons(string controlId)    => RibbonFactory.Buttons[controlId];
        public IRibbonToggle     Toggles(string controlId)    => RibbonFactory.Toggles[controlId];
        public IRibbonDropdown   DropDowns(string controlId)  => RibbonFactory.Dropdowns[controlId];

        protected IRibbonUI      RibbonUI => RibbonFactory?.RibbonUI;

        // All controls (almost) including Groups
        public string            GetDescription(IRibbonControl control) => Controls(control?.Id)?.Description??"Who? Me?";
        public bool              GetEnabled(IRibbonControl control)     => Controls(control?.Id)?.Enabled??false;
        public IPictureDisp      GetImage(IRibbonControl control)       => Controls(control?.Id)?.Image;
        public string            GetImageMso(IRibbonControl control)    => Controls(control?.Id)?.ImageMso;   // obsoleted in Office 2010
        public string            GetKeyTip(IRibbonControl control)      => Controls(control?.Id)?.KeyTip??"";
        public string            GetLabel(IRibbonControl control)       => Controls(control?.Id)?.Label??"A Label";
        public RibbonControlSize GetSize(IRibbonControl control)        => Controls(control?.Id)?.Size??RibbonControlSizeLarge;
        public string            GetScreenTip(IRibbonControl control)   => Controls(control?.Id)?.ScreenTip??control?.Id??"";
        public string            GetSuperTip(IRibbonControl control)    => Controls(control?.Id)?.SuperTip??"SuperTip text";
        public bool              GetVisible(IRibbonControl control)     => Controls(control?.Id)?.Visible??true;

        // Buttons
        public bool GetShowImage(IRibbonControl control) => Buttons(control?.Id)?.ShowImage??false;
        public bool GetShowLabel(IRibbonControl control) => Buttons(control?.Id)?.ShowLabel??false;
        public void OnAction(IRibbonControl control)     => Buttons(control?.Id)?.OnAction();

        // Toggles: checkBoxes & toggleButtons
        public bool GetPressed(IRibbonControl control)             => Toggles(control?.Id)?.IsPressed??false;
        public void OnAction(IRibbonControl control, bool pressed) => Toggles(control?.Id)?.OnAction(pressed);
        #endregion
        // ^^^ Much of this really should be in an abstract base class, but Excel croaks over that for an unknown reason. ^^^

        #region Helpers

        private static string GetResourceText(string resourceName) {
            Assembly asm = Assembly.GetExecutingAssembly();
            string[] resourceNames = asm.GetManifestResourceNames();
            for(int i = 0; i < resourceNames.Length; ++i) {
                if(string.Compare(resourceName, resourceNames[i], StringComparison.OrdinalIgnoreCase) == 0) {
                    using(StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(resourceNames[i]))) {
                        if(resourceReader != null) {
                            return resourceReader.ReadToEnd();
                        }
                    }
                }
            }
            return null;
        }
        #endregion
    }
}
