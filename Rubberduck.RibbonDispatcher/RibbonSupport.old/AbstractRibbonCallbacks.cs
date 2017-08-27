using System;
using stdole;

using Microsoft.Office.Core;
using System.Windows.Forms;

namespace RubberDuck.RibbonSupport {
    using Office = Microsoft.Office.Core;

    //Create callback methods here. For more information about adding callback methods, visit https://go.microsoft.com/fwlink/?LinkID=271226
    public class AbstractRibbonCallbacks {
        protected AbstractRibbonCallbacks () { }

        //public RibbonFactory    RibbonFactory    { get; private set; }

        //public IRibbonCommon    Controls  (string controlId) => RibbonFactory.Controls[controlId];
        //public IRibbonButton    Buttons   (string controlId) => RibbonFactory.Buttons[controlId];
        //public IRibbonToggle    Toggles   (string controlId) => RibbonFactory.Toggles[controlId];
        //public IRibbonDropdown  Dropdowns (string controlId) => RibbonFactory.Dropdowns[controlId];

        //protected Office.IRibbonUI RibbonUI => RibbonFactory?.RibbonUI;

        //public virtual void Ribbon_Load(Office.IRibbonUI ribbonUI) {
        //    RibbonFactory = new RibbonFactory(ribbonUI);
        //    MessageBox.Show ("Ribbon_Load", "Break!",MessageBoxButtons.OK);
        //}

        //// All controls (almost) including Groups
        //public string            GetDescription(IRibbonControl control) => Controls(control.Id).Description;
        //public bool              GetEnabled(IRibbonControl control)     => Controls(control.Id).Enabled;
        //public IPictureDisp      GetImage(IRibbonControl control)       => Controls(control.Id).Image;
        //public string            GetImageMso(IRibbonControl control)    => Controls(control.Id).ImageMso;   // obsoleted in Office 2010
        //public string            GetKeytip(IRibbonControl control)      => Controls(control.Id).KeyTip;
        //public string            GetLabel(IRibbonControl control)       => Controls(control.Id).Label;
        //public RibbonControlSize GetSize(IRibbonControl control)        => Controls(control.Id).Size;
        //public string            GetScreentip(IRibbonControl control)   => Controls(control.Id).ScreenTip;
        //public string            GetSupertip(IRibbonControl control)    => Controls(control.Id).SuperTip;
        //public bool              GetVisible(IRibbonControl control)     => Controls(control.Id).Visible;

        //// Buttons
        //public bool              GetShowImage(IRibbonControl control)   => Buttons(control.Id).ShowImage;
        //public bool              GetShowLabel(IRibbonControl control)   => Buttons(control.Id).ShowLabel;
        //public void              OnAction(IRibbonControl control)       => Buttons(control.Id).OnAction();

        //// Toggles: checkBoxes & toggleButtons
        //public bool              GetPressed(IRibbonControl control)             => Toggles(control.Id).IsPressed;
        //public void              OnAction(IRibbonControl control, bool pressed) => Toggles(control.Id).OnAction(pressed);
    }
}
