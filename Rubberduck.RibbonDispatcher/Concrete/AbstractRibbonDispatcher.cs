////////////////////////////////////////////////////////////////////////////////////////////////////
//                                Copyright (c) 2017 Pieter Geerkens                              //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.Diagnostics.CodeAnalysis;
using System.Drawing;
using System.Globalization;
using System.Reflection;
using System.Resources;
using System.Runtime.InteropServices;
using System.Windows.Forms;

using stdole;
using Microsoft.Office.Core;

using Rubberduck.RibbonDispatcher.Abstract;

namespace Rubberduck.RibbonDispatcher.Concrete {

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
    [Serializable]
    [ComVisible(true)]
    [Guid("2B43D4D0-A674-40CD-B465-FA715DEB74E9")]
    [CLSCompliant(true)]
    public abstract class AbstractRibbonDispatcher {
        /// <summary>TODO</summary>
        protected void InitializeRibbonFactory(IRibbonUI ribbonUI, ResourceManager resourceManager) 
            => RibbonFactory = new RibbonFactory(ribbonUI, resourceManager);

        /// <summary>TODO</summary>
        public RibbonFactory       RibbonFactory { get; private set; }

        /// <summary>TODO</summary>
        public IRibbonCommon       Controls     (string controlId) => RibbonFactory.Controls.GetOrDefault(controlId);
        /// <summary>TODO</summary>
        public IActionItem         Actions      (string controlId) => RibbonFactory.Buttons.GetOrDefault(controlId);
        /// <summary>TODO</summary>
        public IToggleItem         Toggles      (string controlId) => RibbonFactory.Toggles.GetOrDefault(controlId);
        /// <summary>TODO</summary>
        public IRibbonDropDown     DropDowns    (string controlId) => RibbonFactory.DropDowns.GetOrDefault(controlId);
        /// <summary>TODO</summary>
        public IImageableItem      Imageables   (string controlId) => RibbonFactory.Imageables.GetOrDefault(controlId);

        /// <summary>Call back for GetDescription events from ribbon elements.</summary>
        public string              GetDescription (IRibbonControl control) => Controls(control?.Id)?.Description ?? Unknown(control);
        /// <summary>Call back for GetKeyTip events from ribbon elements.</summary>
        public string              GetKeyTip      (IRibbonControl control) => Controls(control?.Id)?.KeyTip      ?? "??";
        /// <summary>Call back for GetLabel events from ribbon elements.</summary>
        public string              GetLabel       (IRibbonControl control) => Controls(control?.Id)?.Label       ?? Unknown(control);
        /// <summary>Call back for GetScreenTip events from ribbon elements.</summary>
        public string              GetScreenTip   (IRibbonControl control) => Controls(control?.Id)?.ScreenTip   ?? Unknown(control);
        /// <summary>Call back for GetSuperTip events from ribbon elements.</summary>
        public string              GetSuperTip    (IRibbonControl control) => Controls(control?.Id)?.SuperTip    ?? Unknown(control);

        /// <summary>Call back for GetEnabled events from ribbon elements.</summary>
        public bool                GetEnabled     (IRibbonControl control) => Controls(control?.Id)?.IsEnabled   ?? false;
        /// <summary>Call back for GetSize events from ribbon elements.</summary>
        public MyRibbonControlSize GetSize        (IRibbonControl control) => Controls(control?.Id)?.Size        ?? MyRibbonControlSize.Large;
        /// <summary>Call back for GetVisible events from ribbon elements.</summary>
        public bool                GetVisible     (IRibbonControl control) => Controls(control?.Id)?.IsVisible   ?? true;

        /// <summary>Call back for GetImage events from ribbon elements.</summary>
        public object              GetImage       (IRibbonControl control) => Imageables(control?.Id)?.Image;
        /// <summary>Call back for GetShowImage events from ribbon elements.</summary>
        public bool                GetShowImage   (IRibbonControl control) => Imageables(control?.Id)?.ShowImage ?? false;
        /// <summary>Call back for GetShowLabel events from ribbon elements.</summary>
        public bool                GetShowLabel   (IRibbonControl control) => Imageables(control?.Id)?.ShowLabel ?? true;

        /// <summary>Call back for OnAction events from the button ribbon elements.</summary>
        public void OnAction(IRibbonControl control)                       => Actions(control?.Id)?.OnAction();

        /// <summary>Call back for GetPressed events from the checkBox and toggleButton ribbon elements.</summary>
        public bool GetPressed(IRibbonControl control)                     => Toggles(control?.Id)?.IsPressed    ?? false;
        /// <summary>Call back for OnAction events from the checkBox and toggleButton ribbon elements.</summary>
        public void OnActionToggle(IRibbonControl control, bool pressed)   => Toggles(control?.Id)?.OnAction(pressed);


        private static string Unknown(IRibbonControl control) 
            => string.Format(CultureInfo.InvariantCulture, $"Unknown control '{control?.Id??""}'");

        /// <summary>TODO</summary>
        protected static ResourceManager GetResourceManager(string resourceSetName) 
            => new ResourceManager(resourceSetName, Assembly.GetExecutingAssembly());

        /// <summary>TODO</summary>
        protected static IPictureDisp    GetResourceImage(string imageName, ResourceManager resourceManager) {
            using (var image = resourceManager?.GetObject(imageName, CultureInfo.InvariantCulture) as Image) {
                return (image == null) ? null : PictureConverter.ImageToPictureDisp(image);
            }
        }

        /// <summary>TODO</summary>
        protected static IPictureDisp    GetResourceIcon(string iconName, ResourceManager resourceManager) {
            using (var icon = resourceManager?.GetObject(iconName, CultureInfo.InvariantCulture) as Icon) {
                return icon == null ? null : PictureConverter.IconToPictureDisp(icon);
            }
        }

        [SuppressMessage("Microsoft.Performance", "CA1812:AvoidUninstantiatedInternalClasses")]
        internal class PictureConverter : AxHost {
            internal PictureConverter() : base(String.Empty) { }

            public static IPictureDisp ImageToPictureDisp(Image image) => GetIPictureDispFromPicture(image) as IPictureDisp;

            public static IPictureDisp IconToPictureDisp(Icon icon)    => ImageToPictureDisp(icon.ToBitmap());
        }
    }
}
