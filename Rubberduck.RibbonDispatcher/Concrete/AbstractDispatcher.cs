////////////////////////////////////////////////////////////////////////////////////////////////////
//                                Copyright (c) 2017 Pieter Geerkens                              //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.Globalization;
using System.Reflection;
using System.Resources;
using System.Runtime.InteropServices;

using Microsoft.Office.Core;

using Rubberduck.RibbonDispatcher.Abstract;
using Rubberduck.RibbonDispatcher.AbstractCOM;

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
    [CLSCompliant(true)]
    [ComDefaultInterface(typeof(IAbstractDispatcher))]
    [Guid(RubberduckGuid.AbstractDispatcher)]
    public abstract class AbstractDispatcher : IAbstractDispatcher {
        /// <summary>TODO</summary>
        protected void InitializeRibbonFactory(IRibbonUI ribbonUI, ResourceManager resourceManager) 
            => _ribbonFactory = new RibbonFactory(ribbonUI, resourceManager);

        /// <inheritdoc/>
        public  IRibbonFactory RibbonFactory => _ribbonFactory; private RibbonFactory _ribbonFactory;

        /// <summary>TODO</summary>
        private IRibbonCommon  Controls   (string controlId) => _ribbonFactory.Controls.GetOrDefault(controlId);
        /// <summary>TODO</summary>
        private IActionItem    Actions    (string controlId) => _ribbonFactory.Buttons.GetOrDefault(controlId);
        /// <summary>TODO</summary>
        private IToggleItem    Toggles    (string controlId) => _ribbonFactory.Toggles.GetOrDefault(controlId);
        /// <summary>TODO</summary>
        private IDropDownItem  DropDowns  (string controlId) => _ribbonFactory.DropDowns.GetOrDefault(controlId);
        /// <summary>TODO</summary>
        private IImageableItem Imageables (string controlId) => _ribbonFactory.Imageables.GetOrDefault(controlId);

        /// <inheritdoc/>
        public string GetDescription(IRibbonControl control)
            => Controls(control?.Id)?.Description ?? Unknown(control);
        /// <inheritdoc/>
        public bool   GetEnabled(IRibbonControl control)
            => Controls(control?.Id)?.IsEnabled ?? false;
        /// <inheritdoc/>
        public string GetKeyTip(IRibbonControl control)
            => Controls(control?.Id)?.KeyTip ?? "??";
        /// <inheritdoc/>
        public string GetLabel(IRibbonControl control)
            => Controls(control?.Id)?.Label ?? Unknown(control);
        /// <inheritdoc/>
        public string GetScreenTip(IRibbonControl control)
            => Controls(control?.Id)?.ScreenTip ?? Unknown(control);
        /// <inheritdoc/>
        public string GetSuperTip(IRibbonControl control)
            => Controls(control?.Id)?.SuperTip ?? Unknown(control);
        /// <inheritdoc/>
        public bool   GetVisible(IRibbonControl control)
            => Controls(control?.Id)?.IsVisible ?? true;

        /// <inheritdoc/>
        public RdControlSize GetSize(IRibbonControl control)
            => Controls(control?.Id)?.Size ?? RdControlSize.rdLarge;

        /// <inheritdoc/>
        public object GetImage(IRibbonControl control)
            => Imageables(control?.Id)?.Image ?? "MacroSecurity";
        /// <inheritdoc/>
        public bool   GetShowImage(IRibbonControl control)
            => Imageables(control?.Id)?.ShowImage ?? true;
        /// <inheritdoc/>
        public bool   GetShowLabel(IRibbonControl control)
            => Imageables(control?.Id)?.ShowLabel ?? true;

        /// <inheritdoc/>
        public bool   GetPressed(IRibbonControl control)
            => Toggles(control?.Id)?.IsPressed ?? false;
        /// <inheritdoc/>
        public void   OnActionToggle(IRibbonControl control, bool pressed)
            => Toggles(control?.Id)?.OnActionToggle(pressed);

        /// <inheritdoc/>
        public void   OnAction(IRibbonControl control) => Actions(control?.Id)?.OnAction();

        /// <inheritdoc/>
        public string GetSelectedItemId(IRibbonControl control)
            => DropDowns(control?.Id)?.SelectedItemId;
        /// <inheritdoc/>
        public int    GetSelectedItemIndex(IRibbonControl control)
            => DropDowns(control?.Id)?.SelectedItemIndex ?? 0;
        /// <inheritdoc/>
        public void   OnActionDropDown(IRibbonControl control, string selectedId, int selectedIndex)
            => DropDowns(control?.Id)?.OnActionDropDown(selectedId, selectedIndex);
 
        /// <inheritdoc/>
        public int    GetItemCount(IRibbonControl control)
            => DropDowns(control?.Id)?.ItemCount ?? 0;
        /// <inheritdoc/>
        public string GetItemId(IRibbonControl control, int index)
            => DropDowns(control?.Id)?.ItemId(index) ?? "";
        /// <inheritdoc/>
        public string GetItemLabel(IRibbonControl control, int index)
            => DropDowns(control?.Id)?.ItemLabel(index) ?? Unknown(control);
        /// <inheritdoc/>
        public string GetItemScreenTip(IRibbonControl control, int index)
            => DropDowns(control?.Id)?.ItemScreenTip(index) ?? Unknown(control);
        /// <inheritdoc/>
        public string GetItemSuperTip(IRibbonControl control, int index)
            => DropDowns(control?.Id)?.ItemSuperTip(index) ?? Unknown(control);

        /// <inheritdoc/>
        public object GetItemImage(IRibbonControl control, int index)
            => DropDowns(control?.Id)?.ItemImage(index) ?? "MacroSecurity";
        /// <inheritdoc/>
        public bool   GetItemShowImage(IRibbonControl control, int index)
            => DropDowns(control?.Id)?.ItemShowImage(index) ?? true;
        /// <inheritdoc/>
        public bool   GetItemShowLabel(IRibbonControl control, int index)
            => DropDowns(control?.Id)?.ItemShowLabel(index) ?? true;

        //private ISelectableItem DropDownItem(IRibbonControl control, int index)
        //    => DropDowns(control?.Id)[index];

        private static string Unknown(IRibbonControl control) 
            => string.Format(CultureInfo.InvariantCulture, $"Unknown control '{control?.Id??""}'");

        /// <summary>TODO</summary>
        protected static ResourceManager GetResourceManager(string resourceSetName) 
            => new ResourceManager(resourceSetName, Assembly.GetExecutingAssembly());
    }
}
