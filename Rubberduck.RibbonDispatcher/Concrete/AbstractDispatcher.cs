using System;
using System.Globalization;
using System.Reflection;
using System.Resources;
using System.Runtime.InteropServices;

using Microsoft.Office.Core;

using Rubberduck.RibbonDispatcher.ControlDecorators;
using Rubberduck.RibbonDispatcher.AbstractCOM;

namespace Rubberduck.RibbonDispatcher.Concrete {

    /// <summary>Implementation of (all) the callbacks for the Fluent Ribbon; for .NET clients.</summary>
    /// <remarks>
    /// DOT NET clients are expected to find it more convenient to inherit their View 
    /// Model class from {AbstractDispatcher} than to compose against an instance of 
    /// {RibbonViewModel}. COM clients will most likely find the reverse true. 
    /// 
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
    [ComDefaultInterface(typeof(IRibbonViewModel))]
    [Guid(RubberduckGuid.AbstractDispatcher)]
    public abstract class AbstractDispatcher : IRibbonViewModel {
        /// <summary>TODO</summary>
        protected void InitializeRibbonFactory(IRibbonUI RibbonUI, ResourceManager ResourceManager) 
            => _ribbonFactory = new RibbonFactory(RibbonUI, ResourceManager);

        /// <inheritdoc/>
        public  IRibbonFactory RibbonFactory => _ribbonFactory; private RibbonFactory _ribbonFactory;

        /// <summary>All of the defined controls.</summary>
        private IRibbonCommon        Controls    (string controlId) => _ribbonFactory.Controls.GetOrDefault(controlId);
        /// <summary>All of the defined controls implementing the {ISizeableDecorator} interface.</summary>
        private ISizeableDecorator   Sizeables   (string controlId) => _ribbonFactory.Sizeables.GetOrDefault(controlId);
        /// <summary>All of the defined controls implementing the {IActionableDecorator} interface.</summary>
        private IActionableDecorator Actionables (string controlId) => _ribbonFactory.Actionables.GetOrDefault(controlId);
        /// <summary>All of the defined controls implementing the {IToggleableDecorator} interface.</summary>
        private IToggleableDecorator Toggleables (string controlId) => _ribbonFactory.Toggleables.GetOrDefault(controlId);
        /// <summary>All of the defined controls implementing the {ISelectableDecorator} interface.</summary>
        private ISelectableDecorator Selectables (string controlId) => _ribbonFactory.Selectables.GetOrDefault(controlId);
        /// <summary>All of the defined controls implementing the {IImageableDecorator} interface.</summary>
        private IImageableDecorator  Imageables  (string controlId) => _ribbonFactory.Imageables.GetOrDefault(controlId);

        /// <inheritdoc/>
        public string GetDescription(IRibbonControl Control)
            => Controls(Control?.Id)?.Description ?? Unknown(Control);
        /// <inheritdoc/>
        public bool   GetEnabled(IRibbonControl Control)
            => Controls(Control?.Id)?.IsEnabled ?? false;
        /// <inheritdoc/>
        public string GetKeyTip(IRibbonControl Control)
            => Controls(Control?.Id)?.KeyTip ?? "??";
        /// <inheritdoc/>
        public string GetLabel(IRibbonControl Control)
            => Controls(Control?.Id)?.Label ?? Unknown(Control);
        /// <inheritdoc/>
        public string GetScreenTip(IRibbonControl Control)
            => Controls(Control?.Id)?.ScreenTip ?? Unknown(Control);
        /// <inheritdoc/>
        public string GetSuperTip(IRibbonControl Control)
            => Controls(Control?.Id)?.SuperTip ?? Unknown(Control);
        /// <inheritdoc/>
        public bool   GetVisible(IRibbonControl Control)
            => Controls(Control?.Id)?.IsVisible ?? true;

        /// <inheritdoc/>
        public RdControlSize GetSize(IRibbonControl Control)
            => Sizeables(Control?.Id)?.Size ?? RdControlSize.rdLarge;

        /// <inheritdoc/>
        public object GetImage(IRibbonControl Control)
            => Imageables(Control?.Id)?.Image ?? "MacroSecurity";
        /// <inheritdoc/>
        public bool   GetShowImage(IRibbonControl Control)
            => Imageables(Control?.Id)?.ShowImage ?? true;
        /// <inheritdoc/>
        public bool   GetShowLabel(IRibbonControl Control)
            => Imageables(Control?.Id)?.ShowLabel ?? true;

        /// <inheritdoc/>
        public bool   GetPressed(IRibbonControl Control)
            => Toggleables(Control?.Id)?.IsPressed ?? false;
        /// <inheritdoc/>
        public void   OnActionToggle(IRibbonControl Control, bool Pressed)
            => Toggleables(Control?.Id)?.OnActionToggle(Pressed);

        /// <inheritdoc/>
        public void   OnAction(IRibbonControl Control) => Actionables(Control?.Id)?.OnAction();

        /// <inheritdoc/>
        public string GetSelectedItemId(IRibbonControl Control)
            => Selectables(Control?.Id)?.SelectedItemId;
        /// <inheritdoc/>
        public int    GetSelectedItemIndex(IRibbonControl Control)
            => Selectables(Control?.Id)?.SelectedItemIndex ?? 0;
        /// <inheritdoc/>
        public void   OnActionDropDown(IRibbonControl Control, string SelectedId, int SelectedIndex)
            => Selectables(Control?.Id)?.OnActionDropDown(SelectedId, SelectedIndex);
 
        /// <inheritdoc/>
        public int    GetItemCount(IRibbonControl Control)
            => Selectables(Control?.Id)?.ItemCount ?? 0;
        /// <inheritdoc/>
        public string GetItemId(IRibbonControl Control, int Index)
            => Selectables(Control?.Id)?.ItemId(Index) ?? "";
        /// <inheritdoc/>
        public string GetItemLabel(IRibbonControl Control, int Index)
            => Selectables(Control?.Id)?.ItemLabel(Index) ?? Unknown(Control);
        /// <inheritdoc/>
        public string GetItemScreenTip(IRibbonControl Control, int Index)
            => Selectables(Control?.Id)?.ItemScreenTip(Index) ?? Unknown(Control);
        /// <inheritdoc/>
        public string GetItemSuperTip(IRibbonControl Control, int Index)
            => Selectables(Control?.Id)?.ItemSuperTip(Index) ?? Unknown(Control);

        /// <inheritdoc/>
        public object GetItemImage(IRibbonControl Control, int Index)
            => Selectables(Control?.Id)?.ItemImage(Index) ?? "MacroSecurity";
        /// <inheritdoc/>
        public bool   GetItemShowImage(IRibbonControl Control, int Index)
            => Selectables(Control?.Id)?.ItemShowImage(Index) ?? true;
        /// <inheritdoc/>
        public bool   GetItemShowLabel(IRibbonControl Control, int Index)
            => Selectables(Control?.Id)?.ItemShowLabel(Index) ?? true;

        //private ISelectableItem DropDownItem(IRibbonControl control, int index)
        //    => DropDowns(control?.Id)[index];

        private static string Unknown(IRibbonControl Control) 
            => string.Format(CultureInfo.InvariantCulture, $"'{Control?.Id??""}' unknown");

        /// <summary>TODO</summary>
        protected static ResourceManager GetResourceManager(string ResourceSetName) 
            => new ResourceManager(ResourceSetName, Assembly.GetExecutingAssembly());
    }
}
