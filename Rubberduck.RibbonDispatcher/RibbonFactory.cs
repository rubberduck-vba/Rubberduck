using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics.CodeAnalysis;
using System.Runtime.InteropServices;

using Microsoft.Office.Core;

using Rubberduck.RibbonDispatcher.Abstract;
using Rubberduck.RibbonDispatcher.Concrete;
using Rubberduck.RibbonDispatcher.EventHandlers;

namespace Rubberduck.RibbonDispatcher
{
    using System.Resources;
    using static MyRibbonControlSize;

    using LanguageStrings     = IRibbonTextLanguageControl;

    /// <summary>TODO</summary>
    /// <remarks>
    /// The {SuppressMessage} attributes are left in the source here, instead of being 'fired and
    /// forgotten' to the Global Suppresion file. This explicitly points out that although "optional
    /// parameters with default values" is often regarded ad an anti-pattern for DOT NET development,
    /// COM does not support overrides. This is (believed) to be the only means of implementing
    /// the equivalent functionality in a COM-compatible way.
    /// </remarks>
    [ComVisible(false)]
    [CLSCompliant(true)]
    public class RibbonFactory {
        internal RibbonFactory(IRibbonUI ribbonUI, ResourceManager resourceManager) {
            _ribbonUI  = ribbonUI;
            _controls  = new Dictionary<string,IRibbonCommon>();
            _buttons   = new Dictionary<string,IRibbonButton>();
            _toggles   = new Dictionary<string,IRibbonToggle>();
            _dropDowns = new Dictionary<string,IRibbonDropDown>();

            _resourceManager = resourceManager;
        }

        private readonly IRibbonUI       _ribbonUI;
        private readonly ResourceManager _resourceManager;

        private void PropertyChanged(object sender, IControlChangedEventArgs e) => _ribbonUI.InvalidateControl(e.ControlId);

        private T Add<T>(T ctrl) where T:IRibbonCommon {
            _controls.Add(ctrl.Id, ctrl);
            var button   = ctrl as IRibbonButton;   if (button   != null) _buttons  .Add(ctrl.Id, button);
            var toggle   = ctrl as IRibbonToggle;   if (toggle   != null) _toggles  .Add(ctrl.Id, toggle);
            var dropDown = ctrl as IRibbonDropDown; if (dropDown != null) _dropDowns.Add(ctrl.Id, dropDown);

            ctrl.Changed += PropertyChanged;
            return ctrl;
        }

        /// <summary>TODO</summary>
        public IReadOnlyDictionary<string,IRibbonCommon>     Controls  => new ReadOnlyDictionary<string,IRibbonCommon>(_controls);
        private IDictionary<string,IRibbonCommon>           _controls  { get; }
        /// <summary>TODO</summary>
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification="Matches COM usage.")]
        public IRibbonCommon NewRibbonGroup(
            string              id,
            bool                visible = true,
            bool                enabled = true,
            MyRibbonControlSize size    = Large
        ) => Add(new RibbonGroup(id, _resourceManager, visible, enabled, size));

        /// <summary>TODO</summary>
        public IReadOnlyDictionary<string,IRibbonButton>     Buttons   => new ReadOnlyDictionary<string,IRibbonButton>(_buttons);
        private IDictionary<string,IRibbonButton>           _buttons   { get; }
        /// <summary>TODO</summary>
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification="Matches COM usage.")]
        public IRibbonButton NewRibbonButton(
            string              id,
            bool                visible   = true,
            bool                enabled   = true,
            MyRibbonControlSize size      = Large,
            bool                showImage = false,
            bool                showLabel = true,
            EventHandler        onClickedAction = null
        ) => Add(new RibbonButton(id, _resourceManager, visible, enabled, size, showImage, showLabel, onClickedAction));

        /// <summary>TODO</summary>
        public IReadOnlyDictionary<string,IRibbonToggle>     Toggles   => new ReadOnlyDictionary<string,IRibbonToggle>(_toggles);
        private IDictionary<string,IRibbonToggle>           _toggles   { get; }
        /// <summary>TODO</summary>
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification="Matches COM usage.")]
        public IRibbonToggle NewRibbonToggle(
            string              id,
            bool                visible   = true,
            bool                enabled   = true,
            MyRibbonControlSize size      = Large,
            bool                showImage = false,
            bool                showLabel = true,
            ToggledEventHandler onClickedAction = null
        ) => Add(new RibbonToggle(id, _resourceManager, visible, enabled, size, showImage, showLabel, onClickedAction));

        /// <summary>TODO</summary>
        public IReadOnlyDictionary<string,IRibbonDropDown>   DropDowns => new ReadOnlyDictionary<string,IRibbonDropDown>(_dropDowns);
        private IDictionary<string,IRibbonDropDown>         _dropDowns { get; }
        /// <summary>TODO</summary>
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification="Matches COM usage.")]
        public IRibbonDropDown NewRibbonDropDown(
            string              id,
            bool                visible = true,
            bool                enabled = true,
            MyRibbonControlSize size    = Large
        ) => Add(new RibbonDropDown(id, _resourceManager, visible, enabled, size));

        /// <summary>TODO</summary>
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification="Matches COM usage.")]
        public static LanguageStrings NewLanguageControlRibbonText(
            string label,
            string screenTip      = null,
            string superTip       = null,
            string keyTip         = null,
            string alternateLabel = null,
            string description    = null
        ) => new RibbonTextLanguageControl(label, screenTip, superTip, keyTip, alternateLabel, description);
    }

    /// <summary>TODO</summary>
    public interface IMain {
        /// <summary>TODO</summary>
        IRibbonViewModel NewRibbonViewModel(IRibbonUI ribbonUI, ResourceManager resourceManager);
    }

    /// <summary>TODO</summary>
    public class Main : IMain {
        /// <summary>TODO</summary>
        public Main() { }
        /// <inheritdoc/>
        public IRibbonViewModel NewRibbonViewModel(IRibbonUI ribbonUI, ResourceManager resourceManager)
            => new RibbonViewModel(ribbonUI, resourceManager);
    }

    /// <summary>TODO</summary>
    public interface IRibbonViewModel {

    }

    /// <summary>TODO</summary>
    public class RibbonViewModel : AbstractRibbonDispatcher, IRibbonViewModel {
        /// <summary>TODO</summary>
        public RibbonViewModel(IRibbonUI ribbonUI, ResourceManager resourceManager) : base() {
            InitializeRibbonFactory(ribbonUI, resourceManager);
        }
    }
}
