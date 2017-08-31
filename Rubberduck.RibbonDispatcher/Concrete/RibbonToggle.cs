using System;
using System.Resources;
using System.Runtime.InteropServices;

using Rubberduck.RibbonDispatcher.Abstract;
using Rubberduck.RibbonDispatcher.EventHandlers;

namespace Rubberduck.RibbonDispatcher.Concrete {
    /// <summary>TODO</summary>
    [Serializable]
    [CLSCompliant(true)]
    [ClassInterface(ClassInterfaceType.None)]
    [ComSourceInterfaces(typeof(IToggledEvents))]
    public class RibbonToggle : RibbonCommon, IRibbonToggle {
        internal RibbonToggle(string id, ResourceManager mgr, bool visible, bool enabled, MyRibbonControlSize size,
                bool showImage, bool showLabel, ToggledEventHandler onToggledAction)
            : base(id, mgr, visible, enabled, size) {
            _showImage = showImage;
            _showLabel = showLabel;
            if (onToggledAction != null) Toggled += onToggledAction;
        }

        /// <summary>TODO</summary>
        public event ToggledEventHandler Toggled;

        /// <summary>TODO</summary>
        public new string Label       => IsPressed && ! String.IsNullOrEmpty(LanguageStrings?.AlternateLabel ?? "")
                                       ? LanguageStrings?.AlternateLabel??Id 
                                       : LanguageStrings?.Label??Id;
        /// <summary>TODO</summary>
        public bool IsPressed         { get; private set; }

        /// <summary>TODO</summary>
        public bool ShowLabel {
            get { return _showLabel; }
            set { _showLabel = value; OnChanged(); }
        }
        private bool _showLabel;
        /// <inheritdoc/>
        public bool ShowImage {
            get { return _showImage; }
            set { _showImage = value; OnChanged(); }
        }
        private bool _showImage;

        /// <summary>TODO</summary>
        public void OnAction(bool isPressed) {
            Toggled?.Invoke(this,new ToggledEventArgs(isPressed));
            IsPressed         = isPressed;
            OnChanged();
        }
   }

}
