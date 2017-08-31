using Rubberduck.RibbonDispatcher.AbstractCOM;
using System;

namespace Rubberduck.RibbonDispatcher.ControlMixins {
    /// <summary>The interface for controls that can be toggled.</summary>
    [CLSCompliant(true)]
    internal interface IToggleableMixin {
        /// <summary>TODO</summary>
        void OnChanged();

        /// <summary>TODO</summary>
        void OnToggled(bool IsPressed);
            
        /// <summary>TODO</summary>
        IRibbonTextLanguageControl LanguageStrings { get; }
    }
}
