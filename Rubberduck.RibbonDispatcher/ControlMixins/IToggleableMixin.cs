using Rubberduck.RibbonDispatcher.AbstractCOM;
using System;

namespace Rubberduck.RibbonDispatcher.ControlMixins {
    /// <summary>The total interface (required to be) exposed externally by RibbonButton objects.</summary>
    [CLSCompliant(true)]
    internal interface IToggleableMixin {
        void OnActionToggle(bool Pressed);
        /// <summary>TODO</summary>
        IRibbonTextLanguageControl LanguageStrings { get; }
    }
}
