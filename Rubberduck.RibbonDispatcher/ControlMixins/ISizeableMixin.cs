using System;
using System.Diagnostics.CodeAnalysis;

namespace Rubberduck.RibbonDispatcher.ControlMixins {
    /// <summary>The interface for controls that can be sized.</summary>
    [SuppressMessage("Microsoft.Design", "CA1040:AvoidEmptyInterfaces", Justification="False positive for Mixins.")]
    [CLSCompliant(true)]
    public interface ISizeableMixin {
    }
}
