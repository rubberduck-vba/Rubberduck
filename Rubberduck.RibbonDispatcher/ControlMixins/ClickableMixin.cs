using System;
using System.Runtime.InteropServices;

namespace Rubberduck.RibbonDispatcher.ControlMixins {
    /// <summary>TODO</summary>
    [CLSCompliant(true)]
    public delegate void ClickedEventHandler();

    internal static class ClickableMixin {
        public static void OnAction(this IClickableMixin mixin, Action clicked) => clicked?.Invoke();
    }
}
