using System;

namespace Rubberduck.RibbonDispatcher.ControlMixins {
    /// <summary>TODO</summary>
    [CLSCompliant(true)]
    public delegate void ClickedEventHandler();

    internal static class ClickableMixin {
        public static void Clicked(this IClickableMixin mixin) => mixin.OnClicked();
    }
}
