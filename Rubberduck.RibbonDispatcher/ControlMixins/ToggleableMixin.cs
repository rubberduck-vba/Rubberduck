using System;
using System.Runtime.CompilerServices;

namespace Rubberduck.RibbonDispatcher.ControlMixins {
    /// <summary>TODO</summary>
    [CLSCompliant(true)]
    public delegate void ToggledEventHandler(bool IsPressed);

    /// <summary>The mixin implementation for IToggleable ribbon controls.</summary>
    internal static class ToggleableMixin {
        static ConditionalWeakTable<IToggleableMixin, Fields> _table = new ConditionalWeakTable<IToggleableMixin, Fields>();

        private sealed class Fields {
            public Fields() {}
            public bool IsPressed { get; set; } = false;
        }
        private static Fields Mixin(this IToggleableMixin mixin) => _table.GetOrCreateValue(mixin);

        public static void OnActionToggle(this IToggleableMixin mixin, bool isPressed) {
            mixin.Mixin().IsPressed = isPressed;
            mixin.OnToggled(isPressed);
            mixin.OnChanged();
        }

        public  static bool   GetPressed(this IToggleableMixin mixin)     => mixin.Mixin().IsPressed;
        public  static string GetLabel(this IToggleableMixin mixin)       => mixin.GetLabel(mixin.Mixin());
        private static string AlternateLabel(this IToggleableMixin mixin) => mixin.LanguageStrings.AlternateLabel;
        private static string Label(this IToggleableMixin mixin)          => mixin.LanguageStrings.Label;

        private  static string GetLabel(this IToggleableMixin mixin, Fields fields)
            => fields.IsPressed && ! string.IsNullOrEmpty(mixin.AlternateLabel()) ? mixin.AlternateLabel()
                                                                                  : mixin.Label();
    }
}
