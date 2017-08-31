﻿using System;
using System.Runtime.CompilerServices;

namespace Rubberduck.RibbonDispatcher.ControlMixins {
    /// <summary>The mixin implementation for IToggleable ribbon controls.</summary>
    internal static class Toggleable {
        static ConditionalWeakTable<IToggleableMixin, Fields> _table = new ConditionalWeakTable<IToggleableMixin, Fields>();

        private sealed class Fields {
            public Fields() {}
            public bool         IsPressed { get; set; } = false;
        }

        public   static bool   GetPressed(this IToggleableMixin mixin)     => _table.GetOrCreateValue(mixin).IsPressed;
        public   static string GetLabel(this IToggleableMixin mixin)       => mixin.GetLabel(_table.GetOrCreateValue(mixin));
        internal static string AlternateLabel(this IToggleableMixin mixin) => mixin.LanguageStrings.AlternateLabel;
        internal static string Label(this IToggleableMixin mixin)          => mixin.LanguageStrings.Label;

        private  static string GetLabel(this IToggleableMixin mixin, Fields fields)
            => fields.IsPressed && ! string.IsNullOrEmpty(mixin.AlternateLabel()) ? mixin.AlternateLabel()
                                                                                  : mixin.Label();

        public static void OnActionToggled(this IToggleableMixin mixin, bool isPressed, Action<bool> toggled) {
            _table.GetOrCreateValue(mixin).IsPressed = isPressed;
            toggled?.Invoke(isPressed);
        }
    }
}
