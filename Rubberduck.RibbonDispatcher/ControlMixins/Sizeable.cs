using System;
using System.Runtime.CompilerServices;

using Rubberduck.RibbonDispatcher.AbstractCOM;

namespace Rubberduck.RibbonDispatcher.ControlMixins {
    /// <summary>The mixin implementation for ISizeable ribbon controls.</summary>
    internal static class Sizeable {
        static ConditionalWeakTable<ISizeableMixin,Fields> _table = new ConditionalWeakTable<ISizeableMixin, Fields>();

        private sealed class Fields {
            public Fields() {}
            public RdControlSize ControlSize = RdControlSize.rdLarge;
        }

        /// <summary>Sets the {RdControlSize} value for an {ISizeableMixin} mixin.</summary>
        public static RdControlSize GetSize(this ISizeableMixin sizeable)
            => _table.GetOrCreateValue(sizeable).ControlSize;

        /// <summary>Sets the {RdControlSize} value for an {ISizeableMixin} mixin.</summary>
        public static void SetSize(this ISizeableMixin sizeable, RdControlSize size, Action onChanged) {
            _table.GetOrCreateValue(sizeable).ControlSize = size; onChanged?.Invoke();
        }
    }
}
