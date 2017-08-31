using Rubberduck.RibbonDispatcher.AbstractCOM;
using System;
using System.Runtime.CompilerServices;

namespace Rubberduck.RibbonDispatcher.ControlDecorators {
    /// <summary>The mixin implementation for Sizeable ribbon controls.</summary>
    public static class Sizeable {
        static ConditionalWeakTable<ISizeableMixin,Fields> _table = new ConditionalWeakTable<ISizeableMixin, Fields>();

        private sealed class Fields {
            public Fields() { ; }
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
