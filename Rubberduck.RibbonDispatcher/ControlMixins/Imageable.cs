using System;
using System.Runtime.CompilerServices;
using stdole;

namespace Rubberduck.RibbonDispatcher.ControlMixins {
    /// <summary>The mixin implementation for IImageable ribbon controls.</summary>
    internal static class Imageable {
        static ConditionalWeakTable<IImageableMixin, Fields> _table = new ConditionalWeakTable<IImageableMixin, Fields>();

        private sealed class Fields {
            public Fields() {}
            public ImageObject  Image     { get; set; } = null;
            public bool         ShowImage { get; set; } = false;
            public bool         ShowLabel { get; set; } = false;
        }

        public static ImageObject GetImage(this IImageableMixin mixin)
            => _table.GetOrCreateValue(mixin).Image;
        public static void SetImage(this IImageableMixin mixin, IPictureDisp image, Action onChanged)
            => mixin.SetImage(new ImageObject(image), onChanged);
        public static void SetImage(this IImageableMixin mixin, string imageMso, Action onChanged)
            => mixin.SetImage(new ImageObject(imageMso), onChanged);
        public static void SetImage(this IImageableMixin mixin, ImageObject image, Action onChanged) {
            _table.GetOrCreateValue(mixin).Image = image;
            onChanged?.Invoke();
        }

        public static bool GetShowImage(this IImageableMixin mixin)
            => _table.GetOrCreateValue(mixin).ShowImage;
        public static void SetShowImage(this IImageableMixin mixin, bool showImage, Action onChanged) {
            _table.GetOrCreateValue(mixin).ShowImage = showImage;
            onChanged?.Invoke();
        }

        public static bool GetShowLabel(this IImageableMixin mixin)
            => _table.GetOrCreateValue(mixin).ShowLabel;
        public static void SetShowLabel(this IImageableMixin mixin, bool showLabel, Action onChanged) {
            _table.GetOrCreateValue(mixin).ShowLabel = showLabel;
            onChanged?.Invoke();
        }
    }
}
