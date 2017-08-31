using System.Runtime.CompilerServices;
using stdole;

namespace Rubberduck.RibbonDispatcher.ControlMixins {
    /// <summary>The mixin implementation for IImageable ribbon controls.</summary>
    internal static class ImageableMixin {
        static ConditionalWeakTable<IImageableMixin, Fields> _table = new ConditionalWeakTable<IImageableMixin, Fields>();

        private sealed class Fields {
            public Fields() {}
            public ImageObject  Image     { get; set; } = null;
            public bool         ShowImage { get; set; } = false;
            public bool         ShowLabel { get; set; } = false;
        }
        private static Fields Mixin(this IImageableMixin imageable) => _table.GetOrCreateValue(imageable);

        public static ImageObject GetImage(this IImageableMixin mixin)
            => _table.GetOrCreateValue(mixin).Image;
        public static void SetImage(this IImageableMixin imageable, IPictureDisp image)
            => imageable.SetImage(new ImageObject(image));
        public static void SetImage(this IImageableMixin imageable, string imageMso)
            => imageable.SetImage(new ImageObject(imageMso));
        public static void SetImage(this IImageableMixin imageable, ImageObject image) {
            imageable.Mixin().Image = image;
            imageable.OnChanged();
        }

        public static bool GetShowImage(this IImageableMixin imageable)
            => _table.GetOrCreateValue(imageable).ShowImage;
        public static void SetShowImage(this IImageableMixin imageable, bool showImage) {
            imageable.Mixin().ShowImage = showImage;
            imageable.OnChanged();
        }

        public static bool GetShowLabel(this IImageableMixin imageable)
            => imageable.Mixin().ShowLabel;
        public static void SetShowLabel(this IImageableMixin imageable, bool showLabel) {
            imageable.Mixin().ShowLabel = showLabel;
            imageable.OnChanged();
        }
    }
}
