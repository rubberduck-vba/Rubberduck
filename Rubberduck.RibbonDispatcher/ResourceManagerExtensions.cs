////////////////////////////////////////////////////////////////////////////////////////////////////
//                                Copyright (c) 2017 Pieter Geerkens                              //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.Diagnostics.CodeAnalysis;
using System.Drawing;
using System.Globalization;
using System.Resources;
using System.Windows.Forms;
using stdole;

namespace Rubberduck.RibbonDispatcher {
    /// <summary>TODO</summary>
    public static class ResourceManagerExtensions {
        /// <summary>TODO</summary>
        public static string GetCurrentUIString(this ResourceManager resourceManager, string name)
            => resourceManager?.GetString(name, CultureInfo.CurrentUICulture) ?? "";

        /// <summary>TODO</summary>
        public static string GetInvariantString(this ResourceManager resourceManager, string name)
            => resourceManager?.GetString(name, CultureInfo.InvariantCulture) ?? "";

        /// <summary>TODO</summary>
        public static IPictureDisp GetResourceIcon(this ResourceManager resourceManager, string iconName) {
            using (var icon = resourceManager?.GetObject(iconName, CultureInfo.InvariantCulture) as Icon) {
                return icon == null ? null : PictureConverter.IconToPictureDisp(icon);
            }
        }

        /// <summary>TODO</summary>
        public static IPictureDisp GetResourceImage(this ResourceManager resourceManager, string imageName) {
            using (var image = resourceManager?.GetObject(imageName, CultureInfo.InvariantCulture) as Image) {
                return (image == null) ? null : PictureConverter.ImageToPictureDisp(image);
            }

        }

        [SuppressMessage("Microsoft.Performance", "CA1812:AvoidUninstantiatedInternalClasses")]
        internal class PictureConverter : AxHost {
            private PictureConverter() : base(string.Empty) { }

            public static IPictureDisp ImageToPictureDisp(Image image) => GetIPictureDispFromPicture(image) as IPictureDisp;

            public static IPictureDisp IconToPictureDisp(Icon icon) => ImageToPictureDisp(icon.ToBitmap());
        }
    }
}
