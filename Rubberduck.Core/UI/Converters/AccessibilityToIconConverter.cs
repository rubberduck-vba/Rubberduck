using System;
using System.Globalization;
using System.Windows.Media;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Resources.CodeExplorer;

namespace Rubberduck.UI.Converters
{
    public class AccessibilityToIconConverter : ImageSourceConverter
    {
        private static readonly ImageSource PrivateOverlay = ToImageSource(CodeExplorerUI.AccessibilityPrivate);
        private static readonly ImageSource FriendOverlay = ToImageSource(CodeExplorerUI.AccessibilityFriend);
        private static readonly ImageSource GlobalOverlay = ToImageSource(CodeExplorerUI.AccessibilityGlobal);
        private static readonly ImageSource StaticOverlay = ToImageSource(CodeExplorerUI.AccessibilityStatic);

        public override object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (!(value is Declaration declaration))
            {
                return null;
            }

            switch (declaration.Accessibility)
            {
                case Accessibility.Private:
                    return PrivateOverlay;
                case Accessibility.Friend:
                    return FriendOverlay;
                case Accessibility.Global:
                    return GlobalOverlay;
                case Accessibility.Static:
                    return StaticOverlay;
                default:
                    return null;
            }
        }
    }
}
