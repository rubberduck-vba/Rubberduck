using System;
using System.Drawing;

namespace Rubberduck.VBEditor.SafeComWrappers.Abstract
{
    public interface ICommandBarButton : ISafeEventedComWrapper, ICommandBarControl
    {
        event EventHandler<CommandBarButtonClickEventArgs> Click;
        bool IsBuiltInFace { get; set; }
        int FaceId { get; set; }
        string ShortcutText { get; set; }
        ButtonState State { get; set; }
        ButtonStyle Style { get; set; }
        Image Picture { get; set; }
        Image Mask { get; set; }
        void ApplyIcon();
    }
}