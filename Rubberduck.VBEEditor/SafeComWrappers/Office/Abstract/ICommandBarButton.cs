using System;
using System.Drawing;
using Rubberduck.VBEditor.SafeComWrappers.VB.Enums;

namespace Rubberduck.VBEditor.SafeComWrappers.Office.Abstract
{
    public interface ICommandBarButton : ICommandBarControl
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