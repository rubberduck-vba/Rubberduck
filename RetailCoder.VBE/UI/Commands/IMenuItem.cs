using System;
using System.Drawing;

namespace Rubberduck.UI.Commands
{
    public interface IMenuItem
    {
        Func<string> Caption { get; }

        string Key { get; }

        bool IsParent { get; }
        Image Image { get; }
        Image Mask { get; }
    }
}