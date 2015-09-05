using System;
using System.Drawing;
using System.Windows.Input;

namespace Rubberduck.UI.Command
{
    public interface IMenuItem
    {
        string Key { get; }
        Func<string> Caption { get; }
        bool BeginGroup { get; }
        int DisplayOrder { get; }
    }

    public interface ICommandMenuItem : IMenuItem
    {
        ICommand Command { get; }
        Image Image { get; }
        Image Mask { get; }
    }
}
