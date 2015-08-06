using System;
using System.Drawing;
using Microsoft.Office.Core;

namespace Rubberduck.UI.Command
{
    public interface IMenuItem
    {
        string Key { get; }
        Func<string> Caption { get; }
        bool BeginGroup { get; }
        int DisplayOrder { get; }
    }

    public interface IParentMenuItem : IMenuItem
    {
        CommandBarPopup Item { get; }
        void Localize();
        void Initialize();
    }

    public interface ICommandMenuItem : IMenuItem
    {
        ICommand Command { get; }
        Image Image { get; }
        Image Mask { get; }
    }
}
