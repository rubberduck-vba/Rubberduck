using System.Windows.Forms;

namespace Rubberduck.Common
{
    public interface IHotKey : IAttachable
    {
        HotKeyInfo HotKeyInfo { get; }
        Keys Combo { get; }
        Keys SecondKey { get; }
        bool IsTwoStepHotKey { get; }
    }
}