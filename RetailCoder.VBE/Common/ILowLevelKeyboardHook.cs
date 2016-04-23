namespace Rubberduck.Common
{
    public interface ILowLevelKeyboardHook : IAttachable
    {
        bool EatNextKey { get; set; }
    }
}