namespace Rubberduck.Common
{
    public interface ILowLevelKeyboardHook : IHook
    {
        bool EatNextKey { get; set; }
    }
}