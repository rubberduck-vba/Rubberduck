namespace Rubberduck.Common.WinAPI
{
    public enum DataCommand : uint
    {
        RID_HEADER = 0x10000005, // Get the header information from the RAWINPUT structure.
        RID_INPUT = 0x10000003   // Get the raw data from the RAWINPUT structure.
    }
}
