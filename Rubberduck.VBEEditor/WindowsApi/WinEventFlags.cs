using System;

namespace Rubberduck.VBEditor.WindowsApi
{
    [Flags]
    public enum WinEventFlags
    {
        //Asynchronous events.
        OutOfContext = 0x0000,
        //No events raised from caller thread. Must be combined with OutOfContext or InContext.
        SkipOwnThread = 0x0001,
        //No events raised from caller process.  Must be combined with OutOfContext or InContext.
        SkipOwnProcess = 0x0002,
        //Synchronous events - injects into *all* processes.
        InContext = 0x0004
    }
}
