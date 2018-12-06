using System.Runtime.InteropServices;

namespace Rubberduck.Common.WinAPI
{
    /// <summary>
    /// Native functions from SHCore.dll
    /// </summary>
    public static class SHCore
    {
        /// <summary>
        /// Sets the DPI awareness level of the current process.
        /// </summary>
        /// <param name="awareness">DPI awareness level.</param>
        /// <returns>HRESULT of S_OK, E_INVALIDARG or E_ACCESSDENIED.</returns>
        /// <remarks>
        /// Only the first DPI awareness call made by a process will have effect, subsequent calls are disregarded.
        /// Thus, calling this method before WPF loads will override the default WPF DPI awareness behavior.
        /// </remarks>
        [DllImport("SHCore.dll", SetLastError = true)]
        public static extern bool SetProcessDpiAwareness(PROCESS_DPI_AWARENESS awareness);
    }

    /// <summary>
    /// Describes DPI awareness of a process.
    /// </summary>
    public enum PROCESS_DPI_AWARENESS
    {
        /// <summary>
        /// Process is not DPI aware. 
        /// </summary>
        Process_DPI_Unaware = 0,

        /// <summary>
        /// Process is aware of the System DPI (monitor 1).
        /// </summary>
        Process_System_DPI_Aware = 1,

        /// <summary>
        /// Process is aware of the DPI of individual monitors.
        /// </summary>
        Process_Per_Monitor_DPI_Aware = 2
    }

}
