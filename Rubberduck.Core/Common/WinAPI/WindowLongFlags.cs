namespace Rubberduck.Common.WinAPI
{
    public enum WindowLongFlags : int
    {
        /// <summary>
        /// Sets a new extended window style.
        /// </summary>
        GWL_EXSTYLE = -20,
        /// <summary>
        /// Sets a new application instance handle.
        /// </summary>
        GWLP_HINSTANCE = -6,
        GWLP_HWNDPARENT = -8,
        /// <summary>
        /// Sets a new identifier of the child window. The window cannot be a top-level window.
        /// </summary>
        GWL_ID = -12,
        /// <summary>
        /// Sets a new window style.
        /// </summary>
        GWL_STYLE = -16,
        /// <summary>
        /// Sets the user data associated with the window. This data is intended for use by the application that created the window. Its value is initially zero.
        /// </summary>
        GWL_USERDATA = -21,
        /// <summary>
        /// Sets a new address for the window procedure. You cannot change this attribute if the window does not belong to the same process as the calling thread.
        /// </summary>
        GWL_WNDPROC = -4,
        /// <summary>
        /// Sets new extra information that is private to the application, such as handles or pointers.
        /// </summary>
        DWLP_USER = 0x8,
        /// <summary>
        /// Sets the return value of a message processed in the dialog box procedure.
        /// </summary>
        DWLP_MSGRESULT = 0x0,
        /// <summary>
        /// Sets the new address of the dialog box procedure.
        /// </summary>
        DWLP_DLGPROC = 0x4
    }
}
