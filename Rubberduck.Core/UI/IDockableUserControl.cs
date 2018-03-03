namespace Rubberduck.UI
{
    public interface IDockableUserControl
    {
        /// <summary>
        /// Gets a string containing the GUID with which the class is registered.
        /// </summary>
        string ClassId { get; }

        /// <summary>
        /// Gets a string containing the caption of the toolwindow.
        /// </summary>
        string Caption { get; }
    }
}
