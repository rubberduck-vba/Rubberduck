namespace Rubberduck.UI
{
    public interface IDockableUserControl
    {
        /// <summary>
        /// Gets a string containing some semi-random GUID to use for positional registration
        /// </summary>
        string GuidIdentifier { get; }

        /// <summary>
        /// Gets a string containing the caption of the toolwindow.
        /// </summary>
        string Caption { get; }
    }
}
