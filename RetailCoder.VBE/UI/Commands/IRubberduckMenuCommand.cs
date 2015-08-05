using System;
using System.Drawing;
using Microsoft.Office.Core;

namespace Rubberduck.UI.Commands
{
    /// <summary>
    /// An object that encapsulates the logic to wire up a number of CommandBarControl instances to a specific command.
    /// </summary>
    public interface IRubberduckMenuCommand
    {
        /// <summary>
        /// Associates a new <see cref="CommandBarButton"/> to the command.
        /// </summary>
        /// <param name="parent">The parent control collection to add the button to.</param>
        /// <param name="caption">The localized caption for the command.</param>
        /// <param name="beginGroup">Optionally specifies that the UI element begins a command group.</param>
        /// <param name="beforeIndex">Optionally specifies the index of the UI element in the parent collection.</param>
        /// <param name="image">An optional icon to represent the command.</param>
        /// <param name="mask">A transparency mask for the command's icon. Required if <see cref="image"/> is not null.</param>
        void AddCommandBarButton(CommandBarControls parent, string caption, bool beginGroup = false, int? beforeIndex = null, Image image = null, Image mask = null);
        
        /// <summary>
        /// Destroys all UI elements associated to the command.
        /// </summary>
        void Release();

        /// <summary>
        /// Requests execution of the command.
        /// </summary>
        event EventHandler RequestExecute;
    }
}