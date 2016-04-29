using System.Windows.Forms;

namespace Rubberduck.Common.WinAPI
{
    public class PreMessageFilter : IMessageFilter
    {
        // true  to filter the message and stop it from being dispatched 
        // false to allow the message to continue to the next filter or control.
        public bool PreFilterMessage(ref Message m)
        {
            return false;
        }
    }
}
