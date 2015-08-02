using System;

namespace Rubberduck.UI.Commands
{
    public class ParentMenuNotFoundException : InvalidOperationException
    {
        private readonly string _caption;
        public string ParentMenuCaption { get { return _caption; } }

        public ParentMenuNotFoundException(string caption)
            : base("Parent menu '" + caption + "' was not found. Cannot create child menu item.")
        {
            _caption = caption;
        }
    }
}