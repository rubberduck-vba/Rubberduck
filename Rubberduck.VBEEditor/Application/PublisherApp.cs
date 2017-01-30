using System;

namespace Rubberduck.VBEditor.Application
{
    public class PublisherApp : HostApplicationBase<Microsoft.Office.Interop.Publisher.Application>
    {
        public PublisherApp() : base("Publisher") { }

        public override void Run(dynamic declaration)
        {
            //Publisher does not support the Run method
            throw new NotImplementedException("Unit Testing not supported for Publisher");
        }
    }
}
