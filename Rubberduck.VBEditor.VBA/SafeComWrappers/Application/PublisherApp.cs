using System;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

// ReSharper disable once CheckNamespace - Special dispensation due to conflicting file vs namespace priorities
namespace Rubberduck.VBEditor.SafeComWrappers.VBA
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
