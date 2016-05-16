﻿using System;
using Microsoft.Office.Interop.Publisher;

namespace Rubberduck.VBEditor.VBEHost
{
    public class PublisherApp : HostApplicationBase<Application>
    {
        public PublisherApp() : base("Publisher") { }

        public override void Run(QualifiedMemberName qualifiedMemberName)
        {
            //Publisher does not support the Run method
            throw new NotImplementedException("Unit Testing not supported for Publisher");
        }
    }
}