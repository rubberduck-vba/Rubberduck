﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Rubberduck.SourceControl;
using NetOffice.VBIDEApi;

namespace Rubberduck.UI.SourceControl
{
    public interface ISourceControlProviderFactory
    {
        ISourceControlProvider CreateProvider(VBProject project);
        ISourceControlProvider CreateProvider(VBProject project, IRepository repository);
    }

    public class SourceControlProviderFactory : ISourceControlProviderFactory
    {
        public ISourceControlProvider CreateProvider(VBProject project)
        {
            return new GitProvider(project);
        }

        public ISourceControlProvider CreateProvider(VBProject project, IRepository repository)
        {
            try
            {
                return new GitProvider(project, repository);
            }
            catch (SourceControlException e)
            {
                throw;
            }
        }
    }
}
