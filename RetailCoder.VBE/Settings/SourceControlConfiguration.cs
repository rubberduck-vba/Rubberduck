using System;
using System.Collections.Generic;
using System.IO;
using Rubberduck.SourceControl;

namespace Rubberduck.Settings
{
    public class SourceControlConfiguration
    {
        public List<Repository> Repositories;

        public SourceControlConfiguration()
        {
            this.Repositories = new List<Repository>();
        }
    }

    public class SourceControlConfigurationService : XmlConfigurationServiceBase<SourceControlConfiguration>
    {

        protected override string ConfigFile
        {
            get { return Path.Combine(this.rootPath, "SourceControl.rubberduck"); }
        }

        public override SourceControlConfiguration LoadConfiguration()
        {
            return base.LoadConfiguration();
        }

        protected override SourceControlConfiguration HandleIOException(IOException ex)
        {
            //couldn't load file
            return new SourceControlConfiguration();
        }

        protected override SourceControlConfiguration HandleInvalidOperationException(InvalidOperationException ex)
        {
            //couldn't load file
            return new SourceControlConfiguration();
        }
    }


}
