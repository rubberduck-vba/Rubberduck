using System;
using System.IO;

namespace Rubberduck.Settings
{
    public class SourceControlConfigurationService : XmlConfigurationServiceBase<SourceControlConfiguration>
    {

        protected override string ConfigFile
        {
            get { return Path.Combine(rootPath, "SourceControl.rubberduck"); }
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
