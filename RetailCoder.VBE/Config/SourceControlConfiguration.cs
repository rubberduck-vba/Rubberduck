using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Rubberduck.SourceControl;
using System.IO;

namespace Rubberduck.Config
{
    public class SourceControlConfiguration
    {
        public List<Repository> Repositories;
    }

    public class SourceControlConfigurationService : XmlConfigurationServiceBase<SourceControlConfiguration>, IConfigurationService<SourceControlConfiguration>
    {

        protected override string ConfigFile
        {
            get { return Path.Combine(this.rootPath, "SourceControl.rubberduck"); }
        }

        public override SourceControlConfiguration LoadConfiguration()
        {
            return base.LoadConfiguration();
        }

        protected override SourceControlConfiguration HandleIOException(System.IO.IOException ex)
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
