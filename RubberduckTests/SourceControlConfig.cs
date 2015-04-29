using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Collections.Generic;
using Rubberduck.SourceControl;
using Rubberduck.Config;

namespace RubberduckTests
{
    // Tests disabled because they aren't meant to be unit tests.
    // These hit the file system and are for ease of debugging.
    [Ignore]
    [TestClass]
    public class SourceControlConfig
    {
        [TestMethod]
        public void Save()
        {
            var repo = new Rubberduck.SourceControl.Repository
            (
                "SourceControlTest",
                @"C:\Users\Christopher\Documents\SourceControlTest",
                @"https://github.com/ckuhn203/SourceControlTest.git"
            );


            var config = new SourceControlConfiguration();
            config.Repositories = new List<Repository>() { repo };

            var service = new SourceControlConfigurationService();
            service.SaveConfiguration(config);

        }

        [TestMethod]
        public void Load()
        {
            var service = new SourceControlConfigurationService();
            var config = service.LoadConfiguration();
        }
    }
}
