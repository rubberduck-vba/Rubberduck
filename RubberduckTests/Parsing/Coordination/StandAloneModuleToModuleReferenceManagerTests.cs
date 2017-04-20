using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Linq;
using Rubberduck.Parsing.VBA;
using RubberduckTests.Mocks;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
namespace RubberduckTests.Parsing.Coordination
{
    [TestClass]
    public class StandAloneModuleToModuleReferenceManagerTests : IModuleToModuleReferenceManagerTestBase
    {
        protected override IModuleToModuleReferenceManager GetNewTestModuleToModuleReferenceManager()
        {
            return new ModuleToModuleReferenceManager();
        }
    }
}
