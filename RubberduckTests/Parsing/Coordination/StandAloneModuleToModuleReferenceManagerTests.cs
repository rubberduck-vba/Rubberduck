using Microsoft.VisualStudio.TestTools.UnitTesting;
using Rubberduck.Parsing.VBA;

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
