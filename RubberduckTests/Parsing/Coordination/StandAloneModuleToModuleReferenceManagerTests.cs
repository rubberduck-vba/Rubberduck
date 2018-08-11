using System;
using NUnit.Framework;
using System.Linq;
using Rubberduck.Parsing.VBA;
using Rubberduck.Parsing.VBA.ReferenceManagement;
using RubberduckTests.Mocks;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
namespace RubberduckTests.Parsing.Coordination
{
    [TestFixture]
    public class StandAloneModuleToModuleReferenceManagerTests : IModuleToModuleReferenceManagerTestBase
    {
        protected override IModuleToModuleReferenceManager GetNewTestModuleToModuleReferenceManager()
        {
            return new ModuleToModuleReferenceManager();
        }
    }
}
