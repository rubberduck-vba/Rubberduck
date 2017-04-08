using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Linq;
using Rubberduck.Parsing.VBA;
using RubberduckTests.Mocks;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using Moq;
using System.Threading;

namespace RubberduckTests.Parsing.Coordination
{
    [TestClass]
    public class ModuleToModuleReferenceManagerTests : IModuleToModuleReferenceManagerTestBase
    {
        protected override IModuleToModuleReferenceManager GetNewTestModuleToModuleReferenceManager()
        {
            var modules = ModulesUsedInTheBaseTests;
            var projectId = modules.Select(qmn => qmn.ProjectId).First();
            var projectName = modules.Select(qmn => qmn.ProjectName).First();
            var projectPath = modules.Select(qmn => qmn.ProjectPath).First();
            var vbeBuilder = new MockVbeBuilder();
            var projectBuilder = vbeBuilder.ProjectBuilder(projectName, projectPath, projectId, Rubberduck.VBEditor.SafeComWrappers.ProjectProtection.Unprotected);
            foreach (var module in modules.Where(qmn => String.Equals(qmn.ProjectId, projectId)))
            {
                projectBuilder.AddComponent(module.ComponentName, Rubberduck.VBEditor.SafeComWrappers.ComponentType.StandardModule, String.Empty);
            }
            var project = projectBuilder.Build();
            vbeBuilder.AddProject(project);
            var vbe = vbeBuilder.Build();
            var state = Parse(vbe);
            return new ModuleToModuleReferenceManager(state);
        }

        private static RubberduckParserState Parse(Mock<IVBE> vbe)
        {
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState(vbe.Object));
            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status != ParserState.Ready)
            {
                Assert.Inconclusive("Parser state should be 'Ready', but returns '{0}'.", parser.State.Status);
            }
            var state = parser.State;
            return state;
        }
    }
}
