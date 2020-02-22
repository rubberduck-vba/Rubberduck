using NUnit.Framework;
using System.Linq;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor;

namespace RubberduckTests.Symbols
{
    [TestFixture]
    public class ProjectDeclarationTests
    {
        [Test]
        [Category("Resolver")]
        public void ProjectsHaveDeclarationTypeProject()
        {
            var projectDeclaration = GetTestProject("testProject");

            Assert.IsTrue(projectDeclaration.DeclarationType.HasFlag(DeclarationType.Project));
        }

            private static ProjectDeclaration GetTestProject(string name)
            {
                var qualifiedProjectName = new QualifiedMemberName(StubQualifiedModuleName(), name);
                return new ProjectDeclaration(qualifiedProjectName, name, true);
            }

                private static QualifiedModuleName StubQualifiedModuleName()
                {
                    return new QualifiedModuleName("dummy", "dummy", "dummy");
                }


        [Test]
        [Category("Resolver")]
        public void ByDefaultProjectsReferenceNoOtherProjects()
        {
            var projectDeclaration = GetTestProject("testProject");

            Assert.IsFalse(projectDeclaration.ProjectReferences.Any());
        }


        [Test]
        [Category("Resolver")]
        public void ProjectsReferencesReturnsTheReferencesAddedViaAddProjectReference()
        {
            var projectDeclaration = GetTestProject("testProject");
            var projectId = "test";
            var priority = 12;
            projectDeclaration.AddProjectReference(projectId, priority);
            var projectReference = projectDeclaration.ProjectReferences.Single();

            Assert.IsTrue(projectReference.ReferencedProjectId == projectId && projectReference.Priority == priority);
        }


        [Test]
        [Category("Resolver")]
        public void ClearProjectsReferencesClearsTheProjectReferences()
        {
            var projectDeclaration = GetTestProject("testProject");
            var projectId = "test";
            var priority = 12;
            projectDeclaration.AddProjectReference(projectId, priority);
            projectDeclaration.ClearProjectReferences();
            var projectReferences = projectDeclaration.ProjectReferences;

            Assert.IsFalse(projectReferences.Any());
        }


        [Test]
        [Category("Resolver")]
        public void ProjectsReferencesIgnoresReferencesWithTheSameIDAsOneAlreadyPresent()
        {
            var projectDeclaration = GetTestProject("testProject");
            var projectId = "test";
            var priority = 12;
            var otherPriority = 1;
            projectDeclaration.AddProjectReference(projectId, priority);
            projectDeclaration.AddProjectReference(projectId, otherPriority);
            var projectReference = projectDeclaration.ProjectReferences.Single();

            Assert.IsTrue(projectReference.ReferencedProjectId == projectId && projectReference.Priority == priority);
        }


        [Test]
        [Category("Resolver")]
        public void ProjectsReferencesReturnsTheReferencesInOrderOfAscendingPriority()
        {
            var projectDeclaration = GetTestProject("testProject");
            var projectId = "test";
            var priority = 12;
            var otherProjectId = "testtest";
            var otherPriority = 1;
            var yetAnotherProjectId = "testtesttest";
            var yetAnotherPriority = 5;
            projectDeclaration.AddProjectReference(projectId, priority);
            projectDeclaration.AddProjectReference(otherProjectId, otherPriority);
            projectDeclaration.AddProjectReference(yetAnotherProjectId, yetAnotherPriority);
            var lowerPriorityProjectReference = projectDeclaration.ProjectReferences.First();

            Assert.IsTrue(lowerPriorityProjectReference.ReferencedProjectId == otherProjectId && lowerPriorityProjectReference.Priority == otherPriority);
        }

    }
}
