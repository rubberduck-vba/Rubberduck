using System;
using System.Collections.Generic;
using System.Linq;
using Moq;
using NUnit.Framework;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.ComManagement;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using RubberduckTests.Mocks;

namespace RubberduckTests.VBEditor
{
    [TestFixture()]
    public class ProjectsRepositoryTests
    {
        private IProjectsRepository TestRepository(IVBE vbe, bool initialRefresh = true)
        {
            var repository = new ProjectsRepository(vbe);
            if (initialRefresh)
            {
                repository.Refresh();
            }
            return repository;
        }

        private Mock<IVBProject> TestProject(
            string projectName, 
            ProjectProtection protection,
            MockVbeBuilder vbeBuilder,
            params (string componentName, ComponentType componentType, string contents)[] componentSpecs)
        {
            return TestProjectAndComponents(projectName, protection, vbeBuilder, componentSpecs).projectMock;
        }
        private (Mock<IVBProject> projectMock, Mock<IVBComponents> componentsMock, List<Mock<IVBComponent>> componentMocks) TestProjectAndComponents(
            string projectName, 
            ProjectProtection protection, 
            MockVbeBuilder vbeBuilder,
            params (string componentName, ComponentType componentType, string contents)[] componentSpecs)
        {
            var projectBuilder = protection == ProjectProtection.Locked 
                ? vbeBuilder.ProjectBuilder(projectName, "projectPath", protection) 
                : vbeBuilder.ProjectBuilder(projectName, protection);

            foreach (var (componentName, componentType, contents) in componentSpecs)
            {
                projectBuilder.AddComponent(componentName, componentType, contents);
            }

            var projectMock = projectBuilder.Build();

            if (protection == ProjectProtection.Locked)
            {
                var projectId = QualifiedModuleName.GetProjectId(projectName, "projectPath");
                projectMock.Setup(m => m.ProjectId).Returns(projectId);
            }

            var componentsMock = projectBuilder.MockVBComponents;
            var componentMocks = projectBuilder.MockComponents;

            return (projectMock, componentsMock, componentMocks);
        }

        [Test]
        [Category("COM")]
        public void ProjectsCollectionReturnsTheOneFromTheVbePassedIn()
        {
            var vbeBuilder = new MockVbeBuilder();
            var vbe = vbeBuilder.Build().Object;
            var repository = TestRepository(vbe);

            var vbeProjectsMock = vbeBuilder.MockProjectsCollection;
            var repositoryProjects = repository.ProjectsCollection();
            
            //We use that the actions on the returned decorators are observable on the original mock to identify the inner projects collection.
            var _ = repositoryProjects.Count;
            vbeProjectsMock.Verify(m => m.Count, Times.Once);
        }

        [Test]
        [Category("COM")]
        public void ProjectsCollectionGetsDisposedOnDisposal()
        {
            var vbeBuilder = new MockVbeBuilder();
            var vbe = vbeBuilder.Build().Object;
            var repository = TestRepository(vbe);

            repository.Dispose();

            var projectsMock = vbeBuilder.MockProjectsCollection;
            projectsMock.Verify(m => m.Dispose(), Times.Once);
        }

        [Test]
        [Category("COM")]
        public void ProjectsCollectionDoesNotGetDisposedOnDisposalOfReturnedValue()
        {
            var vbeBuilder = new MockVbeBuilder();
            var vbe = vbeBuilder.Build().Object;
            var repository = TestRepository(vbe);
            var repositoryProjects = repository.ProjectsCollection();

            repositoryProjects.Dispose();

            var projectsMock = vbeBuilder.MockProjectsCollection;
            projectsMock.Verify(m => m.Dispose(), Times.Never);
        }

        [Test]
        [Category("COM")]
        public void ProjectsCollectionReturnsNullAfterDisposal()
        {
            var vbe = new MockVbeBuilder().Build().Object;
            var repository = TestRepository(vbe);

            repository.Dispose();

            Assert.IsNull(repository.ProjectsCollection());
        }

        [Test]
        [Category("COM")]
        public void ProjectsReturnsProjectsOnVbe()
        {
            var vbeBuilder = new MockVbeBuilder();
            var projectMock = TestProject("project", ProjectProtection.Unprotected, vbeBuilder);
            vbeBuilder.AddProject(projectMock);
            var otherProjectMock = TestProject("otherProject", ProjectProtection.Unprotected, vbeBuilder);
            vbeBuilder.AddProject(otherProjectMock);

            var vbe = vbeBuilder.Build().Object;
            var repository = TestRepository(vbe);
            var projects = repository.Projects().ToList();

            Assert.AreEqual(2, projects.Count);

            foreach (var (projectId, repositoryProject) in projects)
            {
                Assert.AreEqual(projectId, repositoryProject.ProjectId);
                var _ = repositoryProject.VBE;
            }

            projectMock.Verify(m => m.VBE, Times.Once);
            otherProjectMock.Verify(m => m.VBE, Times.Once);
        }

        [Test]
        [Category("COM")]
        public void ProjectsDoesNotReturnsLockedProjectsOnVbe()
        {
            var vbeBuilder = new MockVbeBuilder();
            var projectMock = TestProject("project", ProjectProtection.Locked, vbeBuilder);
            vbeBuilder.AddProject(projectMock);
            var otherProjectMock = TestProject("otherProject", ProjectProtection.Locked, vbeBuilder);
            vbeBuilder.AddProject(otherProjectMock);

            var vbe = vbeBuilder.Build().Object;

            var repository = TestRepository(vbe);
            var projects = repository.Projects().ToList();

            Assert.AreEqual(0, projects.Count);
        }

        [Test]
        [Category("COM")]
        public void ProjectsGetDisposedOnDisposal()
        {
            var vbeBuilder = new MockVbeBuilder();
            var projectMock = TestProject("project", ProjectProtection.Unprotected, vbeBuilder);
            vbeBuilder.AddProject(projectMock);
            var otherProjectMock = TestProject("otherProject", ProjectProtection.Unprotected, vbeBuilder);
            vbeBuilder.AddProject(otherProjectMock);

            var vbe = vbeBuilder.Build().Object;

            var repository = TestRepository(vbe);
            repository.Dispose();

            projectMock.Verify(m => m.Dispose(), Times.Once);
            otherProjectMock.Verify(m => m.Dispose(), Times.Once);
        }

        [Test]
        [Category("COM")]
        public void ProjectsDoNotGetDisposedOnDisposalOfReturnedValue()
        {
            var vbeBuilder = new MockVbeBuilder();
            var projectMock = TestProject("project", ProjectProtection.Unprotected, vbeBuilder);
            vbeBuilder.AddProject(projectMock);
            var otherProjectMock = TestProject("otherProject", ProjectProtection.Unprotected, vbeBuilder);
            vbeBuilder.AddProject(otherProjectMock);

            var vbe = vbeBuilder.Build().Object;

            var repository = TestRepository(vbe);
            var projects = repository.Projects().ToList();

            foreach (var (_, repositoryProject) in projects)
            {
                repositoryProject.Dispose();
            }

            projectMock.Verify(m => m.Dispose(), Times.Never);
            otherProjectMock.Verify(m => m.Dispose(), Times.Never);
        }

        [Test]
        [Category("COM")]
        public void ProjectsReturnsEmptyCollectionAfterDisposal()
        {
            var vbeBuilder = new MockVbeBuilder();
            var projectMock = TestProject("project", ProjectProtection.Unprotected, vbeBuilder);
            vbeBuilder.AddProject(projectMock);
            var otherProjectMock = TestProject("otherProject", ProjectProtection.Unprotected, vbeBuilder);
            vbeBuilder.AddProject(otherProjectMock);

            var vbe = vbeBuilder.Build().Object;

            var repository = TestRepository(vbe);
            repository.Dispose();
            var projects = repository.Projects().ToList();

            Assert.IsEmpty(projects);
        }

        [Test]
        [Category("COM")]
        public void ProjectsReturnsEmptyCollectionBeforeFirstRefresh()
        {
            var vbeBuilder = new MockVbeBuilder();
            var projectMock = TestProject("project", ProjectProtection.Unprotected, vbeBuilder);
            vbeBuilder.AddProject(projectMock);
            var otherProjectMock = TestProject("otherProject", ProjectProtection.Unprotected, vbeBuilder);
            vbeBuilder.AddProject(otherProjectMock);

            var vbe = vbeBuilder.Build().Object;

            var repository = TestRepository(vbe, initialRefresh: false);
            var projects = repository.Projects().ToList();

            Assert.IsEmpty(projects);
        }

        [Test]
        [Category("COM")]
        public void ProjectsDoesNotReturnProjectsAddedToVbeWithoutRefresh()
        {
            var vbeBuilder = new MockVbeBuilder();
            var projectMock = TestProject("project", ProjectProtection.Unprotected, vbeBuilder);
            vbeBuilder.AddProject(projectMock);

            var vbe = vbeBuilder.Build().Object;

            var repository = TestRepository(vbe);
            var otherProject = vbe.VBProjects.Add(ProjectType.HostProject);
            var projectIds = repository.Projects().Select(tpl => tpl.ProjectId).ToList();
            
            Assert.False(projectIds.Contains(otherProject.ProjectId));
        }

        [Test]
        [Category("COM")]
        public void ProjectsReturnsProjectsAddedToVbeAfterRefresh()
        {
            var vbeBuilder = new MockVbeBuilder();
            var projectMock = TestProject("project", ProjectProtection.Unprotected, vbeBuilder);
            vbeBuilder.AddProject(projectMock);

            var vbe = vbeBuilder.Build().Object;

            var repository = TestRepository(vbe);
            var otherProject = vbe.VBProjects.Add(ProjectType.HostProject);
            repository.Refresh();
            var projects = repository.Projects().ToList();

            Assert.AreEqual(2, projects.Count);

            foreach (var (projectId, repositoryProject) in projects)
            {
                Assert.AreEqual(projectId, repositoryProject.ProjectId);
                var _ = repositoryProject.VBE;
            }

            //Old Project still there.
            projectMock.Verify(m => m.VBE, Times.Once);

            //Since we have no access to the decorated project we return, we test for the id only.
            var projectIds = projects.Select(kvp => kvp.ProjectId).ToList();
            var otherProjectId = otherProject.ProjectId;
            Assert.Contains(otherProjectId, projectIds);
        }

        [Test]
        [Category("COM")]
        public void ProjectsReturnsRemovedProjectsBeforeRefresh()
        {
            var vbeBuilder = new MockVbeBuilder();
            var projectMock = TestProject("project", ProjectProtection.Unprotected, vbeBuilder);
            vbeBuilder.AddProject(projectMock);
            var otherProjectMock = TestProject("otherProject", ProjectProtection.Unprotected, vbeBuilder);
            vbeBuilder.AddProject(otherProjectMock);

            var vbe = vbeBuilder.Build().Object;
            var otherProject = otherProjectMock.Object;

            var repository = TestRepository(vbe);
            vbe.VBProjects.Remove(otherProject);
            var projects = repository.Projects().ToList();

            Assert.AreEqual(2, projects.Count);

            foreach (var (projectId, repositoryProject) in projects)
            {
                Assert.AreEqual(projectId, repositoryProject.ProjectId);
                var _ = repositoryProject.VBE;
            }

            projectMock.Verify(m => m.VBE, Times.Once);
            otherProjectMock.Verify(m => m.VBE, Times.Once);
        }

        [Test]
        [Category("COM")]
        public void ProjectsDoesNotReturnRemovedProjectsAfterRefresh()
        {
            var vbeBuilder = new MockVbeBuilder();
            var projectMock = TestProject("project", ProjectProtection.Unprotected, vbeBuilder);
            vbeBuilder.AddProject(projectMock);
            var otherProjectMock = TestProject("otherProject", ProjectProtection.Unprotected, vbeBuilder);
            vbeBuilder.AddProject(otherProjectMock);

            var vbe = vbeBuilder.Build().Object;
            var otherProject = otherProjectMock.Object;

            var repository = TestRepository(vbe);
            vbe.VBProjects.Remove(otherProject);
            repository.Refresh();
            var projects = repository.Projects().ToList();

            foreach (var (_, repositoryProject) in projects)
            {
                var _ = repositoryProject.VBE;
            }
            otherProjectMock.Verify(m => m.VBE, Times.Never);
        }

        [Test]
        [Category("COM")]
        public void ProjectsGetDisposedOnRefresh()
        {
            var vbeBuilder = new MockVbeBuilder();
            var projectMock = TestProject("project", ProjectProtection.Unprotected, vbeBuilder);
            vbeBuilder.AddProject(projectMock);
            var otherProjectMock = TestProject("otherProject", ProjectProtection.Unprotected, vbeBuilder);
            vbeBuilder.AddProject(otherProjectMock);

            var vbe = vbeBuilder.Build().Object;

            var repository = TestRepository(vbe);
            repository.Refresh();

            projectMock.Verify(m => m.Dispose(), Times.Once);
            otherProjectMock.Verify(m => m.Dispose(), Times.Once);
        }

        [Test]
        [Category("COM")]
        public void RemovedProjectsGetDisposedOnRefresh()
        {
            var vbeBuilder = new MockVbeBuilder();
            var projectMock = TestProject("project", ProjectProtection.Unprotected, vbeBuilder);
            vbeBuilder.AddProject(projectMock);
            var otherProjectMock = TestProject("otherProject", ProjectProtection.Unprotected, vbeBuilder);
            vbeBuilder.AddProject(otherProjectMock);

            var vbe = vbeBuilder.Build().Object;
            var otherProject = otherProjectMock.Object;

            var repository = TestRepository(vbe);
            vbe.VBProjects.Remove(otherProject);
            repository.Refresh();

            otherProjectMock.Verify(m => m.Dispose(), Times.Once);
        }

        [Test]
        [Category("COM")]
        public void LockedProjectsReturnsLockedProjectsOnVbe()
        {
            var vbeBuilder = new MockVbeBuilder();
            var projectMock = TestProject("project", ProjectProtection.Locked, vbeBuilder);
            vbeBuilder.AddProject(projectMock);
            var otherProjectMock = TestProject("otherProject", ProjectProtection.Locked, vbeBuilder);
            vbeBuilder.AddProject(otherProjectMock);

            var vbe = vbeBuilder.Build().Object;

            var repository = TestRepository(vbe);
            var projects = repository.LockedProjects().ToList();

            Assert.AreEqual(2, projects.Count);

            foreach (var (projectId, repositoryProject) in projects)
            {
                Assert.AreEqual(projectId, repositoryProject.ProjectId);
                var _ = repositoryProject.VBE;
            }

            projectMock.Verify(m => m.VBE, Times.Once);
            otherProjectMock.Verify(m => m.VBE, Times.Once);
        }

        [Test]
        [Category("COM")]
        public void LockedProjectsDoesNotReturnUnlockedProjectsOnVbe()
        {
            var vbeBuilder = new MockVbeBuilder();
            var projectMock = TestProject("project", ProjectProtection.Unprotected, vbeBuilder);
            vbeBuilder.AddProject(projectMock);
            var otherProjectMock = TestProject("otherProject", ProjectProtection.Unprotected, vbeBuilder);
            vbeBuilder.AddProject(otherProjectMock);

            var vbe = vbeBuilder.Build().Object;

            var repository = TestRepository(vbe);
            var projects = repository.LockedProjects().ToList();

            Assert.AreEqual(0, projects.Count);
        }

        [Test]
        [Category("COM")]
        public void LockedProjectsGetDisposedOnDisposal()
        {
            var vbeBuilder = new MockVbeBuilder();
            var projectMock = TestProject("project", ProjectProtection.Locked, vbeBuilder);
            vbeBuilder.AddProject(projectMock);
            var otherProjectMock = TestProject("otherProject", ProjectProtection.Locked, vbeBuilder);
            vbeBuilder.AddProject(otherProjectMock);

            var vbe = vbeBuilder.Build().Object;

            var repository = TestRepository(vbe);
            repository.Dispose();

            projectMock.Verify(m => m.Dispose(), Times.Once);
            otherProjectMock.Verify(m => m.Dispose(), Times.Once);
        }

        [Test]
        [Category("COM")]
        public void LockedProjectsDoNotGetDisposedOnDisposalOfReturnedValue()
        {
            var vbeBuilder = new MockVbeBuilder();
            var projectMock = TestProject("project", ProjectProtection.Locked, vbeBuilder);
            vbeBuilder.AddProject(projectMock);
            var otherProjectMock = TestProject("otherProject", ProjectProtection.Locked, vbeBuilder);
            vbeBuilder.AddProject(otherProjectMock);

            var vbe = vbeBuilder.Build().Object;

            var repository = TestRepository(vbe);
            var projects = repository.Projects().ToList();

            foreach (var (_, repositoryProject) in projects)
            {
                repositoryProject.Dispose();
            }

            projectMock.Verify(m => m.Dispose(), Times.Never);
            otherProjectMock.Verify(m => m.Dispose(), Times.Never);
        }

        [Test]
        [Category("COM")]
        public void LockedProjectsReturnsEmptyCollectionAfterDisposal()
        {
            var vbeBuilder = new MockVbeBuilder();
            var projectMock = TestProject("project", ProjectProtection.Locked, vbeBuilder);
            vbeBuilder.AddProject(projectMock);
            var otherProjectMock = TestProject("otherProject", ProjectProtection.Locked, vbeBuilder);
            vbeBuilder.AddProject(otherProjectMock);

            var vbe = vbeBuilder.Build().Object;

            var repository = TestRepository(vbe);
            repository.Dispose();
            var projects = repository.LockedProjects().ToList();

            Assert.IsEmpty(projects);
        }

        [Test]
        [Category("COM")]
        public void LockedProjectsReturnsEmptyCollectionBeforeFirstRefresh()
        {
            var vbeBuilder = new MockVbeBuilder();
            var projectMock = TestProject("project", ProjectProtection.Locked, vbeBuilder);
            vbeBuilder.AddProject(projectMock);
            var otherProjectMock = TestProject("otherProject", ProjectProtection.Locked, vbeBuilder);
            vbeBuilder.AddProject(otherProjectMock);

            var vbe = vbeBuilder.Build().Object;

            var repository = TestRepository(vbe, initialRefresh: false);
            var projects = repository.LockedProjects().ToList();

            Assert.IsEmpty(projects);
        }

        [Test]
        [Category("COM")]
        public void LockedProjectsDoesNotReturnLockedProjectsAddedToVbeWithoutRefresh()
        {
            var vbeBuilder = new MockVbeBuilder();
            var projectMock = TestProject("project", ProjectProtection.Locked, vbeBuilder);
            vbeBuilder.AddProject(projectMock);

            var vbe = vbeBuilder.Build().Object;

            var repository = TestRepository(vbe);
            var otherProject = vbe.VBProjects.Open("testPath");
            var projectIds = repository.Projects().Select(tpl => tpl.ProjectId).ToList();

            Assert.False(projectIds.Contains(otherProject.ProjectId));
        }


        [Test]
        [Category("COM")]
        public void LockedProjectsReturnsLockedProjectsAddedToVbeAfterRefresh()
        {
            var vbeBuilder = new MockVbeBuilder();
            var projectMock = TestProject("project", ProjectProtection.Locked, vbeBuilder);
            vbeBuilder.AddProject(projectMock);

            var vbe = vbeBuilder.Build().Object;

            var repository = TestRepository(vbe);
            var otherProject = vbe.VBProjects.Open("testPath");
            repository.Refresh();
            var projects = repository.LockedProjects().ToList();

            Assert.AreEqual(2, projects.Count);

            foreach (var (projectId, repositoryProject) in projects)
            {
                Assert.AreEqual(projectId, repositoryProject.ProjectId);
                var _ = repositoryProject.VBE;
            }

            //Old Project still there.
            projectMock.Verify(m => m.VBE, Times.Once);

            //Since we have no access to the decorated project we return, we test for the id only.
            var projectIds = projects.Select(kvp => kvp.ProjectId).ToList();
            var otherProjectId = otherProject.ProjectId;
            Assert.Contains(otherProjectId, projectIds);
        }

        [Test]
        [Category("COM")]
        public void LockedProjectsReturnsRemovedLockedProjectsBeforeRefresh()
        {
            var vbeBuilder = new MockVbeBuilder();
            var projectMock = TestProject("project", ProjectProtection.Locked, vbeBuilder);
            vbeBuilder.AddProject(projectMock);
            var otherProjectMock = TestProject("otherProject", ProjectProtection.Locked, vbeBuilder);
            vbeBuilder.AddProject(otherProjectMock);

            var vbe = vbeBuilder.Build().Object;
            var otherProject = otherProjectMock.Object;

            var repository = TestRepository(vbe);
            vbe.VBProjects.Remove(otherProject);
            var projects = repository.LockedProjects().ToList();

            Assert.AreEqual(2, projects.Count);

            foreach (var (projectId, repositoryProject) in projects)
            {
                Assert.AreEqual(projectId, repositoryProject.ProjectId);
                var _ = repositoryProject.VBE;
            }

            projectMock.Verify(m => m.VBE, Times.Once);
            otherProjectMock.Verify(m => m.VBE, Times.Once);
        }

        [Test]
        [Category("COM")]
        public void LockedProjectsDoesNotReturnRemovedLockedProjectsAfterRefresh()
        {
            var vbeBuilder = new MockVbeBuilder();
            var projectMock = TestProject("project", ProjectProtection.Locked, vbeBuilder);
            vbeBuilder.AddProject(projectMock);
            var otherProjectMock = TestProject("otherProject", ProjectProtection.Locked, vbeBuilder);
            vbeBuilder.AddProject(otherProjectMock);

            var vbe = vbeBuilder.Build().Object;
            var otherProject = otherProjectMock.Object;

            var repository = TestRepository(vbe);
            vbe.VBProjects.Remove(otherProject);
            repository.Refresh();
            var projects = repository.LockedProjects().ToList();

            foreach (var (_, repositoryProject) in projects)
            {
                var _ = repositoryProject.VBE;
            }
            otherProjectMock.Verify(m => m.VBE, Times.Never);
        }

        [Test]
        [Category("COM")]
        public void LockedProjectsGetDisposedOnRefresh()
        {
            var vbeBuilder = new MockVbeBuilder();
            var projectMock = TestProject("project", ProjectProtection.Locked, vbeBuilder);
            vbeBuilder.AddProject(projectMock);
            var otherProjectMock = TestProject("otherProject", ProjectProtection.Locked, vbeBuilder);
            vbeBuilder.AddProject(otherProjectMock);

            var vbe = vbeBuilder.Build().Object;

            var repository = TestRepository(vbe);
            repository.Refresh();

            projectMock.Verify(m => m.Dispose(), Times.Once);
            otherProjectMock.Verify(m => m.Dispose(), Times.Once);
        }

        [Test]
        [Category("COM")]
        public void RemovedLockedProjectsGetDisposedOnRefresh()
        {
            var vbeBuilder = new MockVbeBuilder();
            var projectMock = TestProject("project", ProjectProtection.Locked, vbeBuilder);
            vbeBuilder.AddProject(projectMock);
            var otherProjectMock = TestProject("otherProject", ProjectProtection.Locked, vbeBuilder);
            vbeBuilder.AddProject(otherProjectMock);

            var vbe = vbeBuilder.Build().Object;
            var otherProject = otherProjectMock.Object;

            var repository = TestRepository(vbe);
            vbe.VBProjects.Remove(otherProject);
            repository.Refresh();

            otherProjectMock.Verify(m => m.Dispose(), Times.Once);
        }

        [Test]
        [Category("COM")]
        public void ProjectReturnsProjectWithMatchingProjectId()
        {
            var vbeBuilder = new MockVbeBuilder();
            var projectMock = TestProject("project", ProjectProtection.Unprotected, vbeBuilder);
            vbeBuilder.AddProject(projectMock);
            var otherProjectMock = TestProject("otherProject", ProjectProtection.Unprotected, vbeBuilder);
            vbeBuilder.AddProject(otherProjectMock);

            var vbe = vbeBuilder.Build().Object;
            var otherProject = otherProjectMock.Object;

            var repository = TestRepository(vbe);
            var returnedProject = repository.Project(otherProject.ProjectId);

            Assert.AreEqual(otherProject.ProjectId, returnedProject.ProjectId);
            var _ = returnedProject.VBE;
            otherProjectMock.Verify(m => m.VBE, Times.Once);
        }

        [Test]
        [Category("COM")]
        public void ProjectReturnsNullForUnknownProjectId()
        {
            var vbeBuilder = new MockVbeBuilder();
            var projectMock = TestProject("project", ProjectProtection.Unprotected, vbeBuilder);
            vbeBuilder.AddProject(projectMock);
            var otherProjectMock = TestProject("otherProject", ProjectProtection.Unprotected, vbeBuilder);
            vbeBuilder.AddProject(otherProjectMock);

            var vbe = vbeBuilder.Build().Object;

            var repository = TestRepository(vbe);
            var returnedProject = repository.Project(new Guid().ToString());

            Assert.IsNull(returnedProject);
        }

        [Test]
        [Category("COM")]
        public void ProjectReturnsNullAfterDisposal()
        {
            var vbeBuilder = new MockVbeBuilder();
            var projectMock = TestProject("project", ProjectProtection.Unprotected, vbeBuilder);
            vbeBuilder.AddProject(projectMock);
            var otherProjectMock = TestProject("otherProject", ProjectProtection.Unprotected, vbeBuilder);
            vbeBuilder.AddProject(otherProjectMock);

            var vbe = vbeBuilder.Build().Object;
            var otherProject = otherProjectMock.Object;

            var repository = TestRepository(vbe);
            repository.Dispose();
            var returnedProject = repository.Project(otherProject.ProjectId);

            Assert.IsNull(returnedProject);
        }

        [Test]
        [Category("COM")]
        public void ProjectDoesNotGetDisposedOnDisposalOfReturnedProject()
        {
            var vbeBuilder = new MockVbeBuilder();
            var projectMock = TestProject("project", ProjectProtection.Unprotected, vbeBuilder);
            vbeBuilder.AddProject(projectMock);
            var otherProjectMock = TestProject("otherProject", ProjectProtection.Unprotected, vbeBuilder);
            vbeBuilder.AddProject(otherProjectMock);

            var vbe = vbeBuilder.Build().Object;
            var otherProject = otherProjectMock.Object;

            var repository = TestRepository(vbe);
            var returnedProject = repository.Project(otherProject.ProjectId);

            Assert.AreEqual(otherProject.ProjectId, returnedProject.ProjectId);
            returnedProject.Dispose();
            otherProjectMock.Verify(m => m.Dispose(), Times.Never);
        }

        [Test]
        [Category("COM")]
        public void ProjectReturnsNullBeforeFirstRefresh()
        {
            var vbeBuilder = new MockVbeBuilder();
            var projectMock = TestProject("project", ProjectProtection.Unprotected, vbeBuilder);
            vbeBuilder.AddProject(projectMock);
            var otherProjectMock = TestProject("otherProject", ProjectProtection.Unprotected, vbeBuilder);
            vbeBuilder.AddProject(otherProjectMock);

            var vbe = vbeBuilder.Build().Object;
            var otherProject = otherProjectMock.Object;

            var repository = TestRepository(vbe, initialRefresh: false);
            var returnedProject = repository.Project(otherProject.ProjectId);

            Assert.IsNull(returnedProject);
        }

        [Test]
        [Category("COM")]
        public void ProjectReturnsNullForProjectIdOfAddedProjectBeforeRefresh()
        {
            var vbeBuilder = new MockVbeBuilder();
            var projectMock = TestProject("project", ProjectProtection.Unprotected, vbeBuilder);
            vbeBuilder.AddProject(projectMock);

            var vbe = vbeBuilder.Build().Object;

            var repository = TestRepository(vbe);
            var otherProject = vbe.VBProjects.Add(ProjectType.HostProject);
            var returnedProject = repository.Project(otherProject.ProjectId);

            Assert.IsNull(returnedProject);
        }

        [Test]
        [Category("COM")]
        public void ProjectReturnsAddedProjectWithMatchingProjectIdAfterRefresh()
        {
            var vbeBuilder = new MockVbeBuilder();
            var projectMock = TestProject("project", ProjectProtection.Unprotected, vbeBuilder);
            vbeBuilder.AddProject(projectMock);

            var vbe = vbeBuilder.Build().Object;

            var repository = TestRepository(vbe);
            var otherProject = vbe.VBProjects.Add(ProjectType.HostProject);
            repository.Refresh();
            var returnedProject = repository.Project(otherProject.ProjectId);

            Assert.AreEqual(otherProject.ProjectId, returnedProject.ProjectId);
        }

        [Test]
        [Category("COM")]
        public void ProjectReturnsRemovedProjectWithMatchingProjectIdBeforeRefresh()
        {
            var vbeBuilder = new MockVbeBuilder();
            var projectMock = TestProject("project", ProjectProtection.Unprotected, vbeBuilder);
            vbeBuilder.AddProject(projectMock);
            var otherProjectMock = TestProject("otherProject", ProjectProtection.Unprotected, vbeBuilder);
            vbeBuilder.AddProject(otherProjectMock);

            var vbe = vbeBuilder.Build().Object;
            var otherProject = otherProjectMock.Object;

            var repository = TestRepository(vbe);
            vbe.VBProjects.Remove(otherProject);
            var returnedProject = repository.Project(otherProject.ProjectId);

            Assert.AreEqual(otherProject.ProjectId, returnedProject.ProjectId);
            var _ = returnedProject.VBE;
            otherProjectMock.Verify(m => m.VBE, Times.Once);
        }

        [Test]
        [Category("COM")]
        public void ProjectReturnsNullForProjectIdOfRemovedProjectAfterRefresh()
        {
            var vbeBuilder = new MockVbeBuilder();
            var projectMock = TestProject("project", ProjectProtection.Unprotected, vbeBuilder);
            vbeBuilder.AddProject(projectMock);
            var otherProjectMock = TestProject("otherProject", ProjectProtection.Unprotected, vbeBuilder);
            vbeBuilder.AddProject(otherProjectMock);

            var vbe = vbeBuilder.Build().Object;
            var otherProject = otherProjectMock.Object;

            var repository = TestRepository(vbe);
            vbe.VBProjects.Remove(otherProject);
            repository.Refresh();
            var returnedProject = repository.Project(otherProject.ProjectId);

            Assert.IsNull(returnedProject);
        }

        [Test]
        [Category("COM")]
        public void ComponentsCollectionsOfProjectsGetDisposedOnDisposal()
        {
            var vbeBuilder = new MockVbeBuilder();
            var (projectMock, componentsMock, _) = TestProjectAndComponents("project", ProjectProtection.Unprotected, vbeBuilder);
            vbeBuilder.AddProject(projectMock);
            var (otherProjectMock, otherComponentsMock, _) = TestProjectAndComponents("otherProject", ProjectProtection.Unprotected, vbeBuilder);
            vbeBuilder.AddProject(otherProjectMock);

            var vbe = vbeBuilder.Build().Object;

            var repository = TestRepository(vbe);
            repository.Dispose();

            componentsMock.Verify(m => m.Dispose(), Times.Once);
            otherComponentsMock.Verify(m => m.Dispose(), Times.Once);
        }

        [Test]
        [Category("COM")]
        public void ComponentsCollectionsOfProjectsGetDisposedOnRefresh()
        {
            var vbeBuilder = new MockVbeBuilder();
            var (projectMock, componentsMock, _) = TestProjectAndComponents("project", ProjectProtection.Unprotected, vbeBuilder);
            vbeBuilder.AddProject(projectMock);
            var (otherProjectMock, otherComponentsMock, _) = TestProjectAndComponents("otherProject", ProjectProtection.Unprotected, vbeBuilder);
            vbeBuilder.AddProject(otherProjectMock);

            var vbe = vbeBuilder.Build().Object;

            var repository = TestRepository(vbe);
            repository.Refresh();

            componentsMock.Verify(m => m.Dispose(), Times.Once);
            otherComponentsMock.Verify(m => m.Dispose(), Times.Once);
        }

        [Test]
        [Category("COM")]
        public void ComponentsCollectionsOfRemovedProjectsGetDisposedOnRefresh()
        {
            var vbeBuilder = new MockVbeBuilder();
            var (projectMock, componentsMock, _) = TestProjectAndComponents("project", ProjectProtection.Unprotected, vbeBuilder);
            vbeBuilder.AddProject(projectMock);
            var otherProjectMock = TestProject("otherProject", ProjectProtection.Unprotected, vbeBuilder);
            vbeBuilder.AddProject(otherProjectMock);

            var vbe = vbeBuilder.Build().Object;
            var otherProject = otherProjectMock.Object;

            var repository = TestRepository(vbe);
            vbe.VBProjects.Remove(otherProject);
            repository.Refresh();

            componentsMock.Verify(m => m.Dispose(), Times.Once);
        }

        [Test]
        [Category("COM")]
        public void ComponentsCollectionReturnsComponentsCollectionOfProjectWithMatchingProjectId()
        {
            var vbeBuilder = new MockVbeBuilder();
            var (projectMock, componentsMock, _) = TestProjectAndComponents("project", ProjectProtection.Unprotected, vbeBuilder);
            vbeBuilder.AddProject(projectMock);
            var (otherProjectMock, otherComponentsMock, _) = TestProjectAndComponents("otherProject", ProjectProtection.Unprotected, vbeBuilder);
            vbeBuilder.AddProject(otherProjectMock);

            var vbe = vbeBuilder.Build().Object;
            var project = projectMock.Object;

            var repository = TestRepository(vbe);
            var returnedCollection = repository.ComponentsCollection(project.ProjectId);
            var _ = returnedCollection.Count;

            componentsMock.Verify(m => m.Count, Times.Once);
            otherComponentsMock.Verify(m => m.Count, Times.Never);
        }

        [Test]
        [Category("COM")]
        public void ComponentsCollectionReturnsNullForUnknownProjectId()
        {
            var vbeBuilder = new MockVbeBuilder();
            var projectMock = TestProject("project", ProjectProtection.Unprotected, vbeBuilder);
            vbeBuilder.AddProject(projectMock);
            var otherProjectMock = TestProject("otherProject", ProjectProtection.Unprotected, vbeBuilder);
            vbeBuilder.AddProject(otherProjectMock);

            var vbe = vbeBuilder.Build().Object;

            var repository = TestRepository(vbe);
            var returnedCollection = repository.ComponentsCollection(new Guid().ToString());

            Assert.IsNull(returnedCollection);
        }

        [Test]
        [Category("COM")]
        public void ComponentsCollectionReturnsNullAfterDisposal()
        {
            var vbeBuilder = new MockVbeBuilder();
            var projectMock = TestProject("project", ProjectProtection.Unprotected, vbeBuilder);
            vbeBuilder.AddProject(projectMock);
            var otherProjectMock = TestProject("otherProject", ProjectProtection.Unprotected, vbeBuilder);
            vbeBuilder.AddProject(otherProjectMock);

            var vbe = vbeBuilder.Build().Object;
            var project = projectMock.Object;

            var repository = TestRepository(vbe);
            repository.Dispose();
            var returnedCollection = repository.ComponentsCollection(project.ProjectId);

            Assert.IsNull(returnedCollection);
        }

        [Test]
        [Category("COM")]
        public void ComponentsCollectionReturnsNullBeforeFirstRefresh()
        {
            var vbeBuilder = new MockVbeBuilder();
            var projectMock = TestProject("project", ProjectProtection.Unprotected, vbeBuilder);
            vbeBuilder.AddProject(projectMock);
            var otherProjectMock = TestProject("otherProject", ProjectProtection.Unprotected, vbeBuilder);
            vbeBuilder.AddProject(otherProjectMock);

            var vbe = vbeBuilder.Build().Object;
            var project = projectMock.Object;

            var repository = TestRepository(vbe, initialRefresh: false);
            var returnedCollection = repository.ComponentsCollection(project.ProjectId);

            Assert.IsNull(returnedCollection);
        }

        [Test]
        [Category("COM")]
        public void ComponentsCollectionReturnsNullForProjectIdOfAddedProjectBeforeRefresh()
        {
            var vbeBuilder = new MockVbeBuilder();
            var projectMock = TestProject("project", ProjectProtection.Unprotected, vbeBuilder);
            vbeBuilder.AddProject(projectMock);

            var vbe = vbeBuilder.Build().Object;

            var repository = TestRepository(vbe);
            var otherProject = vbe.VBProjects.Add(ProjectType.HostProject);
            var returnedCollection = repository.ComponentsCollection(otherProject.ProjectId);

            Assert.IsNull(returnedCollection);
        }

        [Test]
        [Category("COM")]
        public void ComponentCollectionReturnsComponentsCollectionOfAddedProjectWithMatchingProjectIdAfterRefresh()
        {
            var vbeBuilder = new MockVbeBuilder();
            var (projectMock, componentsMock, _) = TestProjectAndComponents("project", ProjectProtection.Unprotected, vbeBuilder);
            vbeBuilder.AddProject(projectMock);

            var vbe = vbeBuilder.Build().Object;

            var repository = TestRepository(vbe);
            var otherProject = vbe.VBProjects.Add(ProjectType.HostProject);
            repository.Refresh();
            var returnedCollection = repository.ComponentsCollection(otherProject.ProjectId);
            var _ = returnedCollection.Count;

            //Since the returned value is not null and it is not the collection of the old project, it must be the new one.
            componentsMock.Verify(m => m.Count, Times.Never);
        }

        [Test]
        [Category("COM")]
        public void ComponentsCollectionReturnsComponentsCollectionOfRemovedProjectWithMatchingProjectIdBeforeRefresh()
        {
            var vbeBuilder = new MockVbeBuilder();
            var (projectMock, componentsMock, _) = TestProjectAndComponents("project", ProjectProtection.Unprotected, vbeBuilder);
            vbeBuilder.AddProject(projectMock);
            var (otherProjectMock, otherComponentsMock, _) = TestProjectAndComponents("otherProject", ProjectProtection.Unprotected, vbeBuilder);
            vbeBuilder.AddProject(otherProjectMock);

            var vbe = vbeBuilder.Build().Object;
            var otherProject = otherProjectMock.Object;

            var repository = TestRepository(vbe);
            vbe.VBProjects.Remove(otherProject);
            var returnedCollection = repository.ComponentsCollection(otherProject.ProjectId);
            var _ = returnedCollection.Count;

            otherComponentsMock.Verify(m => m.Count, Times.Once);
            componentsMock.Verify(m => m.Count, Times.Never);
        }

        [Test]
        [Category("COM")]
        public void ComponentsCollectionReturnsNullForProjectIdOfRemovedProjectAfterRefresh()
        {
            var vbeBuilder = new MockVbeBuilder();
            var projectMock = TestProject("project", ProjectProtection.Unprotected, vbeBuilder);
            vbeBuilder.AddProject(projectMock);
            var otherProjectMock = TestProject("otherProject", ProjectProtection.Unprotected, vbeBuilder);
            vbeBuilder.AddProject(otherProjectMock);

            var vbe = vbeBuilder.Build().Object;
            var otherProject = otherProjectMock.Object;

            var repository = TestRepository(vbe);
            vbe.VBProjects.Remove(otherProject);
            repository.Refresh();
            var returnedCollection = repository.ComponentsCollection(otherProject.ProjectId);

            Assert.IsNull(returnedCollection);
        }

        [Test]
        [Category("COM")]
        public void ComponentsCollectionOfProjectWithMatchingProjectIdGetsDisposedOnRefreshForProjectId()
        {
            var vbeBuilder = new MockVbeBuilder();
            var (projectMock, componentsCollectionMock, _) = TestProjectAndComponents("project", ProjectProtection.Unprotected, vbeBuilder);
            vbeBuilder.AddProject(projectMock);
            var otherProjectMock = TestProject("otherProject", ProjectProtection.Unprotected, vbeBuilder);
            vbeBuilder.AddProject(otherProjectMock);

            var vbe = vbeBuilder.Build().Object;
            var project = projectMock.Object;

            var repository = TestRepository(vbe);
            repository.Refresh(project.ProjectId);

            componentsCollectionMock.Verify(m => m.Dispose(), Times.Once);
        }

        [Test]
        [Category("COM")]
        public void ComponentsCollectionsOfProjectsWithNonMatchingProjectIdDoNotGetDisposedOnRefreshForProjectId()
        {
            var vbeBuilder = new MockVbeBuilder();
            var projectMock = TestProject("project", ProjectProtection.Unprotected, vbeBuilder);
            vbeBuilder.AddProject(projectMock);
            var (otherProjectMock, otherComponentsCollectionMock, _) = TestProjectAndComponents("otherProject", ProjectProtection.Unprotected, vbeBuilder);
            vbeBuilder.AddProject(otherProjectMock);

            var vbe = vbeBuilder.Build().Object;
            var project = projectMock.Object;

            var repository = TestRepository(vbe);
            repository.Refresh(project.ProjectId);

            otherComponentsCollectionMock.Verify(m => m.Dispose(), Times.Never);
        }

        [Test]
        [Category("COM")]
        public void ComponentsReturnsComponentsOnVbeWithQmns()
        {
            var vbeBuilder = new MockVbeBuilder();
            var (projectMock, _, mockComponents) = TestProjectAndComponents(
                "project", 
                ProjectProtection.Unprotected, 
                vbeBuilder,
                ("component1", ComponentType.ClassModule, string.Empty),
                ("component2", ComponentType.ClassModule, string.Empty));
            vbeBuilder.AddProject(projectMock);
            var (otherProjectMock, _, otherMockComponents) = TestProjectAndComponents(
                "otherProject",
                ProjectProtection.Unprotected,
                vbeBuilder,
                ("otherComponent", ComponentType.ClassModule, string.Empty));
            vbeBuilder.AddProject(otherProjectMock);

            var vbe = vbeBuilder.Build().Object;

            var repository = TestRepository(vbe);
            var components = repository.Components().ToList();

            Assert.AreEqual(3, components.Count);

            foreach (var (qmn, component) in components)
            {
                Assert.AreEqual(component.QualifiedModuleName, qmn);
                var _ = component.VBE;
            }

            foreach (var mock in mockComponents.Concat(otherMockComponents))
            {
                mock.Verify(m => m.VBE, Times.Once);
            }
        }

        [Test]
        [Category("COM")]
        public void ComponentsGetDisposedOnDisposal()
        {
            var vbeBuilder = new MockVbeBuilder();
            var (projectMock, _, mockComponents) = TestProjectAndComponents(
                "project",
                ProjectProtection.Unprotected,
                vbeBuilder,
                ("component1", ComponentType.ClassModule, string.Empty),
                ("component2", ComponentType.ClassModule, string.Empty));
            vbeBuilder.AddProject(projectMock);
            var (otherProjectMock, _, otherMockComponents) = TestProjectAndComponents(
                "otherProject",
                ProjectProtection.Unprotected,
                vbeBuilder,
                ("otherComponent", ComponentType.ClassModule, string.Empty));
            vbeBuilder.AddProject(otherProjectMock);

            var vbe = vbeBuilder.Build().Object;

            var repository = TestRepository(vbe);
            repository.Dispose();

            foreach (var mock in mockComponents.Concat(otherMockComponents))
            {
                mock.Verify(m => m.Dispose(), Times.Once);
            }
        }

        [Test]
        [Category("COM")]
        public void ComponentsDoNotGetDisposedOnDisposalOfReturnValues()
        {
            var vbeBuilder = new MockVbeBuilder();
            var (projectMock, _, mockComponents) = TestProjectAndComponents(
                "project",
                ProjectProtection.Unprotected,
                vbeBuilder,
                ("component1", ComponentType.ClassModule, string.Empty),
                ("component2", ComponentType.ClassModule, string.Empty));
            vbeBuilder.AddProject(projectMock);
            var (otherProjectMock, _, otherMockComponents) = TestProjectAndComponents(
                "otherProject",
                ProjectProtection.Unprotected,
                vbeBuilder,
                ("otherComponent", ComponentType.ClassModule, string.Empty));
            vbeBuilder.AddProject(otherProjectMock);

            var vbe = vbeBuilder.Build().Object;

            var repository = TestRepository(vbe);
            var components = repository.Components().ToList();

            foreach (var (_, component) in components)
            {
                component.Dispose();
            }

            foreach (var mock in mockComponents.Concat(otherMockComponents))
            {
                mock.Verify(m => m.Dispose(), Times.Never);
            }
        }

        [Test]
        [Category("COM")]
        public void ComponentsReturnsEmptyCollectionAfterDisposal()
        {
            var vbeBuilder = new MockVbeBuilder();
            var projectMock = TestProject(
                "project",
                ProjectProtection.Unprotected,
                vbeBuilder,
                ("component1", ComponentType.ClassModule, string.Empty),
                ("component2", ComponentType.ClassModule, string.Empty));
            vbeBuilder.AddProject(projectMock);
            var otherProjectMock = TestProject(
                "otherProject",
                ProjectProtection.Unprotected,
                vbeBuilder,
                ("otherComponent", ComponentType.ClassModule, string.Empty));
            vbeBuilder.AddProject(otherProjectMock);

            var vbe = vbeBuilder.Build().Object;

            var repository = TestRepository(vbe);
            repository.Dispose();
            var components = repository.Components().ToList();

            Assert.IsEmpty(components);
        }

        [Test]
        [Category("COM")]
        public void ComponentsReturnsEmptyCollectionBeforeFirstRefresh()
        {
            var vbeBuilder = new MockVbeBuilder();
            var projectMock = TestProject(
                "project",
                ProjectProtection.Unprotected,
                vbeBuilder,
                ("component1", ComponentType.ClassModule, string.Empty),
                ("component2", ComponentType.ClassModule, string.Empty));
            vbeBuilder.AddProject(projectMock);
            var otherProjectMock = TestProject(
                "otherProject",
                ProjectProtection.Unprotected,
                vbeBuilder,
                ("otherComponent", ComponentType.ClassModule, string.Empty));
            vbeBuilder.AddProject(otherProjectMock);

            var vbe = vbeBuilder.Build().Object;

            var repository = TestRepository(vbe, initialRefresh: false);
            var components = repository.Components().ToList();

            Assert.IsEmpty(components);
        }

        [Test]
        [Category("COM")]
        public void ComponentsDoesNotReturnComponentsAddedToVbeBeforeRefresh()
        {
            var vbeBuilder = new MockVbeBuilder();
            var (projectMock, _, mockComponents) = TestProjectAndComponents(
                "project",
                ProjectProtection.Unprotected,
                vbeBuilder,
                ("component1", ComponentType.ClassModule, string.Empty));
            vbeBuilder.AddProject(projectMock);
            var (otherProjectMock, _, otherMockComponents) = TestProjectAndComponents(
                "otherProject",
                ProjectProtection.Unprotected,
                vbeBuilder,
                ("otherComponent", ComponentType.ClassModule, string.Empty));
            vbeBuilder.AddProject(otherProjectMock);

            var vbe = vbeBuilder.Build().Object;
            var project = projectMock.Object;

            var repository = TestRepository(vbe);
            var newComponent = project.VBComponents.Add(ComponentType.ClassModule);
            var componentQmns = repository
                .Components()
                .Select(tpl => tpl.QualifiedModuleName)
                .ToList();

            Assert.False(componentQmns.Contains(newComponent.QualifiedModuleName));
        }

        [Test]
        [Category("COM")]
        public void ComponentsDoesNotReturnComponentsAddedToVbeAfterRefreshForOtherProjectId()
        {
            var vbeBuilder = new MockVbeBuilder();
            var (projectMock, _, mockComponents) = TestProjectAndComponents(
                "project",
                ProjectProtection.Unprotected,
                vbeBuilder,
                ("component1", ComponentType.ClassModule, string.Empty));
            vbeBuilder.AddProject(projectMock);
            var (otherProjectMock, _, otherMockComponents) = TestProjectAndComponents(
                "otherProject",
                ProjectProtection.Unprotected,
                vbeBuilder,
                ("otherComponent", ComponentType.ClassModule, string.Empty));
            vbeBuilder.AddProject(otherProjectMock);

            var vbe = vbeBuilder.Build().Object;
            var project = projectMock.Object;
            var otherProject = otherProjectMock.Object;

            var repository = TestRepository(vbe);
            var newComponent = project.VBComponents.Add(ComponentType.ClassModule);
            repository.Refresh(otherProject.ProjectId);
            var componentQmns = repository
                .Components()
                .Select(tpl => tpl.QualifiedModuleName)
                .ToList();

            Assert.False(componentQmns.Contains(newComponent.QualifiedModuleName));
        }

        [Test]
        [Category("COM")]
        public void ComponentsReturnsComponentsAddedToVbeAfterRefresh()
        {
            var vbeBuilder = new MockVbeBuilder();
            var projectMock = TestProject(
                "project",
                ProjectProtection.Unprotected,
                vbeBuilder,
                ("component1", ComponentType.ClassModule, string.Empty));
            vbeBuilder.AddProject(projectMock);
            var otherProjectMock = TestProject(
                "otherProject",
                ProjectProtection.Unprotected,
                vbeBuilder,
                ("otherComponent", ComponentType.ClassModule, string.Empty));
            vbeBuilder.AddProject(otherProjectMock);

            var vbe = vbeBuilder.Build().Object;
            var project = projectMock.Object;

            var repository = TestRepository(vbe);
            var newComponent = project.VBComponents.Add(ComponentType.ClassModule);
            repository.Refresh();
            var componentQmns = repository.Components().Select(kvp => kvp.QualifiedModuleName).ToList();

            Assert.AreEqual(3, componentQmns.Count);
            Assert.Contains(newComponent.QualifiedModuleName, componentQmns);
        }

        [Test]
        [Category("COM")]
        public void ComponentsReturnsComponentsAddedToVbeAfterRefreshWithProjectIdOfContainingProject()
        {
            var vbeBuilder = new MockVbeBuilder();
            var projectMock = TestProject(
                "project",
                ProjectProtection.Unprotected,
                vbeBuilder,
                ("component1", ComponentType.ClassModule, string.Empty));
            vbeBuilder.AddProject(projectMock);
            var (otherProjectMock, _, otherMockComponents) = TestProjectAndComponents(
                "otherProject",
                ProjectProtection.Unprotected,
                vbeBuilder,
                ("otherComponent", ComponentType.ClassModule, string.Empty));
            vbeBuilder.AddProject(otherProjectMock);

            var vbe = vbeBuilder.Build().Object;
            var project = projectMock.Object;

            var repository = TestRepository(vbe);
            var newComponent = project.VBComponents.Add(ComponentType.ClassModule);
            repository.Refresh(project.ProjectId);
            var componentQmns = repository.Components().Select(kvp => kvp.QualifiedModuleName).ToList();

            Assert.AreEqual(3, componentQmns.Count);
            Assert.Contains(newComponent.QualifiedModuleName, componentQmns);
        }

        [Test]
        [Category("COM")]
        public void ComponentsReturnsRemovedComponentsBeforeRefresh()
        {
            var vbeBuilder = new MockVbeBuilder();
            var (projectMock, _, mockComponents) = TestProjectAndComponents(
                "project",
                ProjectProtection.Unprotected,
                vbeBuilder,
                ("component1", ComponentType.ClassModule, string.Empty),
                ("component2", ComponentType.ClassModule, string.Empty));
            vbeBuilder.AddProject(projectMock);
            var otherProjectMock = TestProject(
                "otherProject",
                ProjectProtection.Unprotected,
                vbeBuilder,
                ("otherComponent", ComponentType.ClassModule, string.Empty));
            vbeBuilder.AddProject(otherProjectMock);

            var vbe = vbeBuilder.Build().Object;
            var project = projectMock.Object;
            var component2Mock = mockComponents[1];
            var component2 = component2Mock.Object;

            var repository = TestRepository(vbe);
            project.VBComponents.Remove(component2);
            var components = repository.Components().ToList();

            foreach (var (_, component) in components)
            {
                var _ = component.VBE;
            }
            
            component2Mock.Verify(m => m.VBE, Times.Once);
        }

        [Test]
        [Category("COM")]
        public void ComponentsReturnsRemovedComponentsAfterRefreshForOtherProjectId()
        {
            var vbeBuilder = new MockVbeBuilder();
            var (projectMock, _, mockComponents) = TestProjectAndComponents(
                "project",
                ProjectProtection.Unprotected,
                vbeBuilder,
                ("component1", ComponentType.ClassModule, string.Empty),
                ("component2", ComponentType.ClassModule, string.Empty));
            vbeBuilder.AddProject(projectMock);
            var otherProjectMock = TestProject(
                "otherProject",
                ProjectProtection.Unprotected,
                vbeBuilder,
                ("otherComponent", ComponentType.ClassModule, string.Empty));
            vbeBuilder.AddProject(otherProjectMock);

            var vbe = vbeBuilder.Build().Object;
            var project = projectMock.Object;
            var otherProject = otherProjectMock.Object;
            var component2Mock = mockComponents[1];
            var component2 = component2Mock.Object;

            var repository = TestRepository(vbe);
            project.VBComponents.Remove(component2);
            repository.Refresh(otherProject.ProjectId);
            var components = repository.Components().ToList();

            foreach (var (_, component) in components)
            {
                var _ = component.VBE;
            }

            component2Mock.Verify(m => m.VBE, Times.Once);
        }

        [Test]
        [Category("COM")]
        public void ComponentsDoesNotReturnRemovedComponentsAfterRefresh()
        {
            var vbeBuilder = new MockVbeBuilder();
            var (projectMock, _, mockComponents) = TestProjectAndComponents(
                "project",
                ProjectProtection.Unprotected,
                vbeBuilder,
                ("component1", ComponentType.ClassModule, string.Empty),
                ("component2", ComponentType.ClassModule, string.Empty));
            vbeBuilder.AddProject(projectMock);
            var otherProjectMock = TestProject(
                "otherProject",
                ProjectProtection.Unprotected,
                vbeBuilder,
                ("otherComponent", ComponentType.ClassModule, string.Empty));
            vbeBuilder.AddProject(otherProjectMock);

            var vbe = vbeBuilder.Build().Object;
            var project = projectMock.Object;
            var component2Mock = mockComponents[1];
            var component2 = component2Mock.Object;

            var repository = TestRepository(vbe);
            project.VBComponents.Remove(component2);
            repository.Refresh();
            var components = repository.Components().ToList();

            foreach (var (_, component) in components)
            {
                var _ = component.VBE;
            }

            component2Mock.Verify(m => m.VBE, Times.Never());
        }

        [Test]
        [Category("COM")]
        public void ComponentsDoesNotReturnRemovedComponentsAfterRefreshForProjectIdOfContainingProject()
        {
            var vbeBuilder = new MockVbeBuilder();
            var (projectMock, _, mockComponents) = TestProjectAndComponents(
                "project",
                ProjectProtection.Unprotected,
                vbeBuilder,
                ("component1", ComponentType.ClassModule, string.Empty),
                ("component2", ComponentType.ClassModule, string.Empty));
            vbeBuilder.AddProject(projectMock);
            var otherProjectMock = TestProject(
                "otherProject",
                ProjectProtection.Unprotected,
                vbeBuilder,
                ("otherComponent", ComponentType.ClassModule, string.Empty));
            vbeBuilder.AddProject(otherProjectMock);

            var vbe = vbeBuilder.Build().Object;
            var project = projectMock.Object;
            var component2Mock = mockComponents[1];
            var component2 = component2Mock.Object;

            var repository = TestRepository(vbe);
            project.VBComponents.Remove(component2);
            repository.Refresh(project.ProjectId);
            var components = repository.Components().ToList();

            foreach (var (_, component) in components)
            {
                var _ = component.VBE;
            }

            component2Mock.Verify(m => m.VBE, Times.Never());
        }

        [Test]
        [Category("COM")]
        public void ComponentsGetDisposedOnRefresh()
        {
            var vbeBuilder = new MockVbeBuilder();
            var (projectMock, _, mockComponents) = TestProjectAndComponents(
                "project",
                ProjectProtection.Unprotected,
                vbeBuilder,
                ("component1", ComponentType.ClassModule, string.Empty));
            vbeBuilder.AddProject(projectMock);
            var (otherProjectMock, _, otherMockComponents) = TestProjectAndComponents(
                "otherProject",
                ProjectProtection.Unprotected,
                vbeBuilder,
                ("otherComponent", ComponentType.ClassModule, string.Empty));
            vbeBuilder.AddProject(otherProjectMock);

            var vbe = vbeBuilder.Build().Object;

            var repository = TestRepository(vbe);
            repository.Refresh();

            foreach (var mock in mockComponents.Concat(otherMockComponents))
            {
                mock.Verify(m => m.Dispose(), Times.Once);
            }
        }

        [Test]
        [Category("COM")]
        public void ComponentsInProjectWithMatchingProjectIdGetDisposedOnRefreshForProjectId()
        {
            var vbeBuilder = new MockVbeBuilder();
            var (projectMock, _, mockComponents) = TestProjectAndComponents(
                "project",
                ProjectProtection.Unprotected,
                vbeBuilder,
                ("component1", ComponentType.ClassModule, string.Empty),
                ("component2", ComponentType.ClassModule, string.Empty));
            vbeBuilder.AddProject(projectMock);
            var otherProjectMock = TestProject(
                "otherProject",
                ProjectProtection.Unprotected,
                vbeBuilder,
                ("otherComponent", ComponentType.ClassModule, string.Empty));
            vbeBuilder.AddProject(otherProjectMock);

            var vbe = vbeBuilder.Build().Object;
            var project = projectMock.Object;

            var repository = TestRepository(vbe);
            repository.Refresh(project.ProjectId);

            foreach (var mock in mockComponents)
            {
                mock.Verify(m => m.Dispose(), Times.Once);
            }
        }

        [Test]
        [Category("COM")]
        public void ComponentsInProjectWithOtherProjectIdDoNotGetDisposedOnRefreshForProjectId()
        {
            var vbeBuilder = new MockVbeBuilder();
            var projectMock = TestProject(
                "project",
                ProjectProtection.Unprotected,
                vbeBuilder,
                ("component1", ComponentType.ClassModule, string.Empty));
            vbeBuilder.AddProject(projectMock);
            var (otherProjectMock, _, otherMockComponents) = TestProjectAndComponents(
                "otherProject",
                ProjectProtection.Unprotected,
                vbeBuilder,
                ("otherComponent", ComponentType.ClassModule, string.Empty));
            vbeBuilder.AddProject(otherProjectMock);

            var vbe = vbeBuilder.Build().Object;
            var project = projectMock.Object;

            var repository = TestRepository(vbe);
            repository.Refresh(project.ProjectId);

            foreach (var mock in otherMockComponents)
            {
                mock.Verify(m => m.Dispose(), Times.Never);
            }
        }

        [Test]
        [Category("COM")]
        public void RemovedComponentsGetDisposedOnRefresh()
        {
            var vbeBuilder = new MockVbeBuilder();
            var (projectMock, _, mockComponents) = TestProjectAndComponents(
                "project",
                ProjectProtection.Unprotected,
                vbeBuilder,
                ("component1", ComponentType.ClassModule, string.Empty),
                ("component2", ComponentType.ClassModule, string.Empty));
            vbeBuilder.AddProject(projectMock);
            var otherProjectMock = TestProject(
                "otherProject",
                ProjectProtection.Unprotected,
                vbeBuilder,
                ("otherComponent", ComponentType.ClassModule, string.Empty));
            vbeBuilder.AddProject(otherProjectMock);

            var vbe = vbeBuilder.Build().Object;
            var project = projectMock.Object;
            var component2Mock = mockComponents[1];
            var component2 = component2Mock.Object;

            var repository = TestRepository(vbe);
            project.VBComponents.Remove(component2);
            repository.Refresh();

            component2Mock.Verify(m => m.Dispose(), Times.Once);
        }

        [Test]
        [Category("COM")]
        public void RemovedComponentsGetDisposedOnRefreshForProjectIdOfFormerlyContainingProject()
        {
            var vbeBuilder = new MockVbeBuilder();
            var (projectMock, _, mockComponents) = TestProjectAndComponents(
                "project",
                ProjectProtection.Unprotected,
                vbeBuilder,
                ("component1", ComponentType.ClassModule, string.Empty),
                ("component2", ComponentType.ClassModule, string.Empty));
            vbeBuilder.AddProject(projectMock);
            var otherProjectMock = TestProject(
                "otherProject",
                ProjectProtection.Unprotected,
                vbeBuilder,
                ("otherComponent", ComponentType.ClassModule, string.Empty));
            vbeBuilder.AddProject(otherProjectMock);

            var vbe = vbeBuilder.Build().Object;
            var project = projectMock.Object;
            var component2Mock = mockComponents[1];
            var component2 = component2Mock.Object;

            var repository = TestRepository(vbe);
            project.VBComponents.Remove(component2);
            repository.Refresh(project.ProjectId);

            component2Mock.Verify(m => m.Dispose(), Times.Once);
        }

        [Test]
        [Category("COM")]
        public void RemovedComponentsDoNotGetDisposedOnRefreshForOtherProjectId()
        {
            var vbeBuilder = new MockVbeBuilder();
            var (projectMock, _, mockComponents) = TestProjectAndComponents(
                "project",
                ProjectProtection.Unprotected,
                vbeBuilder,
                ("component1", ComponentType.ClassModule, string.Empty),
                ("component2", ComponentType.ClassModule, string.Empty));
            vbeBuilder.AddProject(projectMock);
            var otherProjectMock = TestProject(
                "otherProject",
                ProjectProtection.Unprotected,
                vbeBuilder,
                ("otherComponent", ComponentType.ClassModule, string.Empty));
            vbeBuilder.AddProject(otherProjectMock);

            var vbe = vbeBuilder.Build().Object;
            var project = projectMock.Object;
            var otherProject = otherProjectMock.Object;
            var component2Mock = mockComponents[1];
            var component2 = component2Mock.Object;

            var repository = TestRepository(vbe);
            project.VBComponents.Remove(component2);
            repository.Refresh(otherProject.ProjectId);

            component2Mock.Verify(m => m.Dispose(), Times.Never);
        }

        [Test]
        [Category("COM")]
        public void ComponentsForProjectIdReturnsComponentsOnProjectWithMatchingProjectId()
        {
            var vbeBuilder = new MockVbeBuilder();
            var (projectMock, _, mockComponents) = TestProjectAndComponents(
                "project",
                ProjectProtection.Unprotected,
                vbeBuilder,
                ("component1", ComponentType.ClassModule, string.Empty),
                ("component2", ComponentType.ClassModule, string.Empty));
            vbeBuilder.AddProject(projectMock);
            var otherProjectMock = TestProject(
                "otherProject",
                ProjectProtection.Unprotected,
                vbeBuilder,
                ("otherComponent", ComponentType.ClassModule, string.Empty));
            vbeBuilder.AddProject(otherProjectMock);

            var vbe = vbeBuilder.Build().Object;
            var project = projectMock.Object;

            var repository = TestRepository(vbe);
            var components = repository.Components(project.ProjectId).ToList();

            Assert.AreEqual(2, components.Count);
            foreach (var (qmn, component) in components)
            {
                Assert.AreEqual(component.QualifiedModuleName, qmn);
                var _ = component.VBE;
            }
            foreach (var mock in mockComponents)
            {
                mock.Verify(m => m.VBE, Times.Once);
            }
        }

        [Test]
        [Category("COM")]
        public void ComponentsForProjectIdDoNotGetDisposedAfterDisposalOfreturnedComponents()
        {
            var vbeBuilder = new MockVbeBuilder();
            var (projectMock, _, mockComponents) = TestProjectAndComponents(
                "project",
                ProjectProtection.Unprotected,
                vbeBuilder,
                ("component1", ComponentType.ClassModule, string.Empty),
                ("component2", ComponentType.ClassModule, string.Empty));
            vbeBuilder.AddProject(projectMock);
            var otherProjectMock = TestProject(
                "otherProject",
                ProjectProtection.Unprotected,
                vbeBuilder,
                ("otherComponent", ComponentType.ClassModule, string.Empty));
            vbeBuilder.AddProject(otherProjectMock);

            var vbe = vbeBuilder.Build().Object;
            var project = projectMock.Object;

            var repository = TestRepository(vbe);
            var components = repository.Components(project.ProjectId).ToList();

            foreach (var (_, component) in components)
            {
                component.Dispose();
            }
            foreach (var mock in mockComponents)
            {
                mock.Verify(m => m.Dispose(), Times.Never);
            }
        }

        [Test]
        [Category("COM")]
        public void ComponentsForProjectIdReturnsEmptyCollectionAfterDisposal()
        {
            var vbeBuilder = new MockVbeBuilder();
            var projectMock = TestProject(
                "project",
                ProjectProtection.Unprotected,
                vbeBuilder,
                ("component1", ComponentType.ClassModule, string.Empty),
                ("component2", ComponentType.ClassModule, string.Empty));
            vbeBuilder.AddProject(projectMock);
            var otherProjectMock = TestProject(
                "otherProject",
                ProjectProtection.Unprotected,
                vbeBuilder,
                ("otherComponent", ComponentType.ClassModule, string.Empty));
            vbeBuilder.AddProject(otherProjectMock);

            var vbe = vbeBuilder.Build().Object;
            var project = projectMock.Object;

            var repository = TestRepository(vbe);
            repository.Dispose();
            var components = repository.Components(project.ProjectId).ToList();

            Assert.IsEmpty(components);
        }

        [Test]
        [Category("COM")]
        public void ComponentsForProjectIdReturnsEmptyCollectionBeforeFirstRefresh()
        {
            var vbeBuilder = new MockVbeBuilder();
            var projectMock = TestProject(
                "project",
                ProjectProtection.Unprotected,
                vbeBuilder,
                ("component1", ComponentType.ClassModule, string.Empty),
                ("component2", ComponentType.ClassModule, string.Empty));
            vbeBuilder.AddProject(projectMock);
            var otherProjectMock = TestProject(
                "otherProject",
                ProjectProtection.Unprotected,
                vbeBuilder,
                ("otherComponent", ComponentType.ClassModule, string.Empty));
            vbeBuilder.AddProject(otherProjectMock);

            var vbe = vbeBuilder.Build().Object;
            var project = projectMock.Object;

            var repository = TestRepository(vbe, initialRefresh: false);
            var components = repository.Components(project.ProjectId).ToList();

            Assert.IsEmpty(components);
        }

        [Test]
        [Category("COM")]
        public void ComponentsForProjectIdDoesNotReturnComponentsAddedToVbeBeforeRefresh()
        {
            var vbeBuilder = new MockVbeBuilder();
            var (projectMock, _, mockComponents) = TestProjectAndComponents(
                "project",
                ProjectProtection.Unprotected,
                vbeBuilder,
                ("component1", ComponentType.ClassModule, string.Empty));
            vbeBuilder.AddProject(projectMock);
            var (otherProjectMock, _, otherMockComponents) = TestProjectAndComponents(
                "otherProject",
                ProjectProtection.Unprotected,
                vbeBuilder,
                ("otherComponent", ComponentType.ClassModule, string.Empty));
            vbeBuilder.AddProject(otherProjectMock);

            var vbe = vbeBuilder.Build().Object;
            var project = projectMock.Object;

            var repository = TestRepository(vbe);
            var newComponent = project.VBComponents.Add(ComponentType.ClassModule);
            var componentQmns = repository
                .Components(project.ProjectId)
                .Select(tpl => tpl.QualifiedModuleName)
                .ToList();

            Assert.False(componentQmns.Contains(newComponent.QualifiedModuleName));
        }

        [Test]
        [Category("COM")]
        public void ComponentsForProjectIdDoesNotReturnComponentsAddedToVbeAfterRefreshForOtherProjectId()
        {
            var vbeBuilder = new MockVbeBuilder();
            var projectMock = TestProject(
                "project",
                ProjectProtection.Unprotected,
                vbeBuilder,
                ("component1", ComponentType.ClassModule, string.Empty));
            vbeBuilder.AddProject(projectMock);
            var otherProjectMock = TestProject(
                "otherProject",
                ProjectProtection.Unprotected,
                vbeBuilder,
                ("otherComponent", ComponentType.ClassModule, string.Empty));
            vbeBuilder.AddProject(otherProjectMock);

            var vbe = vbeBuilder.Build().Object;
            var project = projectMock.Object;
            var otherProject = otherProjectMock.Object;

            var repository = TestRepository(vbe);
            var newComponent = project.VBComponents.Add(ComponentType.ClassModule);
            repository.Refresh(otherProject.ProjectId);
            var componentQmns = repository
                .Components(project.ProjectId)
                .Select(tpl => tpl.QualifiedModuleName)
                .ToList();

            Assert.False(componentQmns.Contains(newComponent.QualifiedModuleName));
        }

        [Test]
        [Category("COM")]
        public void ComponentsForProjectIdReturnsComponentsAddedToVbeAfterRefresh()
        {
            var vbeBuilder = new MockVbeBuilder();
            var projectMock = TestProject(
                "project",
                ProjectProtection.Unprotected,
                vbeBuilder,
                ("component1", ComponentType.ClassModule, string.Empty));
            vbeBuilder.AddProject(projectMock);
            var otherProjectMock = TestProject(
                "otherProject",
                ProjectProtection.Unprotected,
                vbeBuilder,
                ("otherComponent", ComponentType.ClassModule, string.Empty));
            vbeBuilder.AddProject(otherProjectMock);

            var vbe = vbeBuilder.Build().Object;
            var project = projectMock.Object;

            var repository = TestRepository(vbe);
            var newComponent = project.VBComponents.Add(ComponentType.ClassModule);
            repository.Refresh();
            var componentQmns = repository
                .Components(project.ProjectId)
                .Select(tpl => tpl.QualifiedModuleName)
                .ToList();

            Assert.Contains(newComponent.QualifiedModuleName, componentQmns);
        }

        [Test]
        [Category("COM")]
        public void ComponentsForProjectIdReturnsComponentsAddedToVbeAfterRefreshWithProjectIdOfContainingProject()
        {
            var vbeBuilder = new MockVbeBuilder();
            var projectMock = TestProject(
                "project",
                ProjectProtection.Unprotected,
                vbeBuilder,
                ("component1", ComponentType.ClassModule, string.Empty));
            vbeBuilder.AddProject(projectMock);
            var otherProjectMock = TestProject(
                "otherProject",
                ProjectProtection.Unprotected,
                vbeBuilder,
                ("otherComponent", ComponentType.ClassModule, string.Empty));
            vbeBuilder.AddProject(otherProjectMock);

            var vbe = vbeBuilder.Build().Object;
            var project = projectMock.Object;

            var repository = TestRepository(vbe);
            var newComponent = project.VBComponents.Add(ComponentType.ClassModule);
            repository.Refresh(project.ProjectId);
            var componentQmns = repository
                .Components(project.ProjectId)
                .Select(tpl => tpl.QualifiedModuleName)
                .ToList();

            Assert.Contains(newComponent.QualifiedModuleName, componentQmns);
        }

        [Test]
        [Category("COM")]
        public void ComponentsForProjectIdReturnsRemovedComponentsBeforeRefresh()
        {
            var vbeBuilder = new MockVbeBuilder();
            var (projectMock, _, mockComponents) = TestProjectAndComponents(
                "project",
                ProjectProtection.Unprotected,
                vbeBuilder,
                ("component1", ComponentType.ClassModule, string.Empty),
                ("component2", ComponentType.ClassModule, string.Empty));
            vbeBuilder.AddProject(projectMock);
            var otherProjectMock = TestProject(
                "otherProject",
                ProjectProtection.Unprotected,
                vbeBuilder,
                ("otherComponent", ComponentType.ClassModule, string.Empty));
            vbeBuilder.AddProject(otherProjectMock);

            var vbe = vbeBuilder.Build().Object;
            var project = projectMock.Object;
            var component2Mock = mockComponents[1];
            var component2 = component2Mock.Object;

            var repository = TestRepository(vbe);
            project.VBComponents.Remove(component2);
            var components = repository.Components(project.ProjectId).ToList();

            foreach (var (_, component) in components)
            {
                var _ = component.VBE;
            }

            component2Mock.Verify(m => m.VBE, Times.Once);
        }

        [Test]
        [Category("COM")]
        public void ComponentsForProjectIdReturnsRemovedComponentsAfterRefreshForOtherProjectId()
        {
            var vbeBuilder = new MockVbeBuilder();
            var (projectMock, _, mockComponents) = TestProjectAndComponents(
                "project",
                ProjectProtection.Unprotected,
                vbeBuilder,
                ("component1", ComponentType.ClassModule, string.Empty),
                ("component2", ComponentType.ClassModule, string.Empty));
            vbeBuilder.AddProject(projectMock);
            var otherProjectMock = TestProject(
                "otherProject",
                ProjectProtection.Unprotected,
                vbeBuilder,
                ("otherComponent", ComponentType.ClassModule, string.Empty));
            vbeBuilder.AddProject(otherProjectMock);

            var vbe = vbeBuilder.Build().Object;
            var project = projectMock.Object;
            var otherProject = otherProjectMock.Object;
            var component2Mock = mockComponents[1];
            var component2 = component2Mock.Object;

            var repository = TestRepository(vbe);
            project.VBComponents.Remove(component2);
            repository.Refresh(otherProject.ProjectId);
            var components = repository.Components(project.ProjectId).ToList();

            foreach (var (_, component) in components)
            {
                var _ = component.VBE;
            }

            component2Mock.Verify(m => m.VBE, Times.Once);
        }

        [Test]
        [Category("COM")]
        public void ComponentsForProjectIdDoesNotReturnRemovedComponentsAfterRefresh()
        {
            var vbeBuilder = new MockVbeBuilder();
            var (projectMock, _, mockComponents) = TestProjectAndComponents(
                "project",
                ProjectProtection.Unprotected,
                vbeBuilder,
                ("component1", ComponentType.ClassModule, string.Empty),
                ("component2", ComponentType.ClassModule, string.Empty));
            vbeBuilder.AddProject(projectMock);
            var otherProjectMock = TestProject(
                "otherProject",
                ProjectProtection.Unprotected,
                vbeBuilder,
                ("otherComponent", ComponentType.ClassModule, string.Empty));
            vbeBuilder.AddProject(otherProjectMock);

            var vbe = vbeBuilder.Build().Object;
            var project = projectMock.Object;
            var component2Mock = mockComponents[1];
            var component2 = component2Mock.Object;

            var repository = TestRepository(vbe);
            project.VBComponents.Remove(component2);
            repository.Refresh();
            var components = repository.Components(project.ProjectId).ToList();

            foreach (var (_, component) in components)
            {
                var _ = component.VBE;
            }

            component2Mock.Verify(m => m.VBE, Times.Never);
        }

        [Test]
        [Category("COM")]
        public void ComponentsForProjectIdDoesNotReturnRemovedComponentsAfterRefreshForProjectIdOfContainingProject()
        {
            var vbeBuilder = new MockVbeBuilder();
            var (projectMock, _, mockComponents) = TestProjectAndComponents(
                "project",
                ProjectProtection.Unprotected,
                vbeBuilder,
                ("component1", ComponentType.ClassModule, string.Empty),
                ("component2", ComponentType.ClassModule, string.Empty));
            vbeBuilder.AddProject(projectMock);
            var otherProjectMock = TestProject(
                "otherProject",
                ProjectProtection.Unprotected,
                vbeBuilder,
                ("otherComponent", ComponentType.ClassModule, string.Empty));
            vbeBuilder.AddProject(otherProjectMock);

            var vbe = vbeBuilder.Build().Object;
            var project = projectMock.Object;
            var component2Mock = mockComponents[1];
            var component2 = component2Mock.Object;

            var repository = TestRepository(vbe);
            project.VBComponents.Remove(component2);
            repository.Refresh(project.ProjectId);
            var components = repository.Components(project.ProjectId).ToList();

            foreach (var (_, component) in components)
            {
                var _ = component.VBE;
            }

            component2Mock.Verify(m => m.VBE, Times.Never);
        }

        [Test]
        [Category("COM")]
        public void ComponentReturnsComponentWithMatchingQmn()
        {
            var vbeBuilder = new MockVbeBuilder();
            var (projectMock, _, mockComponents) = TestProjectAndComponents(
                "project",
                ProjectProtection.Unprotected,
                vbeBuilder,
                ("component1", ComponentType.ClassModule, string.Empty),
                ("component2", ComponentType.ClassModule, string.Empty));
            vbeBuilder.AddProject(projectMock);
            var otherProjectMock = TestProject(
                "otherProject",
                ProjectProtection.Unprotected,
                vbeBuilder,
                ("otherComponent", ComponentType.ClassModule, string.Empty));
            vbeBuilder.AddProject(otherProjectMock);

            var vbe = vbeBuilder.Build().Object;
            var component2Mock = mockComponents[1];
            var component2 = component2Mock.Object;

            var repository = TestRepository(vbe);
            var returnedComponent = repository.Component(component2.QualifiedModuleName);

            var _ = returnedComponent.VBE;

            component2Mock.Verify(m => m.VBE, Times.Once);
        }

        [Test]
        [Category("COM")]
        public void ComponentDoesNotGetDisposedOnDisposalOfReturnedComponent()
        {
            var vbeBuilder = new MockVbeBuilder();
            var (projectMock, _, mockComponents) = TestProjectAndComponents(
                "project",
                ProjectProtection.Unprotected,
                vbeBuilder,
                ("component1", ComponentType.ClassModule, string.Empty),
                ("component2", ComponentType.ClassModule, string.Empty));
            vbeBuilder.AddProject(projectMock);
            var otherProjectMock = TestProject(
                "otherProject",
                ProjectProtection.Unprotected,
                vbeBuilder,
                ("otherComponent", ComponentType.ClassModule, string.Empty));
            vbeBuilder.AddProject(otherProjectMock);

            var vbe = vbeBuilder.Build().Object;
            var component2Mock = mockComponents[1];
            var component2 = component2Mock.Object;

            var repository = TestRepository(vbe);
            var returnedComponent = repository.Component(component2.QualifiedModuleName);

            returnedComponent.Dispose();

            component2Mock.Verify(m => m.Dispose(), Times.Never);
        }

        [Test]
        [Category("COM")]
        public void ComponentReturnsNullForUnknownQmn()
        {
            var vbeBuilder = new MockVbeBuilder();
            var projectMock = TestProject(
                "project",
                ProjectProtection.Unprotected,
                vbeBuilder,
                ("component1", ComponentType.ClassModule, string.Empty),
                ("component2", ComponentType.ClassModule, string.Empty));
            vbeBuilder.AddProject(projectMock);
            var otherProjectMock = TestProject(
                "otherProject",
                ProjectProtection.Unprotected,
                vbeBuilder,
                ("otherComponent", ComponentType.ClassModule, string.Empty));
            vbeBuilder.AddProject(otherProjectMock);

            var vbe = vbeBuilder.Build().Object;

            var repository = TestRepository(vbe);
            var newQmn = new QualifiedModuleName(string.Empty, string.Empty, "newComponent");
            var returnedComponent = repository.Component(newQmn);

            Assert.IsNull(returnedComponent);
        }

        [Test]
        [Category("COM")]
        public void ComponentReturnsNullAfterDisposal()
        {
            var vbeBuilder = new MockVbeBuilder();
            var (projectMock, _, mockComponents) = TestProjectAndComponents(
                "project",
                ProjectProtection.Unprotected,
                vbeBuilder,
                ("component1", ComponentType.ClassModule, string.Empty),
                ("component2", ComponentType.ClassModule, string.Empty));
            vbeBuilder.AddProject(projectMock);
            var otherProjectMock = TestProject(
                "otherProject",
                ProjectProtection.Unprotected,
                vbeBuilder,
                ("otherComponent", ComponentType.ClassModule, string.Empty));
            vbeBuilder.AddProject(otherProjectMock);

            var vbe = vbeBuilder.Build().Object;
            var component2 = mockComponents[1].Object;

            var repository = TestRepository(vbe);
            repository.Dispose();
            var returnedComponent = repository.Component(component2.QualifiedModuleName);

            Assert.IsNull(returnedComponent);
        }

        [Test]
        [Category("COM")]
        public void ComponentReturnsNullBeforeFirstRefresh()
        {
            var vbeBuilder = new MockVbeBuilder();
            var (projectMock, _, mockComponents) = TestProjectAndComponents(
                "project",
                ProjectProtection.Unprotected,
                vbeBuilder,
                ("component1", ComponentType.ClassModule, string.Empty),
                ("component2", ComponentType.ClassModule, string.Empty));
            vbeBuilder.AddProject(projectMock);
            var otherProjectMock = TestProject(
                "otherProject",
                ProjectProtection.Unprotected,
                vbeBuilder,
                ("otherComponent", ComponentType.ClassModule, string.Empty));
            vbeBuilder.AddProject(otherProjectMock);

            var vbe = vbeBuilder.Build().Object;
            var component2 = mockComponents[1].Object;

            var repository = TestRepository(vbe, initialRefresh: false);
            var returnedComponent = repository.Component(component2.QualifiedModuleName);

            Assert.IsNull(returnedComponent);
        }

        [Test]
        [Category("COM")]
        public void ComponentReturnsNullForQmnOfAddedComponentBeforeRefresh()
        {
            var vbeBuilder = new MockVbeBuilder();
            var projectMock = TestProject(
                "project",
                ProjectProtection.Unprotected,
                vbeBuilder,
                ("component1", ComponentType.ClassModule, string.Empty));
            vbeBuilder.AddProject(projectMock);
            var otherProjectMock = TestProject(
                "otherProject",
                ProjectProtection.Unprotected,
                vbeBuilder,
                ("otherComponent", ComponentType.ClassModule, string.Empty));
            vbeBuilder.AddProject(otherProjectMock);

            var vbe = vbeBuilder.Build().Object;
            var project = projectMock.Object;

            var repository = TestRepository(vbe);
            var newComponent = project.VBComponents.Add(ComponentType.ClassModule);
            var returnedComponent = repository.Component(newComponent.QualifiedModuleName);

            Assert.IsNull(returnedComponent);
        }

        [Test]
        [Category("COM")]
        public void ComponentReturnsNullForQmnOfAddedComponentAfterRefreshForOtherProjectId()
        {
            var vbeBuilder = new MockVbeBuilder();
            var projectMock = TestProject(
                "project",
                ProjectProtection.Unprotected,
                vbeBuilder,
                ("component1", ComponentType.ClassModule, string.Empty));
            vbeBuilder.AddProject(projectMock);
            var otherProjectMock = TestProject(
                "otherProject",
                ProjectProtection.Unprotected,
                vbeBuilder,
                ("otherComponent", ComponentType.ClassModule, string.Empty));
            vbeBuilder.AddProject(otherProjectMock);

            var vbe = vbeBuilder.Build().Object;
            var project = projectMock.Object;
            var otherProject = otherProjectMock.Object;

            var repository = TestRepository(vbe);
            var newComponent = project.VBComponents.Add(ComponentType.ClassModule);
            repository.Refresh(otherProject.ProjectId);
            var returnedComponent = repository.Component(newComponent.QualifiedModuleName);

            Assert.IsNull(returnedComponent);
        }

        [Test]
        [Category("COM")]
        public void ComponentReturnsAddedComponentWithMatchingQmnAfterRefresh()
        {
            var vbeBuilder = new MockVbeBuilder();
            var (projectMock, _, mockComponents) = TestProjectAndComponents(
                "project",
                ProjectProtection.Unprotected,
                vbeBuilder,
                ("component1", ComponentType.ClassModule, string.Empty),
                ("component2", ComponentType.ClassModule, string.Empty));
            vbeBuilder.AddProject(projectMock);
            var otherProjectMock = TestProject(
                "otherProject",
                ProjectProtection.Unprotected,
                vbeBuilder,
                ("otherComponent", ComponentType.ClassModule, string.Empty));
            vbeBuilder.AddProject(otherProjectMock);

            var vbe = vbeBuilder.Build().Object;
            var project = projectMock.Object;

            var repository = TestRepository(vbe);
            var newComponent = project.VBComponents.Add(ComponentType.ClassModule);
            var newComponentMock = mockComponents.Single(m => m.Object.QualifiedModuleName.Equals(newComponent.QualifiedModuleName));
            repository.Refresh();
            var returnedComponent = repository.Component(newComponent.QualifiedModuleName);

            var _ = returnedComponent.VBE;
            newComponentMock.Verify(m => m.VBE, Times.Once);
        }

        [Test]
        [Category("COM")]
        public void ComponentReturnsAddedComponentWithMatchingQmnAfterRefreshForProjectIdOfContainingProject()
        {
            var vbeBuilder = new MockVbeBuilder();
            var (projectMock, _, mockComponents) = TestProjectAndComponents(
                "project",
                ProjectProtection.Unprotected,
                vbeBuilder,
                ("component1", ComponentType.ClassModule, string.Empty),
                ("component2", ComponentType.ClassModule, string.Empty));
            vbeBuilder.AddProject(projectMock);
            var otherProjectMock = TestProject(
                "otherProject",
                ProjectProtection.Unprotected,
                vbeBuilder,
                ("otherComponent", ComponentType.ClassModule, string.Empty));
            vbeBuilder.AddProject(otherProjectMock);

            var vbe = vbeBuilder.Build().Object;
            var project = projectMock.Object;

            var repository = TestRepository(vbe);
            var newComponent = project.VBComponents.Add(ComponentType.ClassModule);
            var newComponentMock = mockComponents.Single(m => m.Object.QualifiedModuleName.Equals(newComponent.QualifiedModuleName));
            repository.Refresh(project.ProjectId);
            var returnedComponent = repository.Component(newComponent.QualifiedModuleName);

            var _ = returnedComponent.VBE;
            newComponentMock.Verify(m => m.VBE, Times.Once);
        }

        [Test]
        [Category("COM")]
        public void ComponentReturnsRemovedComponentWithMatchingQmnBeforeRefresh()
        {
            var vbeBuilder = new MockVbeBuilder();
            var (projectMock, _, mockComponents) = TestProjectAndComponents(
                "project",
                ProjectProtection.Unprotected,
                vbeBuilder,
                ("component1", ComponentType.ClassModule, string.Empty),
                ("component2", ComponentType.ClassModule, string.Empty));
            vbeBuilder.AddProject(projectMock);
            var otherProjectMock = TestProject(
                "otherProject",
                ProjectProtection.Unprotected,
                vbeBuilder,
                ("otherComponent", ComponentType.ClassModule, string.Empty));
            vbeBuilder.AddProject(otherProjectMock);

            var vbe = vbeBuilder.Build().Object;
            var project = projectMock.Object;
            var component2Mock = mockComponents[1];
            var component2 = component2Mock.Object;

            var repository = TestRepository(vbe);
            project.VBComponents.Remove(component2);
            var returnedComponent = repository.Component(component2.QualifiedModuleName);

            var _ = returnedComponent.VBE;
            component2Mock.Verify(m => m.VBE, Times.Once);
        }

        [Test]
        [Category("COM")]
        public void ComponentReturnsRemovedComponentWithMatchingQmnAfterRefreshFroOtherProjectId()
        {
            var vbeBuilder = new MockVbeBuilder();
            var (projectMock, _, mockComponents) = TestProjectAndComponents(
                "project",
                ProjectProtection.Unprotected,
                vbeBuilder,
                ("component1", ComponentType.ClassModule, string.Empty),
                ("component2", ComponentType.ClassModule, string.Empty));
            vbeBuilder.AddProject(projectMock);
            var otherProjectMock = TestProject(
                "otherProject",
                ProjectProtection.Unprotected,
                vbeBuilder,
                ("otherComponent", ComponentType.ClassModule, string.Empty));
            vbeBuilder.AddProject(otherProjectMock);

            var vbe = vbeBuilder.Build().Object;
            var project = projectMock.Object;
            var otherProject = otherProjectMock.Object;
            var component2Mock = mockComponents[1];
            var component2 = component2Mock.Object;

            var repository = TestRepository(vbe);
            project.VBComponents.Remove(component2);
            repository.Refresh(otherProject.ProjectId);
            var returnedComponent = repository.Component(component2.QualifiedModuleName);

            var _ = returnedComponent.VBE;
            component2Mock.Verify(m => m.VBE, Times.Once);
        }

        [Test]
        [Category("COM")]
        public void ComponentReturnsNullForQmnOfRemovedComponentAfterRefresh()
        {
            var vbeBuilder = new MockVbeBuilder();
            var (projectMock, _, mockComponents) = TestProjectAndComponents(
                "project",
                ProjectProtection.Unprotected,
                vbeBuilder,
                ("component1", ComponentType.ClassModule, string.Empty),
                ("component2", ComponentType.ClassModule, string.Empty));
            vbeBuilder.AddProject(projectMock);
            var otherProjectMock = TestProject(
                "otherProject",
                ProjectProtection.Unprotected,
                vbeBuilder,
                ("otherComponent", ComponentType.ClassModule, string.Empty));
            vbeBuilder.AddProject(otherProjectMock);

            var vbe = vbeBuilder.Build().Object;
            var project = projectMock.Object;
            var component2 = mockComponents[1].Object;

            var repository = TestRepository(vbe);
            project.VBComponents.Remove(component2);
            repository.Refresh();
            var returnedComponent = repository.Component(component2.QualifiedModuleName);

            Assert.IsNull(returnedComponent);
        }

        [Test]
        [Category("COM")]
        public void ComponentReturnsNullForQmnOfRemovedComponentAfterRefreshForProjectIdOfContainingProject()
        {
            var vbeBuilder = new MockVbeBuilder();
            var (projectMock, _, mockComponents) = TestProjectAndComponents(
                "project",
                ProjectProtection.Unprotected,
                vbeBuilder,
                ("component1", ComponentType.ClassModule, string.Empty),
                ("component2", ComponentType.ClassModule, string.Empty));
            vbeBuilder.AddProject(projectMock);
            var otherProjectMock = TestProject(
                "otherProject",
                ProjectProtection.Unprotected,
                vbeBuilder,
                ("otherComponent", ComponentType.ClassModule, string.Empty));
            vbeBuilder.AddProject(otherProjectMock);

            var vbe = vbeBuilder.Build().Object;
            var project = projectMock.Object;
            var component2 = mockComponents[1].Object;

            var repository = TestRepository(vbe);
            project.VBComponents.Remove(component2);
            repository.Refresh(project.ProjectId);
            var returnedComponent = repository.Component(component2.QualifiedModuleName);

            Assert.IsNull(returnedComponent);
        }
    }
}