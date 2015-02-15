using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moq;
using Microsoft.Vbe.Interop;
using Rubberduck.SourceControl;
using System.Collections.Generic;

namespace RubberduckTests
{
    [TestClass]
    public class SourceControlTests
    {

        [TestMethod]
        public void InitVBAProjectIntitializesRepo()
        {
            var component = new Mock<VBComponent>();
            component.Setup(c => c.Name).Returns("Module1");
            component.Setup(c => c.Type).Returns(vbext_ComponentType.vbext_ct_StdModule);
            component.Setup(c => c.Export("foo")).Verifiable();

            var componentList = new List<VBComponent>();
            componentList.Add(component.Object);

            var components = new Mock<VBComponents>();
            components.Setup(c => c.GetEnumerator()).Returns(componentList.GetEnumerator());

            var project = new Mock<VBProject>();

            project.Setup(p => p.VBComponents).Returns(components.Object);
            project.Setup(p => p.Name).Returns("TestProject");

            var git = new GitProvider(project.Object);
            git.InitVBAProject(System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)));
            Assert.AreEqual(project.Object.Name, git.CurrentRepository.Name);

            var repoDir = System.IO.Path.Combine(git.CurrentRepository.LocalLocation, ".git");
            Assert.IsTrue(System.IO.Directory.Exists(repoDir), "Repo directory does not exist.");

            //cleanup file system
            System.IO.Directory.Delete(repoDir, true);
        }
    }
}
