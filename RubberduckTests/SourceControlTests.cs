using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moq;
using Microsoft.Vbe.Interop;
using Rubberduck.SourceControl;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace RubberduckTests
{
    [TestClass]
    public class SourceControlTests
    {

        [TestMethod]
        public void InitVBAProjectIntitializesRepo()
        {
            //arrange
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
            project.Setup(p => p.Name).Returns("RubberduckTestVBAProject");
 
            //act
            var git = new GitProvider(project.Object);
            git.InitVBAProject(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments));

            //assert
            Assert.AreEqual(project.Object.Name, git.CurrentRepository.Name);

            var repoDir = System.IO.Path.Combine(git.CurrentRepository.LocalLocation, ".git");
            Assert.IsTrue(System.IO.Directory.Exists(repoDir), "Repo directory does not exist.");

            //cleanup file system
            var removeDir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "RubberduckTestVBAProject");
            ForceDeleteDirectory(removeDir);
        }

        [TestMethod]
        public void CloneCreatesLocalRepo()
        {
            //arrange
            var project = new Mock<VBProject>();
            var expected = new Repository("SourceControlTest",
                                          Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "SourceControlTest"),
                                          @"https://github.com/ckuhn203/SourceControlTest.git"
                                          );
            var git = new GitProvider(project.Object, expected);

            //act
            var actual = git.Clone(expected.RemoteLocation, expected.LocalLocation);

            //assert
            Assert.AreEqual(expected.Name, actual.Name);
            Assert.AreEqual(expected.LocalLocation, actual.LocalLocation);
            Assert.AreEqual(expected.RemoteLocation, actual.RemoteLocation);
            Assert.IsTrue(Directory.Exists(Path.Combine(expected.LocalLocation, ".git")));

            //cleanup file system
            ForceDeleteDirectory(expected.LocalLocation);
        }

        private void ForceDeleteDirectory(string path)
        {
            var directory = new DirectoryInfo(path) { Attributes = FileAttributes.Normal };

            foreach (var info in directory.GetFileSystemInfos("*", SearchOption.AllDirectories))
            {
                info.Attributes = FileAttributes.Normal;
            }

            directory.Delete(true);
        }
    }
}
