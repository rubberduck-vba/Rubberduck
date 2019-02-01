using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using Moq;
using NUnit.Framework;
using Rubberduck.Templates;

namespace RubberduckTests.Templates
{
    [TestFixture]
    [Category("Templates.Resources")]
    public class TemplateResourceTests
    {
        /// <summary>
        /// The test is primarily a guard against accidental omission/typos in the the
        /// template resource, which must conform to specific format/content
        /// </summary>
        [Test]
        public void VerifyTemplateResourceIsComplete()
        {
            var namesCount = 0;
            var captionsCount = 0;
            var descriptionsCount = 0;
            var codesCount = 0;
            var dictionary = new Dictionary<string, int>();
            var resourceManager = Rubberduck.Resources.Templates.ResourceManager;
            var set = resourceManager.GetResourceSet(CultureInfo.InvariantCulture, true, true);
            foreach (DictionaryEntry entry in set)
            {
                var key = (string) entry.Key;
                var lhs = key.Split('_')[0];
                if (!dictionary.ContainsKey(lhs))
                {
                    dictionary.Add(lhs, 1);
                }
                else
                {
                    dictionary[lhs]++;
                }

                if (key.EndsWith("_Name"))
                {
                    namesCount++;
                }

                if (key.EndsWith("_Caption"))
                {
                    captionsCount++;
                }

                if (key.EndsWith("_Description"))
                {
                    descriptionsCount++;
                }

                if (key.EndsWith("_Code"))
                {
                    codesCount++;
                }
            }

            Assert.IsTrue(dictionary.All(e => e.Value == 4),
                $"Those entries are missing required fields in the template resources: {dictionary.Where(e => e.Value != 4).Select(e => e.Key).ToList()}");
            Assert.AreEqual(namesCount, captionsCount,
                "Caption count did not equal to name count. There must be equal numbers of each type in the resource for {{Name, Caption, Description, Code}}");
            Assert.AreEqual(namesCount, descriptionsCount,
                "Description count did not equal to name count. There must be equal numbers of each type in the resource for {{Name, Caption, Description, Code}}");
            Assert.AreEqual(namesCount, codesCount,
                "Code count did not equal to name count. There must be equal numbers of each type in the resource for {{Name, Caption, Description, Code}}");
        }

        [Test]
        public void BuiltInTemplate_ReturnsFalse_UserDefined()
        {
            var handler = new Mock<ITemplateFileHandler>();
            var template = new Template(Rubberduck.Resources.Templates.PredeclaredClassModule_Name, handler.Object);
            Assert.IsFalse(template.IsUserDefined, "IsUserDefined should be false");
        }

        [Test]
        public void Template_ReturnsCorrectName()
        {
            var handler = new Mock<ITemplateFileHandler>();
            var template = new Template(Rubberduck.Resources.Templates.PredeclaredClassModule_Name, handler.Object);
            Assert.AreEqual(Rubberduck.Resources.Templates.PredeclaredClassModule_Name, template.Name);
        }

        [Test]
        public void BuiltInTemplate_ReturnsCorrectCaption()
        {
            var handler = new Mock<ITemplateFileHandler>();
            var template = new Template(Rubberduck.Resources.Templates.PredeclaredClassModule_Name, handler.Object);
            Assert.AreEqual(Rubberduck.Resources.Templates.PredeclaredClassModule_Caption, template.Caption);
        }

        [Test]
        public void BuiltInTemplate_ReturnsCorrectDescription()
        {
            var handler = new Mock<ITemplateFileHandler>();
            var template = new Template(Rubberduck.Resources.Templates.PredeclaredClassModule_Name, handler.Object);
            Assert.AreEqual(Rubberduck.Resources.Templates.PredeclaredClassModule_Description, template.Description);
        }

        [Test]
        public void UserDefinedTemplate_ReturnsTrue_UserDefined()
        {
            var handler = new Mock<ITemplateFileHandler>();
            var template = new Template("I am the great froo-froo and you are mud", handler.Object);
            Assert.IsTrue(template.IsUserDefined, "IsUserDefined should be true");
        }

        [Test]
        public void BuiltInTemplate_File_Autocreated()
        {
            var handler = new Mock<ITemplateFileHandler>();
            handler.Setup(h => h.Exists).Returns(false);
            var template = new Template(Rubberduck.Resources.Templates.PredeclaredClassModule_Name, handler.Object);
            handler.Verify(h => h.Exists, Times.Once);
            handler.Verify(h => h.Write(Rubberduck.Resources.Templates.PredeclaredClassModule_Code), Times.Once);
        }

        [Test]
        public void BuiltInTemplate_ExistingFile_NotCreated()
        {
            var handler = new Mock<ITemplateFileHandler>();
            handler.Setup(h => h.Exists).Returns(true);
            var template = new Template(Rubberduck.Resources.Templates.PredeclaredClassModule_Name, handler.Object);
            handler.Verify(h => h.Exists, Times.Once);
            handler.Verify(h => h.Write(It.IsAny<string>()), Times.Never());
        }
    }
}