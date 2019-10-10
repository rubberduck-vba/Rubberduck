using NUnit.Framework;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Refactorings.Common;

namespace RubberduckTests
{
    [TestFixture]
    public class VBAIdentifierValidatorTests
    {
        [TestCase("1NoCanDo", DeclarationType.Variable, false)] //Does not starts with a letter
        [TestCase("!NoCanDo", DeclarationType.Variable, false)] //Does not starts with a letter
        [TestCase("No@CanDo", DeclarationType.Variable, false)] //Uses a special character
        [TestCase("Yes_CanDo", DeclarationType.Variable, true)] //Uses an underscore special character
        [TestCase("CStr", DeclarationType.Variable, false)] //ReservedKeyword not OK for variables
        [TestCase("VBA", DeclarationType.Variable, true)] //VBA OK for variables
        [TestCase("VBA", DeclarationType.Module, true)] //VBA OK for modules
        [TestCase("VBA", DeclarationType.Project, false)] //VBA Not OK for projects
        [TestCase("O123456789O123456789O123456789O", DeclarationType.Module, true)] //31 chars OK for modules
        [TestCase("O123456789O123456789O123456789O1", DeclarationType.Module, false)] //32 chars Not OK for modules
        [TestCase("O123456789O123456789O123456789O1", DeclarationType.Variable, true)] //32 chars OK for variables
        [TestCase("O123456789O123456789O123456789O1", DeclarationType.Project, true)] //32 chars OK for projects
        [Category("Rename")]
        public void VBAIdentifierValidator_IsValidName(string identifier, DeclarationType declarationType, bool expected)
        {
            Assert.AreEqual(expected, VBAIdentifierValidator.IsValidIdentifier(identifier, declarationType));
        }

        [TestCase("val1", false)] //ends with number
        [TestCase("abc", true)] //OK
        [TestCase("b1", false)]  //too short
        [TestCase("bbbbbbb", false)] //repeated letter
        [Category("Rename")]
        public void VBAIdentifierValidator_IsMeaningfulName(string identifier, bool expected)
        {
            Assert.AreEqual(expected, VBAIdentifierValidator.IsMeaningfulIdentifier(identifier));
        }

    }
}
