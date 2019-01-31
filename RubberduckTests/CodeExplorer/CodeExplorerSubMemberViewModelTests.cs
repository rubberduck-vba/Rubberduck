using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NUnit.Framework;
using Rubberduck.Navigation.CodeExplorer;
using Rubberduck.Parsing.Symbols;

namespace RubberduckTests.CodeExplorer
{
    [TestFixture]
    [SuppressMessage("ReSharper", "UnusedVariable")]
    public class CodeExplorerSubMemberViewModelTests
    {
        [Test]
        [Category("Code Explorer")]
        [TestCase(CodeExplorerTestCode.TestTypeMemberName)]
        [TestCase(CodeExplorerTestCode.TestEnumMemberName)]
        public void Constructor_SetsDeclaration(string name)
        {
            var declarations = CodeExplorerTestSetup.TestProjectOneDeclarations.TestSubMemberDeclarations(name, out var subMemberDeclaration);
            var subMember = new CodeExplorerSubMemberViewModel(null, subMemberDeclaration);

            Assert.AreSame(subMemberDeclaration, subMember.Declaration);
        }

        [Test]
        [Category("Code Explorer")]
        [TestCase(CodeExplorerTestCode.TestTypeMemberName)]
        [TestCase(CodeExplorerTestCode.TestEnumMemberName)]
        public void Constructor_SetsName(string name)
        {
            var declarations = CodeExplorerTestSetup.TestProjectOneDeclarations.TestSubMemberDeclarations(name, out var subMemberDeclaration);
            var subMember = new CodeExplorerSubMemberViewModel(null, subMemberDeclaration);

            Assert.AreEqual(name, subMember.Name);
        }

        [Test]
        [Category("Code Explorer")]
        [TestCase(CodeExplorerTestCode.TestTypeMemberName)]
        [TestCase(CodeExplorerTestCode.TestEnumMemberName)]
        public void Constructor_NameWithSignatureIsSet(string name)
        {
            var declarations = CodeExplorerTestSetup.TestProjectOneDeclarations.TestSubMemberDeclarations(name, out var subMemberDeclaration);
            var subMember = new CodeExplorerSubMemberViewModel(null, subMemberDeclaration);

            Assert.IsFalse(string.IsNullOrEmpty(subMember.NameWithSignature));
        }

        [Test]
        [Category("Code Explorer")]
        [TestCase(CodeExplorerTestCode.TestTypeMemberName)]
        [TestCase(CodeExplorerTestCode.TestEnumMemberName)]
        public void Constructor_ToolTipIsSet(string name)
        {
            var declarations = CodeExplorerTestSetup.TestProjectOneDeclarations.TestSubMemberDeclarations(name, out var subMemberDeclaration);
            var subMember = new CodeExplorerSubMemberViewModel(null, subMemberDeclaration);

            Assert.IsFalse(string.IsNullOrEmpty(subMember.ToolTip));
        }

        [Test]
        [Category("Code Explorer")]
        [TestCase(CodeExplorerTestCode.TestTypeMemberName)]
        [TestCase(CodeExplorerTestCode.TestEnumMemberName)]
        public void Constructor_SetsIsExpandedFalse(string name)
        {
            var declarations = CodeExplorerTestSetup.TestProjectOneDeclarations.TestSubMemberDeclarations(name, out var subMemberDeclaration);
            var subMember = new CodeExplorerSubMemberViewModel(null, subMemberDeclaration);

            Assert.IsFalse(subMember.IsExpanded);
        }

        [Test]
        [Category("Code Explorer")]
        [TestCase(CodeExplorerSortOrder.Undefined, typeof(CompareByCodeLine))]
        [TestCase(CodeExplorerSortOrder.Name, typeof(CompareByName))]
        [TestCase(CodeExplorerSortOrder.CodeLine, typeof(CompareByCodeLine))]
        [TestCase(CodeExplorerSortOrder.DeclarationType, typeof(CompareByCodeLine))]
        [TestCase(CodeExplorerSortOrder.DeclarationTypeThenName, typeof(CompareByName))]
        [TestCase(CodeExplorerSortOrder.DeclarationTypeThenCodeLine, typeof(CompareByCodeLine))]
        public void SortComparerIsCorrectSortOrderType(CodeExplorerSortOrder order, Type comparerType)
        {
            var declarations =
                CodeExplorerTestSetup.TestProjectOneDeclarations.TestSubMemberDeclarations(
                    CodeExplorerTestCode.TestTypeMemberName, out var subMemberDeclaration);

            var subMember = new CodeExplorerSubMemberViewModel(null, subMemberDeclaration) { SortOrder = order };

            Assert.AreEqual(comparerType, subMember.SortComparer.GetType());
        }

        [Test]
        [Category("Code Explorer")]
        [TestCase(CodeExplorerTestCode.TestTypeMemberName)]
        [TestCase(CodeExplorerTestCode.TestEnumMemberName)]
        public void FilteredIsFalseForSubsetsOfName(string name)
        {
            var declarations = CodeExplorerTestSetup.TestProjectOneDeclarations.TestSubMemberDeclarations(name, out var subMemberDeclaration);
            var subMember = new CodeExplorerSubMemberViewModel(null, subMemberDeclaration);

            for (var characters = 1; characters <= name.Length; characters++)
            {
                subMember.Filter = name.Substring(0, characters);
                Assert.IsFalse(subMember.Filtered);
            }

            for (var position = name.Length - 2; position > 0; position--)
            {
                subMember.Filter = name.Substring(position);
                Assert.IsFalse(subMember.Filtered);
            }
        }

        [Test]
        [Category("Code Explorer")]
        [TestCase(CodeExplorerTestCode.TestTypeMemberName)]
        [TestCase(CodeExplorerTestCode.TestEnumMemberName)]
        public void FilteredIsTrueForCharactersNotInName(string name)
        {
            const string testCharacters = "abcdefghijklmnopqrstuwxyz";

            var declarations = CodeExplorerTestSetup.TestProjectOneDeclarations.TestSubMemberDeclarations(name, out var subMemberDeclaration);
            var subMember = new CodeExplorerSubMemberViewModel(null, subMemberDeclaration);

            var nonMatching = testCharacters.ToCharArray().Except(name.ToLowerInvariant().ToCharArray());

            foreach (var character in nonMatching.Select(letter => letter.ToString()))
            {
                subMember.Filter = character;
                Assert.IsTrue(subMember.Filtered);
            }
        }

        [Test]
        [Category("Code Explorer")]
        [TestCase(CodeExplorerTestCode.TestTypeMemberName)]
        [TestCase(CodeExplorerTestCode.TestEnumMemberName)]
        public void Synchronize_UsesNewDeclaration(string name)
        {
            var declarations = CodeExplorerTestSetup.TestProjectOneDeclarations.TestSubMemberDeclarations(name, out var subMemberDeclaration);
            var subMember = new CodeExplorerSubMemberViewModel(null, subMemberDeclaration);

            var updated = subMemberDeclaration.ShallowCopy();
            var updates = new List<Declaration> { updated };

            subMember.Synchronize(ref updates);

            Assert.AreSame(updated, subMember.Declaration);
        }

        [Test]
        [Category("Code Explorer")]
        [TestCase(CodeExplorerTestCode.TestTypeMemberName)]
        [TestCase(CodeExplorerTestCode.TestEnumMemberName)]
        public void Synchronize_ClearsDeclaration_EmptyList(string name)
        {
            var declarations = CodeExplorerTestSetup.TestProjectOneDeclarations.TestSubMemberDeclarations(name, out var subMemberDeclaration);
            var subMember = new CodeExplorerSubMemberViewModel(null, subMemberDeclaration);
            var updates = new List<Declaration>();

            subMember.Synchronize(ref updates);

            Assert.IsNull(subMember.Declaration);
        }

        [Test]
        [Category("Code Explorer")]
        [TestCase(CodeExplorerTestCode.TestTypeMemberName)]
        [TestCase(CodeExplorerTestCode.TestEnumMemberName)]
        public void Synchronize_ClearsDeclaration_NotInList(string name)
        {
            var declarations = CodeExplorerTestSetup.TestProjectOneDeclarations.TestSubMemberDeclarations(name, out var subMemberDeclaration);
            var subMember = new CodeExplorerSubMemberViewModel(null, subMemberDeclaration);

            var updates = CodeExplorerTestSetup.TestProjectOneDeclarations;
            updates.Remove(subMemberDeclaration);

            subMember.Synchronize(ref updates);

            Assert.IsNull(subMember.Declaration);
        }
    }
}
