using NUnit.Framework;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings;
using Rubberduck.Refactorings.EncapsulateField;
using Rubberduck.Refactorings.EncapsulateFieldUseBackingField;
using Rubberduck.UI.Refactorings.EncapsulateField;
using Rubberduck.VBEditor.SafeComWrappers;
using RubberduckTests.Mocks;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RubberduckTests.Refactoring.EncapsulateField
{
    [TestFixture]
    public class EncapsulateFieldViewModelTests
    {
        private EncapsulateFieldTestSupport Support { get; } = new EncapsulateFieldTestSupport();

        [Test]
        [Category("Refactorings")]
        [Category("Encapsulate Field")]
        [Category(nameof(EncapsulateFieldViewModel))]
        public void ReadOnlyCheckboxDisabledIfExternalWriteReferencesExist()
        {
            var codeClass1 =
@"Public fizz As Integer

Sub Foo()
    fizz = 1
End Sub";
            var codeClass2 =
@"Sub Foo()
    Dim theClass As Class1
    Set theClass = new Class1
    theClass.fizz = 0
    Bar theClass.fizz
End Sub

Sub Bar(ByVal v As Integer)
End Sub";

            var vbe = MockVbeBuilder.BuildFromModules(("Class1", codeClass1, ComponentType.ClassModule), ("Class2", codeClass2, ComponentType.ClassModule)).Object;

            using (var state = MockParser.CreateAndParse(vbe))
            {
                var modelFactory = EncapsulateFieldTestSupport.GetResolver(state).Resolve<IEncapsulateFieldModelFactory>();
                var model = modelFactory.Create(state.DeclarationFinder.MatchName("fizz").Single());

                var viewModel = new EncapsulateFieldViewModel(model);

                Assert.IsFalse(viewModel.EnableReadOnlyOption);
            }
        }
    }
}
