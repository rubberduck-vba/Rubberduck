using System;
using System.Collections.Generic;
using System.Linq;
using NUnit.Framework;
using Rubberduck.Parsing.Annotations;
using Rubberduck.Refactorings.AnnotateDeclaration;
using Rubberduck.UI.Refactorings.AnnotateDeclaration;

namespace RubberduckTests.Refactoring.AnnotateDeclaration
{
    [TestFixture]
    public class AnnotationArgumentViewModelTests
    {
        [Test]
        [TestCase(AnnotationArgumentType.Attribute)]
        [TestCase(AnnotationArgumentType.Inspection)]
        [TestCase(AnnotationArgumentType.Boolean)]
        [TestCase(AnnotationArgumentType.Number)]
        [TestCase(AnnotationArgumentType.Text)]
        public void RecognizesSingleArgumentType(AnnotationArgumentType argumentType)
        {
            var viewModel = TestViewModel(argumentType);

            Assert.AreEqual(1,viewModel.ApplicableArgumentTypes.Count);
            Assert.AreEqual(argumentType, viewModel.ApplicableArgumentTypes.First());
        }

        [Test]
        public void SplitsArgumentTypeIntoFlags()
        {
            var viewModel = TestViewModel(AnnotationArgumentType.Boolean | AnnotationArgumentType.Number | AnnotationArgumentType.Text);
            var applicableTypes = viewModel.ApplicableArgumentTypes.ToList();

            Assert.AreEqual(3, applicableTypes.Count);
            Assert.Contains(AnnotationArgumentType.Boolean, applicableTypes);
            Assert.Contains(AnnotationArgumentType.Number, applicableTypes);
            Assert.Contains(AnnotationArgumentType.Text, applicableTypes);
        }

        [Test]
        public void InitiallySelectedArgumentTypeIsFirstApplicableOne()
        {
            var viewModel = TestViewModel(AnnotationArgumentType.Boolean | AnnotationArgumentType.Number | AnnotationArgumentType.Text);

            var expectedType = viewModel.ApplicableArgumentTypes.First();
            var actualType = viewModel.ArgumentType;

            Assert.AreEqual(expectedType, actualType);
        }

        [Test]
        public void CanEditArgumentTypeForMultipleApplicableArgumentTypes()
        {
            var viewModel = TestViewModel(AnnotationArgumentType.Boolean | AnnotationArgumentType.Number | AnnotationArgumentType.Text);

            Assert.IsTrue(viewModel.CanEditArgumentType);
        }

        [Test]
        public void CannotEditArgumentTypeForSingleApplicableArgumentType()
        {
            var viewModel = TestViewModel(AnnotationArgumentType.Boolean);

            Assert.IsFalse(viewModel.CanEditArgumentType);
        }

        [Test]
        [TestCase(AnnotationArgumentType.Attribute)]
        [TestCase(AnnotationArgumentType.Inspection)]
        [TestCase(AnnotationArgumentType.Boolean)]
        [TestCase(AnnotationArgumentType.Number)]
        [TestCase(AnnotationArgumentType.Text)]
        public void EmptyArgumentsAreIllegal(AnnotationArgumentType argumentType)
        {
            var viewModel = TestViewModel(argumentType, initialArgument: "someText");
            viewModel.ArgumentValue = string.Empty;

            Assert.IsTrue(viewModel.HasErrors);
        }

        [Test]
        [TestCase(AnnotationArgumentType.Attribute)]
        [TestCase(AnnotationArgumentType.Inspection)]
        [TestCase(AnnotationArgumentType.Boolean)]
        [TestCase(AnnotationArgumentType.Number)]
        [TestCase(AnnotationArgumentType.Text)]
        public void ArgumentsLongerThan511CharactersAreIllegal(AnnotationArgumentType argumentType)
        {
            var viewModel = TestViewModel(argumentType, initialArgument: "someText");
            viewModel.ArgumentValue = new string('s', 512);

            Assert.IsTrue(viewModel.HasErrors);
        }

        [Test]
        public void TextArgumentsWith511CharactersAreLegal()
        {
            var viewModel = TestViewModel(AnnotationArgumentType.Text, initialArgument: "someText");
            viewModel.ArgumentValue = new string('s', 511);

            Assert.IsFalse(viewModel.HasErrors);
        }

        [Test]
        [TestCase(AnnotationArgumentType.Attribute)]
        [TestCase(AnnotationArgumentType.Inspection)]
        [TestCase(AnnotationArgumentType.Boolean)]
        [TestCase(AnnotationArgumentType.Number)]
        [TestCase(AnnotationArgumentType.Text)]
        public void NewLinesInArgumentsAreIllegal(AnnotationArgumentType argumentType)
        {
            var viewModel = TestViewModel(argumentType, initialArgument: "someText");
            viewModel.ArgumentValue = $"text with{Environment.NewLine}new line";

            Assert.IsTrue(viewModel.HasErrors);
        }

        [Test]
        [TestCase(AnnotationArgumentType.Attribute)]
        [TestCase(AnnotationArgumentType.Inspection)]
        [TestCase(AnnotationArgumentType.Boolean)]
        [TestCase(AnnotationArgumentType.Number)]
        [TestCase(AnnotationArgumentType.Text)]
        public void ControlCharactersInArgumentsAreIllegal(AnnotationArgumentType argumentType)
        {
            var viewModel = TestViewModel(argumentType, initialArgument: "someText");
            viewModel.ArgumentValue = "text with \u0000 control character";

            Assert.IsTrue(viewModel.HasErrors);
        }

        [Test]
        [TestCase(AnnotationArgumentType.Attribute)]
        [TestCase(AnnotationArgumentType.Inspection)]
        [TestCase(AnnotationArgumentType.Boolean)]
        [TestCase(AnnotationArgumentType.Number)]
        [TestCase(AnnotationArgumentType.Text)]
        public void InitialValueIsValidated(AnnotationArgumentType argumentType)
        {
            var viewModel = TestViewModel(argumentType, initialArgument: string.Empty);

            Assert.IsTrue(viewModel.HasErrors);
        }

        [Test]
        [TestCase(AnnotationArgumentType.Attribute, "VB_Exposed")]
        [TestCase(AnnotationArgumentType.Inspection, "MyInspection")]
        [TestCase(AnnotationArgumentType.Boolean, "True")]
        [TestCase(AnnotationArgumentType.Number, "42")]
        [TestCase(AnnotationArgumentType.Text, "someText")]
        public void SettingValidArgumentClearsErrors(AnnotationArgumentType argumentType, string validArgument)
        {
            var viewModel = TestViewModel(argumentType, initialArgument: "someText", inspectionNames:new []{"MyInspection"});
            viewModel.ArgumentValue = string.Empty;
            viewModel.ArgumentValue = validArgument;

            Assert.IsFalse(viewModel.HasErrors);
        }

        [Test]
        [TestCase(AnnotationArgumentType.Attribute, AnnotationArgumentType.Number, "VB_Exposed")]
        [TestCase(AnnotationArgumentType.Inspection, AnnotationArgumentType.Attribute, "MyInspection")]
        [TestCase(AnnotationArgumentType.Boolean, AnnotationArgumentType.Inspection, "True")]
        [TestCase(AnnotationArgumentType.Number, AnnotationArgumentType.Boolean, "42")]
        [TestCase(AnnotationArgumentType.Text, AnnotationArgumentType.Attribute, "someText")]
        public void ChangingTheArgumentTypeCausesValidation_ToIllegal(AnnotationArgumentType initialArgumentType, AnnotationArgumentType newIllegalArgumentType, string initiallyLegalValue)
        {
            const AnnotationArgumentType allArgumentTypes = AnnotationArgumentType.Attribute 
                                                            | AnnotationArgumentType.Inspection 
                                                            | AnnotationArgumentType.Boolean 
                                                            | AnnotationArgumentType.Number 
                                                            | AnnotationArgumentType.Text;

            var viewModel = TestViewModel(allArgumentTypes, initialArgument: string.Empty, inspectionNames: new[] { "MyInspection" });
            viewModel.ArgumentType = initialArgumentType;
            viewModel.ArgumentValue = initiallyLegalValue;

            viewModel.ArgumentType = newIllegalArgumentType;

            Assert.IsTrue(viewModel.HasErrors);
        }

        [Test]
        [TestCase(AnnotationArgumentType.Number, AnnotationArgumentType.Attribute, "VB_Exposed")]
        [TestCase(AnnotationArgumentType.Number, AnnotationArgumentType.Inspection, "MyInspection")]
        [TestCase(AnnotationArgumentType.Number, AnnotationArgumentType.Boolean, "True")]
        [TestCase(AnnotationArgumentType.Boolean, AnnotationArgumentType.Number, "42")]
        [TestCase(AnnotationArgumentType.Number, AnnotationArgumentType.Text, "someText")]
        public void ChangingTheArgumentTypeCausesValidation_ToLegal(AnnotationArgumentType initialArgumentType, AnnotationArgumentType newLegalArgumentType, string initialyIllegalValueegalValue)
        {
            const AnnotationArgumentType allArgumentTypes = AnnotationArgumentType.Attribute
                                                            | AnnotationArgumentType.Inspection
                                                            | AnnotationArgumentType.Boolean
                                                            | AnnotationArgumentType.Number
                                                            | AnnotationArgumentType.Text;

            var viewModel = TestViewModel(allArgumentTypes, initialArgument: string.Empty, inspectionNames: new[] { "MyInspection" });
            viewModel.ArgumentType = initialArgumentType;
            viewModel.ArgumentValue = initialyIllegalValueegalValue;

            viewModel.ArgumentType = newLegalArgumentType;

            Assert.IsFalse(viewModel.HasErrors);
        }

        [Test]
        public void ChangingTheArgumentTypeChangesItOnTheReturnedModel()
        {
            var viewModel = TestViewModel(AnnotationArgumentType.Text);
            viewModel.ArgumentType = AnnotationArgumentType.Boolean;

            Assert.AreEqual(AnnotationArgumentType.Boolean, viewModel.Model.ArgumentType);
        }

        [Test]
        public void ChangingTheArgumentValueChangesItOnTheReturnedModel()
        {
            var viewModel = TestViewModel(AnnotationArgumentType.Text, string.Empty);
            viewModel.ArgumentValue = "some Text";

            Assert.AreEqual("some Text", viewModel.Model.Argument);
        }

        private AnnotationArgumentViewModel TestViewModel(
            AnnotationArgumentType argumentType,
            string initialArgument = null, 
            IReadOnlyList<string> inspectionNames = null)
        {
            var model = new TypedAnnotationArgument(argumentType, initialArgument ?? string.Empty);
            return new AnnotationArgumentViewModel(model, inspectionNames ?? new List<string>());
        }
    }
}