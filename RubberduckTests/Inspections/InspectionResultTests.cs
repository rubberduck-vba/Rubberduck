using System.Collections.Generic;
using Moq;
using NUnit.Framework;
using Rubberduck.CodeAnalysis.Inspections;
using Rubberduck.CodeAnalysis.Inspections.Results;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Annotations;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Parsing.VBA.DeclarationCaching;
using Rubberduck.Parsing.VBA.ReferenceManagement;
using Rubberduck.VBEditor;

namespace RubberduckTests.Inspections
{
    [TestFixture]
    public class InspectionResultTests
    {
        [Test]
        public void InspectionResultsAreDeemedInvalidatedIfTheModuleWithTheirQualifiedModuleNameHasBeenModified()
        {
            var inspectionMock = new Mock<IInspection>();
            inspectionMock
                .Setup(m =>
                    m.ChangesInvalidateResult(It.IsAny<IInspectionResult>(),
                        It.IsAny<ICollection<QualifiedModuleName>>()))
                .Returns(false);

            var module = new QualifiedModuleName("project", string.Empty,"module");
            var context = new QualifiedContext(module, null);
            var modifiedModules = new HashSet<QualifiedModuleName>{module};

            var inspectionResult = new QualifiedContextInspectionResult(inspectionMock.Object, string.Empty, context);

            Assert.IsTrue(inspectionResult.ChangesInvalidateResult(modifiedModules));
        }

        [Test]
        public void InspectionResultsAreDeemedInvalidatedIfTheInspectionDeemsThemInvalidated()
        {
            var inspectionMock = new Mock<IInspection>();
            inspectionMock
                .Setup(m =>
                    m.ChangesInvalidateResult(It.IsAny<IInspectionResult>(),
                        It.IsAny<ICollection<QualifiedModuleName>>()))
                .Returns(true);

            var module = new QualifiedModuleName("project", string.Empty, "module");
            var context = new QualifiedContext(module, null);
            var modifiedModules = new HashSet<QualifiedModuleName>();

            var inspectionResult = new QualifiedContextInspectionResult(inspectionMock.Object, string.Empty, context);

            Assert.IsTrue(inspectionResult.ChangesInvalidateResult(modifiedModules));
        }

        [Test]
        public void QualifiedContextInspectionResultsAreNotDeemedInvalidatedIfNeitherTheInspectionDeemsThemInvalidatedNorTheirQualifiedModuleNameGetsModified()
        {
            var inspectionMock = new Mock<IInspection>();
            inspectionMock
                .Setup(m =>
                    m.ChangesInvalidateResult(It.IsAny<IInspectionResult>(),
                        It.IsAny<ICollection<QualifiedModuleName>>()))
                .Returns(false);

            var module = new QualifiedModuleName("project", string.Empty, "module");
            var otherModule = new QualifiedModuleName("project", string.Empty, "otherModule");
            var context = new QualifiedContext(module, null);
            var modifiedModules = new HashSet<QualifiedModuleName>{ otherModule };

            var inspectionResult = new QualifiedContextInspectionResult(inspectionMock.Object, string.Empty, context);

            Assert.IsFalse(inspectionResult.ChangesInvalidateResult(modifiedModules));
        }

        [Test]
        public void DeclarationInspectionResultsAreDeemedInvalidatedIfTheirTargetsModuleGetsModified()
        {
            var inspectionMock = new Mock<IInspection>();
            inspectionMock
                .Setup(m =>
                    m.ChangesInvalidateResult(It.IsAny<IInspectionResult>(),
                        It.IsAny<ICollection<QualifiedModuleName>>()))
                .Returns(false);

            var module = new QualifiedModuleName("project", string.Empty, "module");
            var declarationModule = new QualifiedModuleName("project", string.Empty, "declarationModule");
            var declarationMemberName = new QualifiedMemberName(declarationModule, "test");
            var context = new QualifiedContext(module, null);
            var declaration = new Declaration(declarationMemberName, null, string.Empty, string.Empty, string.Empty, false, false,
                Accessibility.Public, DeclarationType.Constant, null, null, default, false, null);
            var modifiedModules = new HashSet<QualifiedModuleName>{declarationModule};
            
            var inspectionResult = new DeclarationInspectionResult(inspectionMock.Object, string.Empty, declaration, context);

            Assert.IsTrue(inspectionResult.ChangesInvalidateResult(modifiedModules));
        }

        [Test]
        public void DeclarationInspectionResultsAreNotDeemedInvalidatedIfNeitherTheInspectionDeemsThemInvalidatedNorTheirModuleNorThatOfTheTargetGetModified()
        {
            var inspectionMock = new Mock<IInspection>();
            inspectionMock
                .Setup(m =>
                    m.ChangesInvalidateResult(It.IsAny<IInspectionResult>(),
                        It.IsAny<ICollection<QualifiedModuleName>>()))
                .Returns(false);

            var module = new QualifiedModuleName("project", string.Empty, "module");
            var declarationModule = new QualifiedModuleName("project", string.Empty, "declarationModule");
            var otherModule = new QualifiedModuleName("project", string.Empty, "otherModule");
            var declarationMemberName = new QualifiedMemberName(declarationModule, "test");
            var context = new QualifiedContext(module, null);
            var declaration = new Declaration(declarationMemberName, null, string.Empty, string.Empty, string.Empty, false, false,
                Accessibility.Public, DeclarationType.Constant, null, null, default, false, null);
            var modifiedModules = new HashSet<QualifiedModuleName> { otherModule };

            var inspectionResult = new DeclarationInspectionResult(inspectionMock.Object, string.Empty, declaration, context);

            Assert.IsFalse(inspectionResult.ChangesInvalidateResult(modifiedModules));
        }

        [Test]
        public void IdentifierRefereneceInspectionResultsAreDeemedInvalidatedIfTheModuleOfTheirReferencedDeclarationGetsModified()
        {
            var inspectionMock = new Mock<IInspection>();
            inspectionMock
                .Setup(m =>
                    m.ChangesInvalidateResult(It.IsAny<IInspectionResult>(),
                        It.IsAny<ICollection<QualifiedModuleName>>()))
                .Returns(false);

            var module = new QualifiedModuleName("project", string.Empty, "module");
            var declarationModule = new QualifiedModuleName("project", string.Empty, "declarationModule");
            var declarationMemberName = new QualifiedMemberName(declarationModule, "test");
            var declaration = new Declaration(declarationMemberName, null, string.Empty, string.Empty, string.Empty, false, false,
                Accessibility.Public, DeclarationType.Constant, null, null, default, false, null);
            var identifierReference = new IdentifierReference(module, null, null, "test", default, null, declaration);
            var modifiedModules = new HashSet<QualifiedModuleName> { declarationModule };

            var finder = new DeclarationFinder(
                new List<Declaration>(), 
                new List<IParseTreeAnnotation>(), 
                new Dictionary<QualifiedModuleName, LogicalLineStore>(),
                new Dictionary<QualifiedModuleName, IFailedResolutionStore>());
            var inspectionResult = new IdentifierReferenceInspectionResult(inspectionMock.Object, string.Empty, finder, identifierReference);

            Assert.IsTrue(inspectionResult.ChangesInvalidateResult(modifiedModules));
        }

        [Test]
        public void IdentifierReferenceInspectionResultsAreNotDeemedInvalidatedIfNeitherTheInspectionDeemsThemInvalidatedNorTheirModuleNorThatOfTheReferencedDeclarationGetModified()
        {
            var inspectionMock = new Mock<IInspection>();
            inspectionMock
                .Setup(m =>
                    m.ChangesInvalidateResult(It.IsAny<IInspectionResult>(),
                        It.IsAny<ICollection<QualifiedModuleName>>()))
                .Returns(false);

            var module = new QualifiedModuleName("project", string.Empty, "module");
            var declarationModule = new QualifiedModuleName("project", string.Empty, "declarationModule");
            var otherModule = new QualifiedModuleName("project", string.Empty, "otherModule");
            var declarationMemberName = new QualifiedMemberName(declarationModule, "test");
            var declaration = new Declaration(declarationMemberName, null, string.Empty, string.Empty, string.Empty, false, false,
                Accessibility.Public, DeclarationType.Constant, null, null, default, false, null);

            var identifierReference = new IdentifierReference(module, null, null, "test", default, null, declaration); 
            var modifiedModules = new HashSet<QualifiedModuleName> { otherModule };

            var finder = new DeclarationFinder(
                new List<Declaration>(), 
                new List<IParseTreeAnnotation>(), 
                new Dictionary<QualifiedModuleName, LogicalLineStore>(),
                new Dictionary<QualifiedModuleName, IFailedResolutionStore>());
            var inspectionResult = new IdentifierReferenceInspectionResult(inspectionMock.Object, string.Empty, finder, identifierReference);

            Assert.IsFalse(inspectionResult.ChangesInvalidateResult(modifiedModules));
        }
    }
}