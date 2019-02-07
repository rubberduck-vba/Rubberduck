using System.Linq;
using System.Threading;
using Moq;
using NUnit.Framework;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.Utility;

namespace RubberduckTests.Rewriter
{
    [TestFixture]
    public class SelectionRecovererTests
    {
        [Test]
        [Category("Rewriting")]
        public void SetsExactlySavedSelectionsOnRecoverSelections()
        {
            var selectionServiceMock = TestSelectionServiceMock();
            var parseManager = new Mock<IParseManager>().Object;
            var selectionRecoverer = new SelectionRecoverer(selectionServiceMock.Object, parseManager);

            selectionRecoverer.SaveSelections(_testModuleSelections.Select(qualifiedSelection => qualifiedSelection.QualifiedName).Take(2));
            selectionRecoverer.RecoverSavedSelections();

            foreach (var qualifiedSelection in _testModuleSelections.Take(2))
            {
                selectionServiceMock.Verify(m => m.TrySetSelection(qualifiedSelection.QualifiedName, qualifiedSelection.Selection), Times.Once);
            }
            selectionServiceMock.Verify(m => m.TrySetSelection(_testModuleSelections[2].QualifiedName, It.IsAny<Selection>()), Times.Never);
        }

        [Test]
        [Category("Rewriting")]
        public void SetsExactlyLastSavedSelectionsOnRecoverSelectionsAfterMultipleSaves()
        {
            var selectionServiceMock = TestSelectionServiceMock();
            var parseManager = new Mock<IParseManager>().Object;
            var selectionRecoverer = new SelectionRecoverer(selectionServiceMock.Object, parseManager);

            selectionRecoverer.SaveSelections(_testModuleSelections.Select(qualifiedSelection => qualifiedSelection.QualifiedName).Take(2));
            selectionRecoverer.SaveSelections(_testModuleSelections.Select(qualifiedSelection => qualifiedSelection.QualifiedName).Skip(1));
            selectionRecoverer.RecoverSavedSelections();

            foreach (var qualifiedSelection in _testModuleSelections.Skip(1))
            {
                selectionServiceMock.Verify(m => m.TrySetSelection(qualifiedSelection.QualifiedName, qualifiedSelection.Selection), Times.Once);
            }
            selectionServiceMock.Verify(m => m.TrySetSelection(_testModuleSelections[0].QualifiedName, It.IsAny<Selection>()), Times.Never);
        }

        [Test]
        [Category("Rewriting")]
        public void SetsModifiedSelectionAfterOffsetIsAppliedOnRecoverSelections()
        {
            var selectionServiceMock = TestSelectionServiceMock();
            var parseManager = new Mock<IParseManager>().Object;
            var selectionRecoverer = new SelectionRecoverer(selectionServiceMock.Object, parseManager);

            var selectionOffset = new Selection(0,2,4,5);
            
            selectionRecoverer.SaveSelections(_testModuleSelections.Select(qualifiedSelection => qualifiedSelection.QualifiedName).Take(2));
            selectionRecoverer.AdjustSavedSelection(_testModuleSelections[0].QualifiedName, selectionOffset);
            selectionRecoverer.RecoverSavedSelections();

            var expectedAdjustedSelection = _testModuleSelections[0].Selection.Offset(selectionOffset);

            selectionServiceMock.Verify(m => m.TrySetSelection(_testModuleSelections[0].QualifiedName, expectedAdjustedSelection), Times.Once);
            selectionServiceMock.Verify(m => m.TrySetSelection(_testModuleSelections[1].QualifiedName, _testModuleSelections[1].Selection), Times.Once);
            selectionServiceMock.Verify(m => m.TrySetSelection(_testModuleSelections[2].QualifiedName, It.IsAny<Selection>()), Times.Never);
        }

        [Test]
        [Category("Rewriting")]
        public void SetsReplacementSelectionOnRecoverSelections_SelectionSavedPreviously()
        {
            var selectionServiceMock = TestSelectionServiceMock();
            var parseManager = new Mock<IParseManager>().Object;
            var selectionRecoverer = new SelectionRecoverer(selectionServiceMock.Object, parseManager);

            var selectionReplacement = new Selection(22, 2, 44, 5);

            selectionRecoverer.SaveSelections(_testModuleSelections.Select(qualifiedSelection => qualifiedSelection.QualifiedName).Take(2));
            selectionRecoverer.ReplaceSavedSelection(_testModuleSelections[2].QualifiedName, selectionReplacement);
            selectionRecoverer.RecoverSavedSelections();

            foreach (var qualifiedSelection in _testModuleSelections.Take(2))
            {
                selectionServiceMock.Verify(m => m.TrySetSelection(qualifiedSelection.QualifiedName, qualifiedSelection.Selection), Times.Once);
            }
            selectionServiceMock.Verify(m => m.TrySetSelection(_testModuleSelections[2].QualifiedName, selectionReplacement), Times.Once);
        }

        [Test]
        [Category("Rewriting")]
        public void SetsReplacementSelectionOnRecoverSelections_SelectionNotSavedPreviously()
        {
            var selectionServiceMock = TestSelectionServiceMock();
            var parseManager = new Mock<IParseManager>().Object;
            var selectionRecoverer = new SelectionRecoverer(selectionServiceMock.Object, parseManager);

            var selectionReplacement = new Selection(22, 2, 44, 5);

            selectionRecoverer.SaveSelections(_testModuleSelections.Select(qualifiedSelection => qualifiedSelection.QualifiedName).Take(2));
            selectionRecoverer.ReplaceSavedSelection(_testModuleSelections[0].QualifiedName, selectionReplacement);
            selectionRecoverer.RecoverSavedSelections();

            selectionServiceMock.Verify(m => m.TrySetSelection(_testModuleSelections[0].QualifiedName, selectionReplacement), Times.Once);
            selectionServiceMock.Verify(m => m.TrySetSelection(_testModuleSelections[1].QualifiedName, _testModuleSelections[1].Selection), Times.Once);
            selectionServiceMock.Verify(m => m.TrySetSelection(_testModuleSelections[2].QualifiedName, It.IsAny<Selection>()), Times.Never);
        }

        [Test]
        [Category("Rewriting")]
        public void ReplacementSelectionOverwritesAdjustmentOnRecoverSelections()
        {
            var selectionServiceMock = TestSelectionServiceMock();
            var parseManager = new Mock<IParseManager>().Object;
            var selectionRecoverer = new SelectionRecoverer(selectionServiceMock.Object, parseManager);

            var selectionOffset = new Selection(0, 2, 4, 5);
            var selectionReplacement = new Selection(22, 2, 44, 5);

            selectionRecoverer.SaveSelections(_testModuleSelections.Select(qualifiedSelection => qualifiedSelection.QualifiedName).Take(1));
            selectionRecoverer.AdjustSavedSelection(_testModuleSelections[0].QualifiedName, selectionOffset);
            selectionRecoverer.ReplaceSavedSelection(_testModuleSelections[0].QualifiedName, selectionReplacement);
            selectionRecoverer.RecoverSavedSelections();

            selectionServiceMock.Verify(m => m.TrySetSelection(_testModuleSelections[0].QualifiedName, selectionReplacement), Times.Once);
        }

        [Test]
        [Category("Rewriting")]
        public void SelectionAdjustmentAddsToReplacementSelectionOnRecoverSelections()
        {
            var selectionServiceMock = TestSelectionServiceMock();
            var parseManager = new Mock<IParseManager>().Object;
            var selectionRecoverer = new SelectionRecoverer(selectionServiceMock.Object, parseManager);

            var selectionOffset = new Selection(0, 2, 4, 5);
            var selectionReplacement = new Selection(22, 2, 44, 5);

            selectionRecoverer.SaveSelections(_testModuleSelections.Select(qualifiedSelection => qualifiedSelection.QualifiedName).Take(1));
            selectionRecoverer.ReplaceSavedSelection(_testModuleSelections[0].QualifiedName, selectionReplacement);
            selectionRecoverer.AdjustSavedSelection(_testModuleSelections[0].QualifiedName, selectionOffset);
            selectionRecoverer.RecoverSavedSelections();

            var expectedAdjustedSelection = selectionReplacement.Offset(selectionOffset);

            selectionServiceMock.Verify(m => m.TrySetSelection(_testModuleSelections[0].QualifiedName, expectedAdjustedSelection), Times.Once);
        }

        [Test]
        [Category("Rewriting")]
        public void RecoverSelectionsOnNextParseDoesNotSetAnythingImmediately()
        {
            var selectionServiceMock = TestSelectionServiceMock();
            var parseManager = new Mock<IParseManager>().Object;
            var selectionRecoverer = new SelectionRecoverer(selectionServiceMock.Object, parseManager);

            var selectionOffset = new Selection(0, 2, 4, 5);
            var selectionReplacement = new Selection(22, 2, 44, 5);

            selectionRecoverer.SaveSelections(_testModuleSelections.Select(qualifiedSelection => qualifiedSelection.QualifiedName).Take(2));
            selectionRecoverer.ReplaceSavedSelection(_testModuleSelections[2].QualifiedName, selectionReplacement);
            selectionRecoverer.AdjustSavedSelection(_testModuleSelections[0].QualifiedName, selectionOffset);
            selectionRecoverer.RecoverSavedSelectionsOnNextParse();

            foreach (var qualifiedSelection in _testModuleSelections)
            {
                selectionServiceMock.Verify(m => m.TrySetSelection(qualifiedSelection.QualifiedName, qualifiedSelection.Selection), Times.Never);
            }
        }


        [Test]
        [Category("Rewriting")]
        public void SetsExactlySavedSelectionsOnParseAfterRecoverSelectionsOnNextParse()
        {
            var selectionServiceMock = TestSelectionServiceMock();
            var parseManagerMock = new Mock<IParseManager>();
            var selectionRecoverer = new SelectionRecoverer(selectionServiceMock.Object, parseManagerMock.Object);

            selectionRecoverer.SaveSelections(_testModuleSelections.Select(qualifiedSelection => qualifiedSelection.QualifiedName).Take(2));
            selectionRecoverer.RecoverSavedSelectionsOnNextParse();

            var stateEventArgs = new ParserStateEventArgs(_stateExpectedToTriggerTheRecovery, ParserState.Pending, CancellationToken.None);
            parseManagerMock.Raise(m => m.StateChanged += null, stateEventArgs);

            foreach (var qualifiedSelection in _testModuleSelections.Take(2))
            {
                selectionServiceMock.Verify(m => m.TrySetSelection(qualifiedSelection.QualifiedName, qualifiedSelection.Selection), Times.Once);
            }
            selectionServiceMock.Verify(m => m.TrySetSelection(_testModuleSelections[2].QualifiedName, It.IsAny<Selection>()), Times.Never);
        }

        [Test]
        [Category("Rewriting")]
        public void SetsExactlyLastSavedSelectionsOnParseAfterRecoverSelectionsOnNextParseAfterMultipleSaves()
        {
            var selectionServiceMock = TestSelectionServiceMock();
            var parseManagerMock = new Mock<IParseManager>();
            var selectionRecoverer = new SelectionRecoverer(selectionServiceMock.Object, parseManagerMock.Object);

            selectionRecoverer.SaveSelections(_testModuleSelections.Select(qualifiedSelection => qualifiedSelection.QualifiedName).Take(2));
            selectionRecoverer.SaveSelections(_testModuleSelections.Select(qualifiedSelection => qualifiedSelection.QualifiedName).Skip(1));
            selectionRecoverer.RecoverSavedSelectionsOnNextParse();

            var stateEventArgs = new ParserStateEventArgs(_stateExpectedToTriggerTheRecovery, ParserState.Pending, CancellationToken.None);
            parseManagerMock.Raise(m => m.StateChanged += null, stateEventArgs);

            foreach (var qualifiedSelection in _testModuleSelections.Skip(1))
            {
                selectionServiceMock.Verify(m => m.TrySetSelection(qualifiedSelection.QualifiedName, qualifiedSelection.Selection), Times.Once);
            }
            selectionServiceMock.Verify(m => m.TrySetSelection(_testModuleSelections[0].QualifiedName, It.IsAny<Selection>()), Times.Never);
        }

        [Test]
        [Category("Rewriting")]
        public void SetsModifiedSelectionAfterOffsetIsAppliedOnParseAfterRecoverSelectionsOnNextParse()
        {
            var selectionServiceMock = TestSelectionServiceMock();
            var parseManagerMock = new Mock<IParseManager>();
            var selectionRecoverer = new SelectionRecoverer(selectionServiceMock.Object, parseManagerMock.Object);

            var selectionOffset = new Selection(0, 2, 4, 5);

            selectionRecoverer.SaveSelections(_testModuleSelections.Select(qualifiedSelection => qualifiedSelection.QualifiedName).Take(2));
            selectionRecoverer.AdjustSavedSelection(_testModuleSelections[0].QualifiedName, selectionOffset);
            selectionRecoverer.RecoverSavedSelectionsOnNextParse();

            var stateEventArgs = new ParserStateEventArgs(_stateExpectedToTriggerTheRecovery, ParserState.Pending, CancellationToken.None);
            parseManagerMock.Raise(m => m.StateChanged += null, stateEventArgs);

            var expectedAdjustedSelection = _testModuleSelections[0].Selection.Offset(selectionOffset);

            selectionServiceMock.Verify(m => m.TrySetSelection(_testModuleSelections[0].QualifiedName, expectedAdjustedSelection), Times.Once);
            selectionServiceMock.Verify(m => m.TrySetSelection(_testModuleSelections[1].QualifiedName, _testModuleSelections[1].Selection), Times.Once);
            selectionServiceMock.Verify(m => m.TrySetSelection(_testModuleSelections[2].QualifiedName, It.IsAny<Selection>()), Times.Never);
        }

        [Test]
        [Category("Rewriting")]
        public void SetsReplacementSelectionOnParseAfterRecoverSelectionsOnNextParse_SelectionSavedPreviously()
        {
            var selectionServiceMock = TestSelectionServiceMock();
            var parseManagerMock = new Mock<IParseManager>();
            var selectionRecoverer = new SelectionRecoverer(selectionServiceMock.Object, parseManagerMock.Object);

            var selectionReplacement = new Selection(22, 2, 44, 5);

            selectionRecoverer.SaveSelections(_testModuleSelections.Select(qualifiedSelection => qualifiedSelection.QualifiedName).Take(2));
            selectionRecoverer.ReplaceSavedSelection(_testModuleSelections[2].QualifiedName, selectionReplacement);
            selectionRecoverer.RecoverSavedSelections();

            foreach (var qualifiedSelection in _testModuleSelections.Take(2))
            {
                selectionServiceMock.Verify(m => m.TrySetSelection(qualifiedSelection.QualifiedName, qualifiedSelection.Selection), Times.Once);
            }
            selectionServiceMock.Verify(m => m.TrySetSelection(_testModuleSelections[2].QualifiedName, selectionReplacement), Times.Once);
        }

        [Test]
        [Category("Rewriting")]
        public void SetsReplacementSelectionOnParseAfterRecoverSelectionsOnNextParse_SelectionNotSavedPreviously()
        {
            var selectionServiceMock = TestSelectionServiceMock();
            var parseManagerMock = new Mock<IParseManager>();
            var selectionRecoverer = new SelectionRecoverer(selectionServiceMock.Object, parseManagerMock.Object);

            var selectionReplacement = new Selection(22, 2, 44, 5);

            selectionRecoverer.SaveSelections(_testModuleSelections.Select(qualifiedSelection => qualifiedSelection.QualifiedName).Take(2));
            selectionRecoverer.ReplaceSavedSelection(_testModuleSelections[0].QualifiedName, selectionReplacement);
            selectionRecoverer.RecoverSavedSelectionsOnNextParse();

            var stateEventArgs = new ParserStateEventArgs(_stateExpectedToTriggerTheRecovery, ParserState.Pending, CancellationToken.None);
            parseManagerMock.Raise(m => m.StateChanged += null, stateEventArgs);

            selectionServiceMock.Verify(m => m.TrySetSelection(_testModuleSelections[0].QualifiedName, selectionReplacement), Times.Once);
            selectionServiceMock.Verify(m => m.TrySetSelection(_testModuleSelections[1].QualifiedName, _testModuleSelections[1].Selection), Times.Once);
            selectionServiceMock.Verify(m => m.TrySetSelection(_testModuleSelections[2].QualifiedName, It.IsAny<Selection>()), Times.Never);
        }

        [Test]
        [Category("Rewriting")]
        public void ReplacementSelectionOverwritesAdjustmentOnParseAfterRecoverSelectionsOnNextParse()
        {
            var selectionServiceMock = TestSelectionServiceMock();
            var parseManagerMock = new Mock<IParseManager>();
            var selectionRecoverer = new SelectionRecoverer(selectionServiceMock.Object, parseManagerMock.Object);

            var selectionOffset = new Selection(0, 2, 4, 5);
            var selectionReplacement = new Selection(22, 2, 44, 5);

            selectionRecoverer.SaveSelections(_testModuleSelections.Select(qualifiedSelection => qualifiedSelection.QualifiedName).Take(1));
            selectionRecoverer.AdjustSavedSelection(_testModuleSelections[0].QualifiedName, selectionOffset);
            selectionRecoverer.ReplaceSavedSelection(_testModuleSelections[0].QualifiedName, selectionReplacement);
            selectionRecoverer.RecoverSavedSelectionsOnNextParse();

            var stateEventArgs = new ParserStateEventArgs(_stateExpectedToTriggerTheRecovery, ParserState.Pending, CancellationToken.None);
            parseManagerMock.Raise(m => m.StateChanged += null, stateEventArgs);

            selectionServiceMock.Verify(m => m.TrySetSelection(_testModuleSelections[0].QualifiedName, selectionReplacement), Times.Once);
        }

        [Test]
        [Category("Rewriting")]
        public void SelectionAdjustmentAddsToReplacementSelectionOnParseAfterRecoverSelectionsOnNextParse()
        {
            var selectionServiceMock = TestSelectionServiceMock();
            var parseManagerMock = new Mock<IParseManager>();
            var selectionRecoverer = new SelectionRecoverer(selectionServiceMock.Object, parseManagerMock.Object);

            var selectionOffset = new Selection(0, 2, 4, 5);
            var selectionReplacement = new Selection(22, 2, 44, 5);

            selectionRecoverer.SaveSelections(_testModuleSelections.Select(qualifiedSelection => qualifiedSelection.QualifiedName).Take(1));
            selectionRecoverer.ReplaceSavedSelection(_testModuleSelections[0].QualifiedName, selectionReplacement);
            selectionRecoverer.AdjustSavedSelection(_testModuleSelections[0].QualifiedName, selectionOffset);
            selectionRecoverer.RecoverSavedSelectionsOnNextParse();

            var stateEventArgs = new ParserStateEventArgs(_stateExpectedToTriggerTheRecovery, ParserState.Pending, CancellationToken.None);
            parseManagerMock.Raise(m => m.StateChanged += null, stateEventArgs);

            var expectedAdjustedSelection = selectionReplacement.Offset(selectionOffset);

            selectionServiceMock.Verify(m => m.TrySetSelection(_testModuleSelections[0].QualifiedName, expectedAdjustedSelection), Times.Once);
        }


        private Mock<ISelectionService> TestSelectionServiceMock()
        {
            var mock = new Mock<ISelectionService>();
            foreach(var qualifiedSelection in _testModuleSelections)
            {
                mock.Setup(m => m.Selection(qualifiedSelection.QualifiedName)).Returns(qualifiedSelection.Selection);
            }

            mock.Setup(m => m.TrySetSelection(It.IsAny<QualifiedModuleName>(), It.IsAny<Selection>()));
            return mock;
        }

        private ParserState _stateExpectedToTriggerTheRecovery = ParserState.LoadingReference;

        private readonly QualifiedSelection[] _testModuleSelections = new[]
        {
            new QualifiedSelection(new QualifiedModuleName("testProject", string.Empty, "module1"),
                new Selection(2, 4, 5, 6)),
            new QualifiedSelection(new QualifiedModuleName("testProject", string.Empty, "module2"),
                new Selection(3, 9)),
            new QualifiedSelection(new QualifiedModuleName("testProject", string.Empty, "module3"),
                new Selection(15, 17, 1, 4))
        };
    }
}