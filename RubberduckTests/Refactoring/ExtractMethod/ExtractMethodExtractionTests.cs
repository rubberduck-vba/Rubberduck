//namespace RubberduckTests.Refactoring.ExtractMethod
//{
//    [TestClass]
//    public class ExtractMethodExtractionTests
//    {

//        #region inputCode
//        string inputCode = @"
//Option explicit
//Public Sub CodeWithDeclaration()
//    Dim x as long
//    Dim y as long
//    Dim z as long
//
//    x = 1 + 2
//    DebugPrint x
//    y = x + 1
//    x = 2
//    DebugPrint y
//
//    z = x
//    DebugPrint z
//
//End Sub
//Public Sub DebugPrint(byval g as long)
//End Sub
//
//
//";
//        #endregion

//        [TestClass]
//        public class WhenRemoveSelectionIsCalledWithValidSelection
//        {

//            [TestMethod]
//            [TestCategory("ExtractedMethodRefactoringTests")]
//            public void shouldRemoveLinesFromCodeModuleFromBottomUp()
//            {

//                var notifyCalls = new List<Tuple<int, int>>();
//                var codeModule = new Mock<ICodeModuleWrapper>();
//                codeModule.Setup(cm => cm.DeleteLines(It.IsAny<int>(), It.IsAny<int>()))
//                    .Callback<int, int>((start, count) => notifyCalls.Add(Tuple.Create(start, count)));
//                var selections = new List<Selection>() { new Selection(5, 1, 5, 20), new Selection(10, 1, 12, 20) };
//                var SUT = new ExtractMethodExtraction();

//                //Act
//                SUT.RemoveSelection(codeModule.Object, selections);

//                //Assert
//                Assert.AreEqual(Tuple.Create(5, 1), notifyCalls[1]);
//                Assert.AreEqual(Tuple.Create(10, 3), notifyCalls[0]);
//            }

//        }

//        [TestClass]
//        public class WhenApplyIsCalled
//        {

//            [TestMethod]
//            [TestCategory("ExtractedMethodRefactoringTests")]
//            public void shouldExtractTheTextForTheNewProcByCallingConstructLinesOfProc()
//            {
//                var newProc = @"
//Public Sub NewMethod()
//    DebugPrint ""a""
//End Sub";
//                var extraction = new Mock<ExtractMethodExtraction>() { CallBase = true };
//                IExtractMethodExtraction SUT = extraction.Object;
//                var codeModule = new Mock<ICodeModuleWrapper>();
//                var model = new Mock<IExtractMethodModel>();
//                var selection = new Selection(1, 1, 1, 1);
//                var methodMock = new Mock<ExtractedMethod>() { CallBase = true };
//                var method = methodMock.Object;
//                method.Accessibility = Accessibility.Private;
//                method.Parameters = new List<ExtractedParameter>();
//                method.MethodName = "NewMethod";
//                methodMock.Setup(m => m.NewMethodCall()).Returns("theMethodCall");
//                model.Setup(m => m.PositionForNewMethod).Returns(selection);
//                model.Setup(m => m.Method).Returns(method);
//                extraction.Setup(em => em.ConstructLinesOfProc(It.IsAny<ICodeModuleWrapper>(), It.IsAny<IExtractMethodModel>())).Returns(newProc);

//                SUT.Apply(codeModule.Object, model.Object, selection);

//                extraction.Verify(extr => extr.ConstructLinesOfProc(codeModule.Object, model.Object));


//            }

//            [TestMethod]
//            [TestCategory("ExtractedMethodRefactoringTests")]
//            public void shouldRemoveSelection()
//            {

//                var newProc = @"
//Public Sub NewMethod()
//    DebugPrint ""a""
//End Sub";
//                var extraction = new Mock<ExtractMethodExtraction>() { CallBase = true };
//                IExtractMethodExtraction SUT = extraction.Object;
//                var codeModule = new Mock<ICodeModuleWrapper>();
//                var selection = new Selection(1, 1, 1, 1);
//                var selections = new List<Selection>() { new Selection(5, 1, 5, 20), new Selection(10, 1, 12, 20) };
//                var methodMock = new Mock<ExtractedMethod>() { CallBase = true };
//                var method = methodMock.Object;
//                method.Accessibility = Accessibility.Private;
//                method.Parameters = new List<ExtractedParameter>();
//                method.MethodName = "NewMethod";
//                methodMock.Setup(m => m.NewMethodCall()).Returns("theMethodCall");
//                var model = new Mock<IExtractMethodModel>();
//                model.Setup(m => m.PositionForNewMethod).Returns(selection);
//                model.Setup(m => m.Method).Returns(method);
//                model.Setup(m => m.RowsToRemove).Returns(selections);

//                extraction.Setup(em => em.ConstructLinesOfProc(It.IsAny<ICodeModuleWrapper>(), It.IsAny<IExtractMethodModel>())).Returns("theMethodCall");

//                SUT.Apply(codeModule.Object, model.Object, selection);

//                extraction.Verify(ext => ext.RemoveSelection(codeModule.Object, selections));
//            }

//            [TestMethod]
//            [TestCategory("ExtractedMethodRefactoringTests")]
//            public void shouldInsertMethodCall()
//            {

//                var extraction = new Mock<ExtractMethodExtraction>() { CallBase = true };
//                IExtractMethodExtraction SUT = extraction.Object;
//                var codeModule = new Mock<ICodeModuleWrapper>();
//                var model = new Mock<IExtractMethodModel>();
//                var selection = new Selection(7, 1, 7, 1);
//                model.Setup(m => m.PositionForNewMethod).Returns(selection);
//                var methodMock = new Mock<ExtractedMethod>() { CallBase = true };
//                var method = methodMock.Object;
//                method.Accessibility = Accessibility.Private;
//                method.Parameters = new List<ExtractedParameter>();
//                method.MethodName = "NewMethod";
//                methodMock.Setup(m => m.NewMethodCall()).Returns("theMethodCall");
//                model.Setup(m => m.Method).Returns(method);

//                extraction.Setup(em => em.ConstructLinesOfProc(It.IsAny<ICodeModuleWrapper>(), It.IsAny<IExtractMethodModel>())).Returns("theMethodCall");
//                extraction.Setup(em => em.RemoveSelection(It.IsAny<ICodeModuleWrapper>(), It.IsAny<IEnumerable<Selection>>()));

//                var inserted = new List<Tuple<int, string>>();
//                codeModule.Setup(cm => cm.InsertLines(It.IsAny<int>(), It.IsAny<string>()))
//                    .Callback<int, string>((line, data) => inserted.Add(Tuple.Create(line, data)));

//                SUT.Apply(codeModule.Object, model.Object, selection);

//                // selection.StartLine = 7
//                var expected = Tuple.Create(7, "theMethodCall");
//                var actual = inserted[1];
//                //Make sure the second insert inserted the methodCall higher up.
//                Assert.AreEqual(expected, actual);
//            }

//            [TestMethod]
//            [TestCategory("ExtractedMethodRefactoringTests")]
//            public void shouldInsertNewMethodAtGivenLineNoBeforeInsertingMethodCall()
//            {
//                var newProc = @"
//Public Sub NewMethod()
//    DebugPrint ""a""
//End Sub";
//                var extraction = new Mock<ExtractMethodExtraction>() { CallBase = true };
//                IExtractMethodExtraction SUT = extraction.Object;
//                var codeModule = new Mock<ICodeModuleWrapper>();
//                var model = new Mock<IExtractMethodModel>();
//                var selection = new Selection(1, 1, 1, 1);
//                model.Setup(m => m.PositionForNewMethod).Returns(selection);
//                var method = new ExtractedMethod();
//                method.Accessibility = Accessibility.Private;
//                method.Parameters = new List<ExtractedParameter>();
//                method.MethodName = "NewMethod";
//                model.Setup(m => m.Method).Returns(method);
//                extraction.Setup(em => em.ConstructLinesOfProc(It.IsAny<ICodeModuleWrapper>(), It.IsAny<IExtractMethodModel>())).Returns(newProc);
//                extraction.Setup(em => em.RemoveSelection(It.IsAny<ICodeModuleWrapper>(), It.IsAny<IEnumerable<Selection>>()));

//                var inserted = new List<Tuple<int, string>>();
//                codeModule.Setup(cm => cm.InsertLines(It.IsAny<int>(), It.IsAny<string>())).Callback<int, string>((line, data) => inserted.Add(Tuple.Create(line, data)));
//                SUT.Apply(codeModule.Object, model.Object, selection);

//                var expected = Tuple.Create(selection.StartLine, newProc);
//                var actual = inserted[0];
//                //Make sure the first insert inserted the rows.
//                Assert.AreEqual(expected, actual);

//            }

//        }

//        [TestClass]
//        public class WhenConstructLinesOfProcIsCalledWithAListOfSelections
//        {

//            [TestMethod]
//            [TestCategory("ExtractedMethodRefactoringTests")]
//            public void shouldConcatenateASeriesOfLines()
//            {

//                var notifyCalls = new List<Tuple<int, int>>();
//                var codeModule = new Mock<ICodeModuleWrapper>();
//                codeModule.Setup(cm => cm.get_Lines(It.IsAny<int>(), It.IsAny<int>()))
//                    .Callback<int, int>((start, count) => notifyCalls.Add(Tuple.Create(start, count)));
//                var selections = new List<Selection>() { new Selection(5, 1, 5, 20), new Selection(10, 1, 12, 20) };
//                var model = new Mock<IExtractMethodModel>();
//                var method = new ExtractedMethod();
//                method.Accessibility = Accessibility.Private;
//                method.Parameters = new List<ExtractedParameter>();
//                method.MethodName = "NewMethod";
//                model.Setup(m => m.RowsToRemove).Returns(selections);
//                model.Setup(m => m.Method).Returns(method);

//                var SUT = new ExtractMethodExtraction();
//                //Act
//                SUT.ConstructLinesOfProc(codeModule.Object, model.Object);

//                //Assert
//                Assert.AreEqual(Tuple.Create(5, 1), notifyCalls[0]);
//                Assert.AreEqual(Tuple.Create(10, 3), notifyCalls[1]);
//            }

//        }

//        /// <summary>
//        /// https://github.com/rubberduck-vba/Rubberduck/issues/844
//        /// </summary>
//        [TestMethod]
//        [TestCategory("ExtractMethodModelTests")]
//        public void shouldNotProduceDuplicateDimOfz()
//        {

//            #region inputCode
//            var inputCode = @"
//Option explicit
//Public Sub CodeWithDeclaration()
//    Dim x as long
//    Dim y as long
//
//    x = 1 + 2
//    DebugPrint x                      '8
//    y = x + 1
//    Dim z as long
//    z = x  
//    DebugPrint z                      '12
//    x = 2
//    DebugPrint y
//
//
//End Sub                                '17
//Public Sub DebugPrint(byval g as long)
//End Sub
//                                       '20
//
//";

//            var selectedCode = @"
//    DebugPrint x                      '8
//    y = x + 1
//    Dim z as long
//    z = x  
//    DebugPrint z                      '12
//";
//            var expectedCode = @"
//Option explicit
//Public Sub CodeWithDeclaration()
//    Dim x as long
//    Dim y as long
//
//    x = 1 + 2
//NewMethod x, y
//    x = 2
//    DebugPrint y
//
//
//End Sub                                '17
//Private Sub NewMethod(ByRef x As long, ByRef y As long) 
//    Dim z as long
//    DebugPrint x                      '8
//    y = x + 1
//    z = x  
//    DebugPrint z                      '12
//End Sub
//Public Sub DebugPrint(byval g as long)
//End Sub
//                                       '20
//
//";
//            #endregion

//            QualifiedModuleName qualifiedModuleName;
//            RubberduckParserState state;
//            MockParser.ParseString(inputCode, out qualifiedModuleName, out state);
//            var declarations = state.AllDeclarations;

//            var selection = new Selection(8, 1, 12, 50);
//            QualifiedSelection? qSelection = new QualifiedSelection(qualifiedModuleName, selection);

//            List<IExtractMethodRule> emRules = new List<IExtractMethodRule>(){
//                        new ExtractMethodRuleInSelection(),
//                        new ExtractMethodRuleIsAssignedInSelection(),
//                        new ExtractMethodRuleUsedBefore(),
//                        new ExtractMethodRuleUsedAfter(),
//                        new ExtractMethodRuleExternalReference()};

//            var codeModule = new CodeModuleWrapper(qualifiedModuleName.Component.CodeModule);
//            var extractedMethod = new ExtractedMethod();
//            var paramClassify = new ExtractMethodParameterClassification(emRules);
//            var model = new ExtractMethodModel(extractedMethod, paramClassify);
//            model.extract(declarations, qSelection.Value, selectedCode);

//            var SUT = new ExtractMethodExtraction();

//            //Act
//            SUT.Apply(codeModule, model, selection);

//            //Assert
//            var actual = codeModule.get_Lines(1, 1000);
//            Assert.AreEqual(expectedCode, actual);
//        }
//    }
//}
