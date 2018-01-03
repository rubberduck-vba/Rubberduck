//namespace RubberduckTests.Refactoring.ExtractMethod
//{
//    [TestFixture]
//    public class ExtractMethodRefactoringTests
//    {
//        [TestFixture]
//        public class WhenRefactorIsCalled
//        {
//            #region codeSnipets
//            const string inputCode = @"
//    Public Sub ChangeMeIntoDecs()
//        Dim x As Integer
//        x = 1
//        x = 1 + 2
//    End Sub
//";
//            const string newMethod = @"
//Private Sub Bar(ByRef x as integer)
//    x = 1
//End Function";
//            const string extractCode = "x = 1 + 2";
//            #endregion codeSnipets

//            [Test]
//            public void shouldCallParseRequest()
//            {

//                QualifiedModuleName qualifiedModuleName;
//                RubberduckParserState state;
//                MockParser.ParseString(inputCode, out qualifiedModuleName, out state);

//                var declarations = state.AllDeclarations;
//                var selection = new Selection(4, 4, 4, 14);
//                QualifiedSelection? qualifiedSelection = new QualifiedSelection(qualifiedModuleName, selection);
//                var codeModule = new Mock<ICodeModuleWrapper>();
//                var extractedMethod = new Mock<IExtractedMethod>();
//                var wasParsed = false;
//                Action<object> onParseRequest = (obj) => { wasParsed = true; };
//                extractedMethod.Setup(em => em.MethodName).Returns("Bar");

//                var paramClassify = new Mock<IExtractMethodParameterClassification>();
//                var model = new ExtractMethodModel(extractedMethod.Object, paramClassify.Object);
//                model.extract(declarations, qualifiedSelection.Value, extractCode);
//                model.extract(declarations, qualifiedSelection.Value, extractCode);
//                var insertCode = "Bar x";

//                Func<QualifiedSelection?, string, IExtractMethodModel> createMethodModel = (q, s) => { return model; };

//                codeModule.SetupGet(cm => cm.QualifiedSelection).Returns(qualifiedSelection);
//                codeModule.Setup(cm => cm.GetLines(selection)).Returns(extractCode);
//                codeModule.Setup(cm => cm.DeleteLines(It.IsAny<Selection>()));
//                codeModule.Setup(cm => cm.InsertLines(It.IsAny<int>(), It.IsAny<String>()));

//                var extraction = new Mock<IExtractMethodExtraction>();

//                var SUT = new ExtractMethodRefactoring(codeModule.Object, onParseRequest, createMethodModel, extraction.Object);

//                SUT.Refactor();

//                Assert.AreEqual(true, wasParsed);

//            }

//            [Test]
//            public void shouldCallApplyOnExtraction()
//            {

//                QualifiedModuleName qualifiedModuleName;
//                RubberduckParserState state;
//                MockParser.ParseString(inputCode, out qualifiedModuleName, out state);

//                var declarations = state.AllDeclarations;
//                var selection = new Selection(4, 4, 4, 14);
//                QualifiedSelection? qualifiedSelection = new QualifiedSelection(qualifiedModuleName, selection);
//                var codeModule = new Mock<ICodeModuleWrapper>();
//                var extractedMethod = new Mock<IExtractedMethod>();
//                Action<object> onParseRequest = (obj) => { };
//                extractedMethod.Setup(em => em.MethodName).Returns("Bar");

//                var paramClassify = new Mock<IExtractMethodParameterClassification>();
//                var model = new ExtractMethodModel(extractedMethod.Object, paramClassify.Object);
//                model.extract(declarations, qualifiedSelection.Value, extractCode);
//                var insertCode = "Bar x";

//                Func<QualifiedSelection?, string, IExtractMethodModel> createMethodModel = (q, s) => { return model; };

//                codeModule.SetupGet(cm => cm.QualifiedSelection).Returns(qualifiedSelection);
//                codeModule.Setup(cm => cm.GetLines(selection)).Returns(extractCode);
//                codeModule.Setup(cm => cm.DeleteLines(It.IsAny<Selection>()));
//                codeModule.Setup(cm => cm.InsertLines(It.IsAny<int>(), It.IsAny<String>()));

//                var extraction = new Mock<IExtractMethodExtraction>();

//                var SUT = new ExtractMethodRefactoring(codeModule.Object, onParseRequest, createMethodModel, extraction.Object);

//                SUT.Refactor();

//                extraction.Verify(ext => ext.Apply(It.IsAny<ICodeModuleWrapper>(), It.IsAny<IExtractMethodModel>(), It.IsAny<Selection>()));
                
//            }
//        }
//    }

//    [TestFixture]
//    public class Issues
//    {
//        [TestFixture]
//        public class Issue_844_InternalDimIsDuplicatedWhenExtracted
//        {
//        }
//    }

//    [TestFixture]
//    public class Example
//    {

//        #region desired process
//        const string inputCode = @"
//Option Explicit
//Private Sub Foo()
//    Dim x As Integer
//    x = 1 + 2
//End Sub";

//        const string extractCode = "x = 1 + 2";
//        const string insertCode = "Bar x ";
//        const string newMethod = @"
//Private Function Bar(byref integer x) As Integer
//    x = 1 + 2
//End Function
//";
//        #endregion codeparts
//        [TestFixture]
//        public class WhenASimpleExampleIsRun : Example
//        {
//            const string inputCode = @"
//    Public Sub ChangeMeIntoDecs()
//        Dim x As Integer
//        x = 1
//        x = 1 + 2
//    End Sub
//";
//            const string newMethod = @"
//Private Sub Bar(ByRef x as integer)
//    x = 1
//End Function";

//            [Category("ExtractedMethodRefactoringTests")]
//            [Test]
//            public void shouldCallTheExtraction()
//            {
//                QualifiedModuleName qualifiedModuleName;
//                RubberduckParserState state;
//                MockParser.ParseString(inputCode, out qualifiedModuleName, out state);

//                var declarations = state.AllDeclarations;
//                var selection = new Selection(4, 4, 4, 14);
//                QualifiedSelection? qualifiedSelection = new QualifiedSelection(qualifiedModuleName, selection);
//                var codeModule = new Mock<ICodeModuleWrapper>();

//                var emRules = new List<IExtractMethodRule>() { 
//                    new ExtractMethodRuleUsedAfter(), new ExtractMethodRuleUsedBefore(), new ExtractMethodRuleInSelection(), new ExtractMethodRuleIsAssignedInSelection()};
//                var extractedMethod = new Mock<IExtractedMethod>();
//                Action<object> onParseRequest = (obj) => { };
//                extractedMethod.Setup(em => em.MethodName).Returns("Bar");
//                var paramClassify = new Mock<IExtractMethodParameterClassification>();
//                var model = new ExtractMethodModel(extractedMethod.Object, paramClassify.Object);
//                model.extract(declarations, qualifiedSelection.Value, extractCode);
//                var insertCode = "Bar x";

//                Func<QualifiedSelection?, string, IExtractMethodModel> createMethodModel = (q, s) => { return model; };

//                codeModule.SetupGet(cm => cm.QualifiedSelection).Returns(qualifiedSelection);
//                codeModule.Setup(cm => cm.GetLines(selection)).Returns(extractCode);
//                codeModule.Setup(cm => cm.DeleteLines(It.IsAny<Selection>()));
//                codeModule.Setup(cm => cm.InsertLines(It.IsAny<int>(), It.IsAny<String>()));

//                var extraction = new Mock<IExtractMethodExtraction>();

//                var SUT = new ExtractMethodRefactoring(codeModule.Object, onParseRequest, createMethodModel, extraction.Object);

//                SUT.Refactor();

//                extraction.Verify(ext => ext.Apply(codeModule.Object, It.IsAny<IExtractMethodModel>(), selection));
//            }
//        }

//        /* Initially I want default output to be only Subs with Byref Params */

//        /* tests to ignore therefore : 
//         * - When there is no output needed - refactoring extracts a Sub *
//         * - When there is only one possible output - refactoring extracts a Function and returns that value/reference *
//         * - When there are multiple possible outputs, refactoring extracts a Function and returns whichever selected output the user selected; other outputs are ByRef parameters *
//         * - When the return value is a reference - the return assignment is Set initially implement with return values returned as ByRef */


//        //[TestFixture]
//        public class when_local_variable_is_only_used_before_the_selection : ExtractMethodModelTests
//        {
//            /* When a local variable/constant is only used before the selection, 
//             * its declaration remains where it was */
//            //[Test]
//            public void should_leave_declaration_in_source_method()
//            {
//            }
//        }
//        //[TestFixture]
//        public class when_local_variable_is_only_used_after_the_selection : ExtractMethodModelTests
//        {
//            /* When a local variable/constant is only used after the selection, 
//             * its declaration remains where it was */
//            //[Test]
//            public void should_leave_declaration_in_source_method()
//            {

//            }

//        }
//        //[TestFixture]
//        public class when_local_variable_is_used_before_and_within_the_selection : ExtractMethodModelTests
//        {
//            /* When a local variable is used before and within the selction, 
//             * it's considered an input */
//            //[Test]
//            public void should_be_passed_as_a_byref_parameter()
//            {
//            }
//        }
//        //[TestFixture]
//        public class when_local_variable_is_used_after_and_within_the_selection : ExtractMethodModelTests
//        {
//            /* When a local variable is used after and within the selection, 
//             * it's considered an output */
//            //[Test]
//            public void should_be_passed_as_a_byref_parameter()
//            {
//            }
//        }
//        //[TestFixture]
//        public class when_multiple_values_are_updated_within_selection : ExtractMethodModelTests
//        {

//            public void should_add_byref_param_for_each()
//            {
//            }

//        }
//        //[TestFixture]
//        public class when_selection_contains_a_line_label_refered_to_before_the_selection : ExtractMethodModelTests
//        {
//            /* This rules out extracting ErrHandler subroutines 
//             * and avoids having to deal with Return and Resume statements. */
//            public void should_return_null()
//            {
//            }
//            public void should_report_invalid_selection_label_conflict()
//            {
//            }

//        }

//        //[TestFixture]
//        public class when_selection_contains_a_line_label_only_referred_to_within_the_selection : ExtractMethodModelTests
//        {
//            [Test]
//            public void should_move_the_label_and_reference_to_destination_method()
//            {

//            }
//        }
//    }

//}
