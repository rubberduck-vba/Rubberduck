using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moq;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.ExtractMethod;
using Rubberduck.VBEditor;

namespace RubberduckTests.Refactoring.ExtractMethod
{
    [TestClass]
    class ExtractMethodProcTests
    {

        [TestClass]
        public class WhenLocalVariableConstantIsInternal : ExtractMethodProcTests
        {
            /* When a local variable/constant is only used within the selection
             * its declaration is moved to the extracted method */
            [TestMethod]
            [TestCategory("ExtractMethodProcTests")]
            public void shouldMoveDeclarationToDestinationMethod()
            {
                /*
                 * identify block being moved.
                 * identify declarations found in block
                 */

                #region inputCode

                var inputCode = @"
Option explicit
Public Sub CodeWithDeclaration()
    Dim x as long
    Dim y as long

    x = 1 + 2
    Debug.Print x
    y = x + 1
    Debug.Print y

    x = 4
End Sub
";

                var outputMethod = @"
Private Sub NewMethod(Byref x as long)
    Dim y as long
    y = x + 1
    Debug.Print y
End Sub";

                var selectedCode = @"
y = x + 1 
Debug.Print y";
                #endregion

                QualifiedModuleName qualifiedModuleName;
                RubberduckParserState state;
                MockParser.ParseString(inputCode, out qualifiedModuleName, out state);
                var declarations = state.AllDeclarations;

                var selection = new Selection(9, 1, 10, 17);
                QualifiedSelection? qSelection = new QualifiedSelection(qualifiedModuleName, selection);
                var emRules = new List<IExtractMethodRule>() { };
                var extractedMethod = new Mock<IExtractedMethod>();

                var extractedMethodModel = new ExtractMethodModel(emRules,extractedMethod.Object);
                var SUT = new ExtractMethodProc();
                var actual = SUT.createProc(extractedMethodModel);





            }
        }
    }
}
