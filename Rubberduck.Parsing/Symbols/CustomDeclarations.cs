﻿using NLog;
using Rubberduck.Parsing.Annotations;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor;
using System.Collections.Generic;
using System.Diagnostics;

namespace Rubberduck.Parsing.Symbols
{
    public static class CustomDeclarations
    {
        public static Declaration DEBUG_PRINT;
        private static readonly Logger _logger = LogManager.GetCurrentClassLogger();

        public static IReadOnlyList<Declaration> Load(Declaration parentProject, Declaration parentModule)
        {
            _logger.Debug("Loading custom declarations with {0} as parent project and {1} as parent module.", parentProject.IdentifierName, parentModule.IdentifierName);
            List<Declaration> declarations = new List<Declaration>();
            var debugModuleName = new QualifiedModuleName(parentProject.QualifiedName.QualifiedModuleName.ProjectName, parentProject.QualifiedName.QualifiedModuleName.ProjectPath, "DebugClass");
            var debugModule = new ProceduralModuleDeclaration(new QualifiedMemberName(debugModuleName, "DebugModule"), parentProject, "DebugModule", true, new List<IAnnotation>(), new Attributes());
            var debugClassName = new QualifiedModuleName(parentProject.QualifiedName.QualifiedModuleName.ProjectName, parentProject.QualifiedName.QualifiedModuleName.ProjectPath, "DebugClass");
            var debugClass = new ClassModuleDeclaration(new QualifiedMemberName(debugClassName, "DebugClass"), parentProject, "DebugClass", true, new List<IAnnotation>(), new Attributes(), true);
            var debugObject = new Declaration(new QualifiedMemberName(debugClassName, "Debug"), debugModule, "Global", "DebugClass", null, true, false, Accessibility.Global, DeclarationType.Variable, false, null);
            var debugAssert = new SubroutineDeclaration(new QualifiedMemberName(debugClassName, "Assert"), debugClass, debugClass, null, Accessibility.Global, null, Selection.Home, true, null, new Attributes());
            var debugPrint = new SubroutineDeclaration(new QualifiedMemberName(debugClassName, "Print"), debugClass, debugClass, null, Accessibility.Global, null, Selection.Home, true, null, new Attributes());
            declarations.Add(debugModule);
            declarations.Add(debugClass);
            declarations.Add(debugObject);
            declarations.Add(debugAssert);
            declarations.Add(debugPrint);
            declarations.AddRange(AddSpecialFormDeclarations(parentProject, parentModule));
            // Debug.Print has the same special syntax as the print and write statement.
            // Because of that it is treated specially in the grammar and normally wouldn't be resolved.
            // Since we still want it to be resolved we make it easier for the resolver to access the debug print
            // declaration by exposing it in this way.
            DEBUG_PRINT = debugPrint;
            return declarations;
        }

        private static List<Declaration> AddSpecialFormDeclarations(Declaration parentProject, Declaration parentModule)
        {
            List<Declaration> declarations = new List<Declaration>();
            Debug.Assert(parentModule != null);
            var arrayFunction = new FunctionDeclaration(
                new QualifiedMemberName(parentModule.QualifiedName.QualifiedModuleName, "Array"),
                parentModule,
                parentModule,
                "Variant",
                null,
                null,
                Accessibility.Public,
                null,
                Selection.Home,
                false,
                true,
                null,
                new Attributes());
            var inputFunction = new SubroutineDeclaration(new QualifiedMemberName(parentModule.QualifiedName.QualifiedModuleName, "Input"), parentModule, parentModule, "Variant", Accessibility.Public, null, Selection.Home, true, null, new Attributes());
            var numberParam = new ParameterDeclaration(new QualifiedMemberName(parentModule.QualifiedName.QualifiedModuleName, "Number"), inputFunction, "Integer", null, null, false, false);
            var filenumberParam = new ParameterDeclaration(new QualifiedMemberName(parentModule.QualifiedName.QualifiedModuleName, "Filenumber"), inputFunction, "Integer", null, null, false, false);
            inputFunction.AddParameter(numberParam);
            inputFunction.AddParameter(filenumberParam);
            var inputBFunction = new SubroutineDeclaration(new QualifiedMemberName(parentModule.QualifiedName.QualifiedModuleName, "InputB"), parentModule, parentModule, "Variant", Accessibility.Public, null, Selection.Home, true, null, new Attributes());
            var numberBParam = new ParameterDeclaration(new QualifiedMemberName(parentModule.QualifiedName.QualifiedModuleName, "Number"), inputBFunction, "Integer", null, null, false, false);
            var filenumberBParam = new ParameterDeclaration(new QualifiedMemberName(parentModule.QualifiedName.QualifiedModuleName, "Filenumber"), inputBFunction, "Integer", null, null, false, false);
            inputBFunction.AddParameter(numberBParam);
            inputBFunction.AddParameter(filenumberBParam);
            var lboundFunction = new FunctionDeclaration(
                new QualifiedMemberName(parentModule.QualifiedName.QualifiedModuleName, "LBound"),
                parentModule,
                parentModule,
                "Long",
                null,
                null,
                Accessibility.Public,
                null,
                Selection.Home,
                false,
                true,
                null,
                new Attributes());
            var arrayNameParam = new ParameterDeclaration(new QualifiedMemberName(parentModule.QualifiedName.QualifiedModuleName, "Arrayname"), lboundFunction, "Integer", null, null, false, false);
            var dimensionParam = new ParameterDeclaration(new QualifiedMemberName(parentModule.QualifiedName.QualifiedModuleName, "Dimension"), lboundFunction, "Integer", null, null, true, false);
            lboundFunction.AddParameter(arrayNameParam);
            lboundFunction.AddParameter(dimensionParam);
            var uboundFunction = new FunctionDeclaration(
                new QualifiedMemberName(parentModule.QualifiedName.QualifiedModuleName, "UBound"),
                parentModule,
                parentModule,
                "Integer",
                null,
                null,
                Accessibility.Public,
                null,
                Selection.Home,
                false,
                true,
                null,
                new Attributes());
            var arrayParam = new ParameterDeclaration(new QualifiedMemberName(parentModule.QualifiedName.QualifiedModuleName, "Array"), uboundFunction, "Variant", null, null, false, false, true);
            var rankParam = new ParameterDeclaration(new QualifiedMemberName(parentModule.QualifiedName.QualifiedModuleName, "Rank"), uboundFunction, "Integer", null, null, true, false);
            uboundFunction.AddParameter(arrayParam);
            uboundFunction.AddParameter(rankParam);
            declarations.Add(arrayFunction);
            declarations.Add(inputFunction);
            declarations.Add(inputBFunction);
            declarations.Add(lboundFunction);
            declarations.Add(uboundFunction);
            return declarations;
        }
    }
}
