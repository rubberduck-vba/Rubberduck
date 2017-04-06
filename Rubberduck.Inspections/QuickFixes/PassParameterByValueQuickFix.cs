using System;
using System.Collections.Generic;
using System.Linq;
using Rubberduck.Common;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Inspections.Resources;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor;

namespace Rubberduck.Inspections.QuickFixes
{
    public class PassParameterByValueQuickFix : IQuickFix
    {
        private readonly RubberduckParserState _state;
        private static readonly HashSet<Type> _supportedInspections = new HashSet<Type>
        {
            typeof(ParameterCanBeByValInspection)
        };

        public PassParameterByValueQuickFix(RubberduckParserState state)
        {
            _state = state;
        }

        public static IReadOnlyCollection<Type> SupportedInspections => _supportedInspections.ToList();

        public static void AddSupportedInspectionType(Type inspectionType)
        {
            if (!inspectionType.GetInterfaces().Contains(typeof(IInspection)))
            {
                throw new ArgumentException("Type must implement IInspection", nameof(inspectionType));
            }

            _supportedInspections.Add(inspectionType);
        }

        public void Fix(IInspectionResult result)
        {
            if (result.Target.ParentDeclaration.DeclarationType == DeclarationType.Event ||
                _state.AllUserDeclarations.FindInterfaceMembers().Contains(result.Target.ParentDeclaration))
            {
                FixMethods(result.Target);
            }
            else
            {
                FixMethod((VBAParser.ArgContext)result.Target.Context, result.QualifiedSelection);
            }
        }

        public string Description(IInspectionResult result)
        {
            return InspectionsUI.PassParameterByValueQuickFix;
        }

        private void FixMethods(Declaration target)
        {
            var declarationParameters =
                _state.AllUserDeclarations.Where(declaration => declaration.DeclarationType == DeclarationType.Parameter &&
                                                                Equals(declaration.ParentDeclaration, target.ParentDeclaration))
                    .OrderBy(o => o.Selection.StartLine)
                    .ThenBy(t => t.Selection.StartColumn)
                    .ToList();

            var parameterIndex = declarationParameters.IndexOf(target);
            if (parameterIndex == -1)
            {
                return; // should only happen if the parse results are stale; prevents a crash in that case
            }

            var members = target.ParentDeclaration.DeclarationType == DeclarationType.Event
                ? _state.AllUserDeclarations.FindHandlersForEvent(target.ParentDeclaration)
                    .Select(s => s.Item2)
                    .ToList()
                : _state.AllUserDeclarations.FindInterfaceImplementationMembers(target.ParentDeclaration).ToList();

            foreach (var member in members)
            {
                var parameters =
                    _state.AllUserDeclarations.Where(declaration => declaration.DeclarationType == DeclarationType.Parameter &&
                                                                    Equals(declaration.ParentDeclaration, member))
                        .OrderBy(o => o.Selection.StartLine)
                        .ThenBy(t => t.Selection.StartColumn)
                        .ToList();

                FixMethod((VBAParser.ArgContext)parameters[parameterIndex].Context,
                    parameters[parameterIndex].QualifiedSelection);
            }

            FixMethod((VBAParser.ArgContext)declarationParameters[parameterIndex].Context,
                declarationParameters[parameterIndex].QualifiedSelection);
        }

        private void FixMethod(VBAParser.ArgContext context, QualifiedSelection qualifiedSelection)
        {
            var parameter = context.GetText();
            var argList = context.parent.GetText();

            var module = qualifiedSelection.QualifiedName.Component.CodeModule;
            {
                string result;
                if (context.BYREF() != null)
                {
                    result = parameter.Replace(Tokens.ByRef, Tokens.ByVal);
                }
                else if (context.OPTIONAL() != null)
                {
                    result = parameter.Replace(Tokens.Optional, Tokens.Optional + ' ' + Tokens.ByVal);
                }
                else
                {
                    result = Tokens.ByVal + ' ' + parameter;
                }

                var startLine = 0;
                var stopLine = 0;
                try
                {
                    dynamic proc = context.parent.parent;
                    startLine = proc.GetType().GetProperty("Start").GetValue(proc).Line;
                    stopLine =  proc.GetType().GetProperty("Stop").GetValue(proc).Line;
                }
                catch { return; }

                var code = module.GetLines(startLine, stopLine - startLine + 1);
                result = code.Replace(argList, argList.Replace(parameter, result));

                foreach (var line in result.Split(new[] { "\r\n" }, StringSplitOptions.None))
                {
                    module.ReplaceLine(startLine++, line);
                }  
            }
        }

        public bool CanFixInProcedure => true;
        public bool CanFixInModule => true;
        public bool CanFixInProject => true;
    }
}