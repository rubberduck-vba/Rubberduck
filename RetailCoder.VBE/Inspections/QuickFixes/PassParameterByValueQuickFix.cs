using System;
using System.Linq;
using Antlr4.Runtime;
using Rubberduck.Common;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Resources;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor;

namespace Rubberduck.Inspections.QuickFixes
{
    public class PassParameterByValueQuickFix : QuickFixBase
    {
        private readonly RubberduckParserState _state;
        private readonly Declaration _target;

        public PassParameterByValueQuickFix(RubberduckParserState state, Declaration target, ParserRuleContext context, QualifiedSelection selection)
            : base(context, selection, InspectionsUI.PassParameterByValueQuickFix)
        {
            _state = state;
            _target = target;
        }

        public override void Fix()
        {
            if (_target.ParentDeclaration.DeclarationType == DeclarationType.Event ||
                _state.AllUserDeclarations.FindInterfaceMembers().Contains(_target.ParentDeclaration))
            {
                FixMethods();
            }
            else
            {
                FixMethod((VBAParser.ArgContext)Context, Selection);
            }
        }

        private void FixMethods()
        {
            var declarationParameters =
                _state.AllUserDeclarations.Where(declaration => declaration.DeclarationType == DeclarationType.Parameter &&
                                                                Equals(declaration.ParentDeclaration, _target.ParentDeclaration))
                    .OrderBy(o => o.Selection.StartLine)
                    .ThenBy(t => t.Selection.StartColumn)
                    .ToList();

            var parameterIndex = declarationParameters.IndexOf(_target);
            if (parameterIndex == -1)
            {
                return; // should only happen if the parse results are stale; prevents a crash in that case
            }

            var members = _target.ParentDeclaration.DeclarationType == DeclarationType.Event
                ? _state.AllUserDeclarations.FindHandlersForEvent(_target.ParentDeclaration)
                    .Select(s => s.Item2)
                    .ToList()
                : _state.AllUserDeclarations.FindInterfaceImplementationMembers(_target.ParentDeclaration).ToList();

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
    }
}