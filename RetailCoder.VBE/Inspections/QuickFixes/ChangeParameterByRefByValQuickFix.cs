using System;
using Antlr4.Runtime;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Parsing.Grammar;
using Rubberduck.VBEditor;

namespace Rubberduck.Inspections.QuickFixes
{
    public class ChangeParameterByRefByValQuickFix : QuickFixBase
    {
        private readonly string _newToken;

        public ChangeParameterByRefByValQuickFix(ParserRuleContext context, QualifiedSelection selection, string description, string newToken) 
            : base(context, selection, description)
        {
            _newToken = newToken;
        }

        public override void Fix()
        {
            try
            {
                dynamic context = Context;
                var parameter = Context.GetText();
                dynamic args = Context.parent;
                var argList = args.GetText();
                var module = Selection.QualifiedName.Component.CodeModule;
                {
                    string result;
                    if (context.OPTIONAL() != null)
                    {
                        result = parameter.Replace(Tokens.Optional, Tokens.Optional + ' ' + _newToken);
                    }
                    else
                    {
                        result = _newToken + ' ' + parameter;
                    }

                    dynamic proc = args.parent;
                    var startLine = proc.GetType().GetProperty("Start").GetValue(proc).Line;
                    var stopLine = proc.GetType().GetProperty("Stop").GetValue(proc).Line;
                    var code = module.GetLines(startLine, stopLine - startLine + 1);
                    result = code.Replace(argList, argList.Replace(parameter, result));

                    foreach (var line in result.Split(new[] {"\r\n"}, StringSplitOptions.None))
                    {
                        module.ReplaceLine(startLine++, line);
                    }
                }
            }
            // ReSharper disable once EmptyGeneralCatchClause
            catch { }
        }
    }
}