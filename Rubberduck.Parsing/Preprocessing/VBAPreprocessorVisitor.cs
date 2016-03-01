using Antlr4.Runtime;
using Antlr4.Runtime.Misc;
using Rubberduck.Parsing.Date;
using Rubberduck.Parsing.Like;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace Rubberduck.Parsing.Preprocessing
{
    public sealed class VBAPreprocessorVisitor : VBAConditionalCompilationBaseVisitor<object>
    {
        private readonly VBAExpressionEvaluator _evaluator;

        public VBAPreprocessorVisitor(SymbolTable symbolTable, VBAPredefinedCompilationConstants predefinedConstants)
        {
            _evaluator = new VBAExpressionEvaluator(symbolTable);
            _evaluator.AddConstant(VBAPredefinedCompilationConstants.VBA6_NAME, predefinedConstants.VBA6);
            _evaluator.AddConstant(VBAPredefinedCompilationConstants.VBA7_NAME, predefinedConstants.VBA7);
            _evaluator.AddConstant(VBAPredefinedCompilationConstants.WIN64_NAME, predefinedConstants.Win64);
            _evaluator.AddConstant(VBAPredefinedCompilationConstants.WIN32_NAME, predefinedConstants.Win32);
            _evaluator.AddConstant(VBAPredefinedCompilationConstants.WIN16_NAME, predefinedConstants.Win16);
            _evaluator.AddConstant(VBAPredefinedCompilationConstants.MAC_NAME, predefinedConstants.Mac);
        }

        public override object VisitCompilationUnit([NotNull] VBAConditionalCompilationParser.CompilationUnitContext context)
        {
            return Visit(context.ccBlock());
        }

        public override object VisitLogicalLine([NotNull] VBAConditionalCompilationParser.LogicalLineContext context)
        {
            return context.GetText();
        }

        public override object VisitCcBlock([NotNull] VBAConditionalCompilationParser.CcBlockContext context)
        {
            var results = new List<object>();
            if (context.children == null)
            {
                return string.Empty;
            }
            foreach (var child in context.children)
            {
                results.Add(Visit(child));
            }
            return string.Join(string.Empty, results);
        }

        public override object VisitCcConst([NotNull] VBAConditionalCompilationParser.CcConstContext context)
        {
            // 3.4.1: If <cc-var-lhs> is a <TYPED-NAME> with a <type-suffix>, the <type-suffix> is ignored.
            var identifier = context.ccVarLhs().name().IDENTIFIER().GetText();
            _evaluator.EvaluateConstant(identifier, context.ccExpression());
            return MarkLineAsDead(context.GetText());
        }

        public override object VisitCcIfBlock([NotNull] VBAConditionalCompilationParser.CcIfBlockContext context)
        {
            StringBuilder builder = new StringBuilder();
            List<bool> conditions = new List<bool>();
            builder.Append(MarkLineAsDead(context.ccIf().GetText()));
            var ifIsAlive = _evaluator.EvaluateCondition(context.ccIf().ccExpression());
            conditions.Add(ifIsAlive);
            var ifBlock = (string)Visit(context.ccBlock());
            builder.Append(EvaluateLiveliness(ifBlock, ifIsAlive));
            foreach (var elseIfBlock in context.ccElseIfBlock())
            {
                builder.Append(MarkLineAsDead(elseIfBlock.ccElseIf().GetText()));
                var elseIfIsAlive = _evaluator.EvaluateCondition(elseIfBlock.ccElseIf().ccExpression());
                conditions.Add(elseIfIsAlive);
                var block = (string)Visit(elseIfBlock.ccBlock());
                builder.Append(EvaluateLiveliness(block, elseIfIsAlive));
            }
            if (context.ccElseBlock() != null)
            {
                builder.Append(MarkLineAsDead(context.ccElseBlock().ccElse().GetText()));
                var block = (string)Visit(context.ccElseBlock().ccBlock());
                var elseIsAlive = conditions.All(condition => !condition);
                builder.Append(EvaluateLiveliness(block, elseIsAlive));
            }
            builder.Append(MarkLineAsDead(context.ccEndIf().GetText()));
            return builder.ToString();
        }

        private string EvaluateLiveliness(string code, bool isAlive)
        {
            if (isAlive)
            {
                return code;
            }
            else
            {
                return MarkAsDead(code);
            }
        }

        private string MarkAsDead(string code)
        {
            bool hasNewLine = false;
            if (code.EndsWith(Environment.NewLine))
            {
                hasNewLine = true;
            }
            // Remove parsed new line.
            code = code.TrimEnd('\r', '\n');
            var lines = code.Split(new string[] { Environment.NewLine }, StringSplitOptions.None);
            var result = string.Join(Environment.NewLine, lines.Select(_ => string.Empty));
            if (hasNewLine)
            {
                result += Environment.NewLine;
            }
            return result;
        }

        private string MarkLineAsDead(string line)
        {
            var result = string.Empty;
            if (line.EndsWith(Environment.NewLine))
            {
                result += Environment.NewLine;
            }
            return result;
        }
    }
}
