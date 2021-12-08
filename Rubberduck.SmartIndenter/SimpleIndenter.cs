using System;
using System.Collections.Generic;
using System.Linq;

namespace Rubberduck.SmartIndenter
{
    /// <summary>
    /// An indenter implementation that does not use Rubberduck.VBEditor types, for use with public API; serves as a base class for Rubberduck's indenter.
    /// </summary>
    public class SimpleIndenter : ISimpleIndenter
    {
        protected virtual Func<IIndenterSettings> Settings { get; }

        /// <summary>
        /// Indents the code contained in the passed string. NOTE: In Rubberduck this overload should only be used on procedures or modules.
        /// </summary>
        /// <remarks>
        /// This overload is intended for use with a public API, not necessarily Rubberduck.
        /// </remarks>
        /// <param name="code">The code block to indent</param>
        /// <returns>Indented code lines</returns>
        public IEnumerable<string> Indent(string code, IIndenterSettings settings = null) => Indent(code.Replace("\r", string.Empty).Split('\n'), false, settings ?? Settings?.Invoke());

        /// <summary>
        /// Indents a range of code lines. NOTE: If inserting procedures, use the forceTrailingNewLines overload to preserve vertical spacing in the module.
        /// Do not call directly on selections. Use Indent(IVBComponent, Selection) instead.
        /// </summary>
        /// <param name="codeLines">Code lines to indent</param>
        /// <returns>Indented code lines</returns>
        public IEnumerable<string> Indent(IEnumerable<string> codeLines, IIndenterSettings settings = null) => Indent(codeLines, false, settings ?? Settings?.Invoke());

        /// <summary>
        /// Indents a range of code lines. Do not call directly on selections. Use Indent(IVBComponent, Selection) instead.
        /// </summary>
        /// <param name="codeLines">Code lines to indent</param>
        /// <param name="forceTrailingNewLines">If true adds a number of blank lines after the last procedure based on VerticallySpaceProcedures settings</param>
        /// <returns>Indented code lines</returns>
        public IEnumerable<string> Indent(IEnumerable<string> codeLines, bool forceTrailingNewLines, IIndenterSettings settings = null) => Indent(codeLines, forceTrailingNewLines, false, settings ?? Settings?.Invoke());

        protected IEnumerable<string> Indent(IEnumerable<string> codeLines, bool forceTrailingNewLines, bool procedure, IIndenterSettings settings)
        {
            if (settings is null)
            {
                settings = new IndenterSettings();
            }

            var logical = BuildLogicalCodeLines(codeLines, settings).ToList();
            var indents = 0;
            var start = false;
            var enumStart = false;
            var inEnumType = false;
            var inProcedure = false;

            foreach (var line in logical)
            {
                inEnumType &= !line.IsEnumOrTypeEnd;
                if (inEnumType)
                {
                    line.AtEnumTypeStart = enumStart;
                    enumStart = line.IsCommentBlock;
                    line.IsEnumOrTypeMember = true;
                    line.InsideProcedureTypeOrEnum = true;
                    line.IndentationLevel = line.EnumTypeIndents;
                    continue;
                }

                if (line.IsProcedureStart)
                {
                    inProcedure = true;
                }
                line.InsideProcedureTypeOrEnum = inProcedure || enumStart;
                inProcedure = inProcedure && !line.IsProcedureEnd && !line.IsEnumOrTypeEnd;
                if (line.IsProcedureStart || line.IsEnumOrTypeStart)
                {
                    indents = 0;
                }

                line.AtProcedureStart = start;
                line.IndentationLevel = indents - line.Outdents;
                indents += line.NextLineIndents;
                start = line.IsProcedureStart ||
                        line.AtProcedureStart && line.IsDeclaration ||
                        line.AtProcedureStart && line.IsCommentBlock ||
                        settings.IgnoreEmptyLinesInFirstBlocks && line.AtProcedureStart && line.IsEmpty;
                inEnumType = line.IsEnumOrTypeStart;
                enumStart = inEnumType;
            }

            return GenerateCodeLineStrings(logical, forceTrailingNewLines, settings, procedure);
        }

        protected IEnumerable<LogicalCodeLine> BuildLogicalCodeLines(IEnumerable<string> lines, IIndenterSettings settings)
        {
            var logical = new List<LogicalCodeLine>();
            LogicalCodeLine current = null;
            AbsoluteCodeLine previous = null;

            foreach (var line in lines)
            {
                var absolute = new AbsoluteCodeLine(line, settings, previous);
                if (current == null)
                {
                    current = new LogicalCodeLine(absolute, settings);
                    logical.Add(current);
                }
                else
                {
                    current.AddContinuationLine(absolute);
                }

                if (!absolute.HasContinuation)
                {
                    current = null;
                }
                previous = absolute;
            }
            return logical;
        }

        protected IEnumerable<string> GenerateCodeLineStrings(IEnumerable<LogicalCodeLine> logical, bool forceTrailingNewLines, IIndenterSettings settings, bool procedure = false)
        {
            var output = new List<string>();

            List<LogicalCodeLine> indent;
            if (!procedure && settings.VerticallySpaceProcedures)
            {
                indent = new List<LogicalCodeLine>();
                var lines = logical.ToArray();
                var header = true;
                var inEnumType = false;
                var propertyGroupIdentifier = string.Empty;

                for (var i = 0; i < lines.Length; i++)
                {
                    indent.Add(lines[i]);

                    if (header && lines[i].IsEnumOrTypeStart)
                    {
                        inEnumType = true;
                    }
                    if (header && lines[i].IsEnumOrTypeEnd)
                    {
                        inEnumType = false;
                    }

                    if (header && !inEnumType && lines[i].IsProcedureStart)
                    {
                        header = false;
                        SpaceHeader(indent, settings);

                        propertyGroupIdentifier = lines[i].IsPropertyStart
                            ? ExtractPropertyIdentifier(lines[i].ToString())
                            : string.Empty;

                        continue;
                    }
                    if (!lines[i].IsEnumOrTypeEnd && !lines[i].IsProcedureEnd)
                    {
                        continue;
                    }

                    while (++i < lines.Length && lines[i].IsEmpty) { }
                    if (i != lines.Length)
                    {
                        var linesBetweenProcedures = settings.LinesBetweenProcedures;

                        if (lines[i].IsPropertyStart)
                        {
                            var propertyIdentifier = ExtractPropertyIdentifier(lines[i].ToString());
                            if (propertyIdentifier.Equals(propertyGroupIdentifier, StringComparison.InvariantCultureIgnoreCase)
                                && settings.GroupRelatedProperties)
                            {
                                linesBetweenProcedures = 0;
                            }
                            else
                            {
                                propertyGroupIdentifier = propertyIdentifier;
                            }
                        }

                        if (linesBetweenProcedures > 0)
                        {
                            indent.Add(new LogicalCodeLine(Enumerable.Repeat(new AbsoluteCodeLine(string.Empty, settings), linesBetweenProcedures), settings));
                        }

                        indent.Add(lines[i]);
                    }
                    else if (forceTrailingNewLines && i == lines.Length)
                    {
                        indent.Add(new LogicalCodeLine(Enumerable.Repeat(new AbsoluteCodeLine(string.Empty, settings), Math.Max(settings.LinesBetweenProcedures, 1)), settings));
                    }
                }
            }
            else
            {
                indent = logical.ToList();
            }

            foreach (var line in indent)
            {
                output.AddRange(line.Indented().Split(new[] { Environment.NewLine }, StringSplitOptions.None));
            }
            return output;
        }

        protected static string ExtractPropertyIdentifier(string line)
        {
            var declarationElementsStartingAtGetLetOrSet = line.ToString().Split(' ').SkipWhile(c => !c.EndsWith("et")).ToArray();
            return declarationElementsStartingAtGetLetOrSet[1].Split(new string[] { "(" }, StringSplitOptions.None).FirstOrDefault();
        }

        protected static void SpaceHeader(IList<LogicalCodeLine> header, IIndenterSettings settings)
        {
            var commentSkipped = false;
            var commentLines = 0;
            for (var i = header.Count - 2; i >= 0; i--)
            {
                if (!commentSkipped && header[i].IsCommentBlock)
                {
                    commentLines++;
                    continue;
                }

                commentSkipped = true;
                if (header[i].IsEmpty)
                {
                    header.RemoveAt(i);
                }
                else
                {
                    header.Insert(header.Count - 1 - commentLines,
                        new LogicalCodeLine(
                            Enumerable.Repeat(new AbsoluteCodeLine(string.Empty, settings),
                                settings.LinesBetweenProcedures), settings));
                    return;
                }
            }
        }
    }
}
