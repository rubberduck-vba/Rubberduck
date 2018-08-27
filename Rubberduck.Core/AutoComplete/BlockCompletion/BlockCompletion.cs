using Rubberduck.Parsing.VBA.Parsing;
using Rubberduck.VBEditor;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace Rubberduck.AutoComplete.BlockCompletion
{
    //public class BlockCompletion
    //{
    //    private static readonly string PlaceholderPattern =
    //        @"\$(?<placeholder>({(?<namedtabindex>\d+):(?<identifier>[^}]+)})|(?<tabindex>\d+))";

    //    public BlockCompletion(string name, string prefix, string[] body, string description, ParserStartRule startRule = null, bool onlyValidInScope = true, bool validInScope = true)
    //    {
    //        StartRule = startRule;
    //        Prefix = prefix;
    //        Description = description;
    //        IsOnlyValidInScope = onlyValidInScope;
    //        IsValidInScope = validInScope;
    //        _body = body;

    //        var stops = body.SelectMany((line, i) => 
    //            Regex.Matches(line, PlaceholderPattern).Cast<Match>()
    //                 .Select(m => new
    //                 {
    //                     Line = i,
    //                     Key = m.Groups["tabindex"].Success ? int.Parse(m.Groups["tabindex"].Value) : int.Parse(m.Groups["namedtabindex"].Value),
    //                     Match = m
    //                 }))
    //                 .OrderBy(m => (m.Match.Groups["namedtabindex"].Success ? 0 : 1))
    //                 .ThenBy(m => m.Key);
    //        var escapedBody = string.Join("\n", body)
    //            .Replace(@"\", @"\\")
    //            .Replace(@"(", @"\(")
    //            .Replace(@")", @"\)")
    //            .Replace(@"^", @"\^")
    //            .Replace(@"[", @"\[")
    //            .Replace(@"]", @"\]")
    //            .Replace(@"<", @"\<")
    //            .Replace(@">", @"\>")
    //            .Replace(@"*", @"\*")
    //            .Replace(@".", @"\.")
    //            .Replace(@",", @"\,")
    //            .Replace(@"+", @"\+")
    //            .Replace(@"?", @"\?");

    //        Syntax = Regex.Replace(escapedBody, PlaceholderPattern, m => $"(.*)")
    //            .Replace("\t", @" "); // ??

    //        _tabStops = (from m in stops
    //                    let startCol = m.Match.Index
    //                    let startLine = m.Line
    //                    let endCol = startCol + m.Match.Value.Length
    //                    let endLine = m.Line
    //                    select new { Identifier = m.Match.Groups["identifier"].Value, Position = new Selection(startLine, startCol, endLine, endCol) })
    //                .Select(item => new TabStop(item.Identifier, item.Position))
    //                .ToList();
    //    }

    //    public ParserStartRule StartRule { get; }

    //    private readonly IList<TabStop> _tabStops;
    //    public IReadOnlyList<TabStop> TabStops => _tabStops.ToArray();

    //    public TabStop GetTabStop(int index)
    //    {
    //        return _tabStops[index];
    //    }

    //    public string Syntax { get; }
    //    public string Name { get; }
    //    public string Prefix { get; }

    //    public bool IsOnlyValidInScope { get; }
    //    public bool IsValidInScope { get; }

    //    private readonly string[] _body;

    //    /// <summary>
    //    /// Gets the completion body, with placeholder syntax.
    //    /// </summary>
    //    public string Body => string.Join("\r\n", _body);

    //    /// <summary>
    //    /// Gets the <see cref="Body"/> as it would appear in the code pane.
    //    /// </summary>
    //    public string[] CodeBody => _body.Select(line => Regex.Replace(line, PlaceholderPattern, m => m.Groups["identifier"].Success ? m.Groups["identifier"].Value : string.Empty)).ToArray();
    //    public string Description { get; }

    //    public class TabStop
    //    {
    //        public TabStop(string content, Selection position)
    //        {
    //            Content = content;
    //            Position = position;
    //        }

    //        public string Content { get; set; }

    //        public Selection Position { get; set; }
    //    }
    //}
}
