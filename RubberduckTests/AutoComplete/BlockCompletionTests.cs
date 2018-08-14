using Moq;
using NUnit.Framework;
//using Rubberduck.AutoComplete.BlockCompletion;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using System.Windows.Forms;

namespace RubberduckTests.AutoComplete
{
    //[TestFixture]
    //public class BlockCompletionTests
    //{
    //    private (string NewCode, Selection NewSelection) Run(Keys keypress, string currentLine, Selection pSelection, BlockCompletionService service = null)
    //    {
    //        var sut = service ?? new BlockCompletionService(TestBlockCompletions);
    //        var module = new Mock<ICodeModule>();
    //        var (newCode, newSelection) = sut.Run(Keys.None, currentLine, pSelection, module.Object);
    //        return (newCode, newSelection);
    //    }

    //    [Test]
    //    public void NoKeyDoesNothing()
    //    {
    //        var prefix = TestBlockCompletions[0].Prefix;

    //        var code = prefix;
    //        var initialSelection = new Selection(1, prefix.Length);

    //        // Keys.None means we had a WM_CHAR and a char content, rather than a keypress (e.g. DELETE, TAB, etc.).
    //        var (newCode, newSelection) = Run(Keys.None, code, initialSelection);

    //        var passCode = code == newCode;
    //        var passSelection = initialSelection == newSelection;

    //        Assert.AreEqual(code, newCode, "Code mismatch.");
    //        Assert.AreEqual(initialSelection, newSelection, "Selection mismatch.");
    //    }

    //    [Test]
    //    public void MatchingSelectsPrefix()
    //    {
    //        var prefix = TestBlockCompletions[0].Prefix;

    //        var code = prefix; // must match
    //        var initialSelection = new Selection(1, prefix.Length);

    //        var (newCode, newSelection) = Run(Keys.None, code, initialSelection);

    //        Assert.AreEqual(new Selection(1, 1, 1, prefix.Length), newSelection);
    //    }

    //    [Test]
    //    public void MatchingSetsCurrentCompletion()
    //    {
    //        var prefix = TestBlockCompletions[0].Prefix;

    //        var code = prefix;
    //        var initialSelection = new Selection(1, prefix.Length);

    //        var sut = new BlockCompletionService(TestBlockCompletions);
    //        var (newCode, newSelection) = Run(Keys.None, code, initialSelection, sut);

    //        Assert.IsNotNull(sut.Current);
    //    }

    //    [Test]
    //    public void EscapeCurrentMatchResetsSelection()
    //    {
    //        var prefix = TestBlockCompletions[0].Prefix;

    //        var code = prefix;
    //        var initialSelection = new Selection(1, 1, 1, prefix.Length);

    //        var (newCode, newSelection) = Run(Keys.Escape, code, initialSelection);

    //        Assert.AreEqual(new Selection(1, prefix.Length + 1), newSelection);
    //    }

    //    [Test]
    //    public void EscapeMatchingLeavesCodeUnchanged()
    //    {
    //        var prefix = TestBlockCompletions[0].Prefix;

    //        var code = prefix;
    //        var initialSelection = new Selection(1, 1, 1, prefix.Length);

    //        var (newCode, newSelection) = Run(Keys.Escape, code, initialSelection);

    //        Assert.AreEqual(code, newCode);
    //    }

    //    [Test]
    //    public void EscapeMatchingResetsCurrentCompletion()
    //    {
    //        var prefix = TestBlockCompletions[0].Prefix;

    //        var code = prefix;
    //        var initialSelection = new Selection(1, 1, 1, prefix.Length);

    //        var sut = new BlockCompletionService(TestBlockCompletions);
    //        var (newCode, newSelection) = Run(Keys.Escape, code, initialSelection, sut);

    //        Assert.IsNull(sut.Current);
    //    }

    //    [Test]
    //    public void MatchingTabCompletesBlock()
    //    {
    //        var completion = TestBlockCompletions[0];
    //        var prefix = completion.Prefix;

    //        var code = prefix;
    //        var initialSelection = new Selection(1, 1, 1, prefix.Length);

    //        var sut = new BlockCompletionService(TestBlockCompletions)
    //        {
    //            Current = completion
    //        };
    //        var (newCode, newSelection) = Run(Keys.Tab, code, initialSelection, sut);

    //        Assert.IsNull(sut.Current);
    //    }

    //    [Test]
    //    public void GivenCompletedMatch_TabMovesToPlaceholder()
    //    {
    //        var completion = TestBlockCompletions[0];
    //        var prefix = completion.Prefix;
    //        var placeholder = completion.TabStops[0];

    //        var code = string.Join("\r\n", completion.CodeBody);
    //        var initialSelection = new Selection(1, 1, 1, prefix.Length);

    //        var sut = new BlockCompletionService(TestBlockCompletions)
    //        {
    //            Current = completion
    //        };
    //        var (newCode, newSelection) = Run(Keys.Tab, code, initialSelection, sut);

    //        Assert.AreEqual(placeholder.Position.ToOneBased(), newSelection);
    //    }

    //    [Test]
    //    public void GivenTabStops_TabMovesToNextPlaceholder()
    //    {
    //        var completion = TestBlockCompletions[0];
    //        var prefix = completion.Prefix;
    //        var placeholder = completion.TabStops[1];

    //        var code = string.Join("\r\n", completion.CodeBody);
    //        var initialSelection = new Selection(1, 1, 1, prefix.Length);

    //        var sut = new BlockCompletionService(TestBlockCompletions)
    //        {
    //            Current = completion
    //        };

    //        var (newCodeA, newSelectionA) = Run(Keys.Tab, code, initialSelection, sut);
    //        var (newCodeB, newSelectionB) = Run(Keys.Tab, newCodeA, newSelectionA, sut);

    //        Assert.AreEqual(placeholder.Position.ToOneBased(), newSelectionB);
    //    }

    //    [Test]
    //    public void GivenTabStops_TabOnLastPlaceholderCompletesBlock()
    //    {
    //        var completion = TestBlockCompletions[0];
    //        var prefix = completion.Prefix;
    //        var placeholder = completion.TabStops[1];

    //        var code = string.Join("\r\n", completion.CodeBody);
    //        var initialSelection = new Selection(1, 1, 1, prefix.Length);

    //        var sut = new BlockCompletionService(TestBlockCompletions)
    //        {
    //            Current = completion
    //        };

    //        for (int i = 0; i < completion.TabStops.Count; i++)
    //        {
    //            Assert.IsNotNull(sut.Current);
    //            Run(Keys.Tab, code, initialSelection);
    //        }

    //        Assert.IsNull(sut.Current);
    //    }

    //    [Test]
    //    public void GivenInputForPlaceholder_TabMovesToNextPlaceholder()
    //    {
    //        var completion = TestBlockCompletions[0];
    //        var prefix = completion.Prefix;
    //        var placeholder = completion.TabStops[0].Position;

    //        var code = string.Join("\r\n", completion.CodeBody);
    //        var initialSelection = placeholder.ToOneBased();
    //        var sut = new BlockCompletionService(TestBlockCompletions)
    //        {
    //            Current = completion
    //        };

    //        var (newCodeA, newSelectionA) = Run(Keys.Tab, code, initialSelection);
    //        Assert.AreEqual(placeholder.ToOneBased(), newSelectionA);

    //        var (newCodeB, newSelectionB) = Run(Keys.None, newCodeA, newSelectionA);
    //        Assert.AreEqual(placeholder.ToOneBased(), newSelectionB);

    //        var (newCode, newSelection) = Run(Keys.Tab, newCodeB, newSelectionB);
    //        Assert.AreEqual(completion.TabStops[1].Position.ToOneBased(), newSelection);
    //    }

    //    [Test]
    //    public void FirstTestBlockCompletionBody_HasThreePlaceholders()
    //    {
    //        Assert.IsTrue(TestBlockCompletions[0].TabStops.Count == 3);
    //    }

    //    private BlockCompletion[] TestBlockCompletions => new[] 
    //    {
    //        new BlockCompletion("Test", "foo", "Foo ${1:identifier} As ${2:something}$0".Split('\n'), "Test completion"),
    //    };
    //}
}
