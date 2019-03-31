using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using System.Collections.Generic;

// ReSharper disable once CheckNamespace - Special dispensation due to conflicting file vs namespace priorities
namespace Rubberduck.VBEditor.SafeComWrappers.VBA
{
    public class WordApp : HostApplicationBase<Microsoft.Office.Interop.Word.Application>
    {
        public WordApp(IVBE vbe) : base(vbe, "Word", true) { }

        public override IEnumerable<HostAutoMacro> AutoMacroIdentifiers => new HostAutoMacro[]
        {
            //ref: https://docs.microsoft.com/en-us/office/vba/word/concepts/customizing-word/auto-macros
            new HostAutoMacro(new[] {ComponentType.StandardModule, ComponentType.Document}, false, null, "AutoExec"),
            new HostAutoMacro(new[] {ComponentType.StandardModule, ComponentType.Document}, false, null, "AutoNew"),
            new HostAutoMacro(new[] {ComponentType.StandardModule, ComponentType.Document}, false, null, "AutoOpen"),
            new HostAutoMacro(new[] {ComponentType.StandardModule, ComponentType.Document}, false, null, "AutoClose"),
            new HostAutoMacro(new[] {ComponentType.StandardModule, ComponentType.Document}, false, null, "AutoExit"),
            new HostAutoMacro(new[] {ComponentType.StandardModule}, false, "AutoExec", "Main"),
            new HostAutoMacro(new[] {ComponentType.StandardModule}, false, "AutoNew", "Main"),
            new HostAutoMacro(new[] {ComponentType.StandardModule}, false, "AutoOpen", "Main"),
            new HostAutoMacro(new[] {ComponentType.StandardModule}, false, "AutoClose", "Main"),
            new HostAutoMacro(new[] {ComponentType.StandardModule}, false, "AutoExit", "Main")
        };
    }
}
