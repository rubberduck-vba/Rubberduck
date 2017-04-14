using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.UI;
using Rubberduck.VBEditor.SafeComWrappers;
using System.Diagnostics;

namespace Rubberduck.Refactorings.Rename
{
    public class RenameRefactorModule : RenameRefactorBase
    {
        public RenameRefactorModule(RubberduckParserState state)
            : base(state) { }

        public override string ErrorMessage => RubberduckUI.RenameDialog_ModuleRenameError;

        public override bool RequestParseAfterRename => false;

        public override void Rename(Declaration renameTarget, string newName)
        {
            var component = renameTarget.QualifiedName.QualifiedModuleName.Component;
            var module = component.CodeModule;
            Debug.Assert(!module.IsWrappingNullReference, "input validation fail: Code Module is wrapping a null reference");

            RenameUsages(renameTarget, newName);

            if (component.Type == ComponentType.Document)
            {
                var properties = component.Properties;
                var property = properties["_CodeName"];
                {
                    Rewrite();
                    property.Value = newName;
                }
            }
            else if (component.Type == ComponentType.UserForm)
            {
                var properties = component.Properties;
                var property = properties["Caption"];
                {
                    Rewrite();
                    if ((string)property.Value == renameTarget.IdentifierName)
                    {
                        property.Value = newName;
                    }
                    component.Name = newName;
                }
            }
            else
            {
                Rewrite();
                module.Name = newName;
            }
        }
    }
}
