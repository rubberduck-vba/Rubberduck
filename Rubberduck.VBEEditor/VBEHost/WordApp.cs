using Microsoft.Vbe.Interop;
using Microsoft.Office.Interop.Word;
using System.Linq;

namespace Rubberduck.VBEditor.VBEHost
{
    public class WordApp : HostApplicationBase<Microsoft.Office.Interop.Word.Application>
    {
        public WordApp() : base("Word") { }
        public WordApp(VBE vbe) : base(vbe, "Word") { }

        public override void Run(QualifiedMemberName qualifiedMemberName)
        {
            var call = GenerateMethodCall(qualifiedMemberName);

            ActivateProjectDocument(qualifiedMemberName);
            //Prevent TE hanging should Application.Run fail
            try
            {
                Application.Run(call);
            }
            //TODO - Let TestEngine know that the method failed
            catch { };
        }

        protected virtual string GenerateMethodCall(QualifiedMemberName qualifiedMemberName)
        {
            var moduleName = qualifiedMemberName.QualifiedModuleName.Component.Name;
            return string.Concat(moduleName, ".", qualifiedMemberName.MemberName);
        }

        protected virtual void ActivateProjectDocument(QualifiedMemberName qualifiedMemberName)
        {
            // Word requires that the document be active for Application.Run to find the target Method in scope.
            // Check the project's document or a document referring to a project's template is active.
            var activeDoc = Application.ActiveDocument;
            var template = activeDoc.get_AttachedTemplate();
            var targetDoc = Application.Documents.Cast<Document>()
                    .FirstOrDefault(doc => doc.Name == qualifiedMemberName.QualifiedModuleName.ProjectDisplayName
                                        || doc.get_AttachedTemplate().Name == qualifiedMemberName.QualifiedModuleName.ProjectDisplayName);
            if (activeDoc != targetDoc) { 
                targetDoc.Activate();
            }
        }
    }
}
