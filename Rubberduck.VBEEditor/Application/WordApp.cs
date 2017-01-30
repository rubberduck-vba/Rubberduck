using System.Linq;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.VBEditor.Application
{
    public class WordApp : HostApplicationBase<Microsoft.Office.Interop.Word.Application>
    {
        public WordApp() : base("Word") { }
        public WordApp(IVBE vbe) : base(vbe, "Word") { }

        public override void Run(dynamic declaration)
        {
            var qualifiedMemberName = declaration.QualifiedName;
            var call = GenerateMethodCall(qualifiedMemberName);

            ActivateProjectDocument(declaration);
            //Prevent TE hanging should Application.Run fail
            try
            {
                Application.Run(call);
            }
            catch 
            { 
                //TODO - Let TestEngine know that the method failed 
            }
        }

        protected virtual string GenerateMethodCall(QualifiedMemberName qualifiedMemberName)
        {
            var moduleName = qualifiedMemberName.QualifiedModuleName.ComponentName;
            return string.Concat(moduleName, ".", qualifiedMemberName.MemberName);
        }

        protected virtual void ActivateProjectDocument(dynamic declaration)
        {
            // Word requires that the document be active for Application.Run to find the target Method in scope.
            // Check the project's document or a document referring to a project's template is active.
            var activeDoc = Application.ActiveDocument;
            //var template = activeDoc.get_AttachedTemplate();
            var targetDoc = Application.Documents.Cast<Microsoft.Office.Interop.Word.Document>()
                    .FirstOrDefault(doc => doc.Name == declaration.ProjectDisplayName
                                        || doc.get_AttachedTemplate().Name == declaration.ProjectDisplayName);
            if (activeDoc != targetDoc && targetDoc != null) 
            { 
                targetDoc.Activate();
            }
        }
    }
}
