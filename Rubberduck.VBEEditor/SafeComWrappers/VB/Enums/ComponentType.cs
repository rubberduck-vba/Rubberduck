namespace Rubberduck.VBEditor.SafeComWrappers
{
    // Note, values 4-10 (inclusive) are available in VB6 only
    public enum ComponentType
    {
        ComComponent = -1,
        Undefined = 0,
        StandardModule = 1,
        ClassModule = 2,
        UserForm = 3,
        ResFile = 4,
        VBForm = 5,
        MDIForm = 6,
        PropPage = 7,
        UserControl = 8,
        DocObject = 9,
        RelatedDocument = 10,
        ActiveXDesigner = 11,
        Document = 100       
    }
}