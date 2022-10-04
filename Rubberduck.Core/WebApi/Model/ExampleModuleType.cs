using System.ComponentModel;

namespace Rubberduck.Core.WebApi.Model
{
    public enum ExampleModuleType
    {
        None = 0,
        [Description("(Any)")] Any,
        [Description("Class Module")] ClassModule,
        [Description("Document Module")] DocumentModule,
        [Description("Interface Module")] InterfaceModule,
        [Description("Predeclared Class")] PredeclaredClass,
        [Description("Standard Module")] StandardModule,
        [Description("UserForm Module")] UserFormModule
    }
}
