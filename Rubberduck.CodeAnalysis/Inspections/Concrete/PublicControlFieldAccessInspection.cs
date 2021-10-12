using Rubberduck.CodeAnalysis.Inspections.Abstract;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Parsing.VBA.DeclarationCaching;
using Rubberduck.Resources.Inspections;

namespace Rubberduck.CodeAnalysis.Inspections.Concrete
{
    /// <summary>
    /// Flags MSForms controls being accessed from outside the UserForm that contains them.
    /// </summary>
    /// <why>
    /// MSForms exposes UserForm controls as public fields; accessing these fields outside the UserForm class breaks encapsulation and needlessly couples code with specific form controls.
    /// Consider encapsulating the desired values into their own 'model' class, making event handlers in the form manipulate these 'model' properties, and then the calling code can query this encapsulated state instead of querying form controls.
    /// </why>
    /// <example hasResult="true">
    /// <module name="Module1" type="Standard Module">
    /// <![CDATA[
    /// Option Explicit
    /// 
    /// Public Sub Test()
    ///     With New UserForm1
    ///         .Show
    ///         If .ExportPathBox.Text <> vbNullString Then
    ///             MsgBox .FileNameBox.Text
    ///         End If
    ///     End With
    /// End Sub
    /// ]]>
    /// </module>
    /// <example hasResult="false">
    /// <module name="Module1" type="Standard Module">
    /// <![CDATA[
    /// Option Explicit
    /// 
    /// Public Sub Test()
    ///     With New UserForm1
    ///         .Show
    ///         If .ExportPath <> vbNullString Then
    ///             MsgBox .FileName
    ///         End If
    ///     End With
    /// End Sub
    /// ]]>
    /// </module>
    /// <module name="UserForm1" type="UserForm Module">
    /// <![CDATA[
    /// ' simple solution: embed the model in the form itself, and expose getters for each desired property.
    /// ' > pros: simple to implement.
    /// ' > cons: view vs model responsibilities are fuzzy.
    /// Option Explicit
    /// 
    /// Public Property Get ExportPath() As String
    ///     ExportPath = ExportPathBox.Text
    /// End Property
    ///
    /// Public Property Get FileName() As String
    ///     FileName = FileNameBox.Text
    /// End Property
    /// ]]>
    /// </module>
    /// </example>
    /// </example>
    /// <example hasResult="false">
    /// <module name="Module1" type="Standard Module">
    /// <![CDATA[
    /// Option Explicit
    /// 
    /// Public Sub Test()
    ///     Dim Model As TestModel
    ///     Set Model = New TestModel
    ///     With New UserForm1
    ///         Set .Model = Model
    ///         .Show
    ///         If Model.ExportPath <> vbNullString Then
    ///             MsgBox Model.FileName
    ///         End If
    ///     End With
    /// End Sub
    /// ]]>
    /// </module>
    /// <module name="TestMmodel" type="Class Module">
    /// <![CDATA[
    /// Option Explicit
    /// Private Type TModel
    ///     ExportPath As String
    ///     FileName As String
    /// End Type
    /// Private This As TModel
    ///
    /// Public Property Get ExportPath() As String
    ///     ExportPath = This.ExportPath
    /// End Property
    ///
    /// Public Property Let ExportPath(ByVal RHS As String)
    ///     This.ExportPath = RHS
    /// End Property
    ///
    /// Public Property Get FileName() As String
    ///     FileName = This.FileName
    /// End Property
    ///
    /// Public Property Let FileName(ByVal RHS As String)
    ///     This.FileName = RHS
    /// End Property
    /// ]]>
    /// </module>
    /// <module name="UserForm1" type="UserForm Module">
    /// <![CDATA[
    /// ' thorough solution: encapsulate the model data into its own data type.
    /// ' > pros: easily extended, cleanly separates data from presentation concerns.
    /// ' > cons: model-view-presenter architecture requires more modules and can feel "overkill" for simpler scenarios.
    /// Option Explicit
    /// Private Type TView
    ///     Model As TestModel
    /// End Type
    /// Private This As TView
    ///
    /// Public Property Get Model() As TestModel
    ///     Set Model = This.Model
    /// End Property
    ///
    /// Public Property Set Model(ByVal RHS As TestModel)
    ///     Set This.Model = RHS
    /// End Property
    ///
    /// Private Sub ExportPathBox_Change()
    ///     Model.ExportPath = ExportPathBox.Text
    /// End Sub
    ///
    /// Private Sub FileNameBox_Change()
    ///     Model.FileName = FileNameBox.Text
    /// End Sub
    /// ]]>
    /// </module>
    /// </example>
    internal sealed class PublicControlFieldAccessInspection : IdentifierReferenceInspectionBase
    {
        public PublicControlFieldAccessInspection(IDeclarationFinderProvider declarationFinderProvider)
            : base(declarationFinderProvider)
        {
        }

        protected override bool IsResultReference(IdentifierReference reference, DeclarationFinder finder)
        {
            return reference.Declaration.DeclarationType == DeclarationType.Control &&
                   !reference.ParentScoping.ParentDeclaration.Equals(reference.Declaration.ParentDeclaration);
        }

        protected override string ResultDescription(IdentifierReference reference)
        {
            return string.Format(InspectionResults.PublicControlFieldAccessInspection, reference.Declaration.ParentDeclaration.IdentifierName, reference.IdentifierName);
        }
    }
}