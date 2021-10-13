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
    /// MSForms exposes UserForm controls as public fields; accessing these fields outside the UserForm class breaks encapsulation and couples
    /// the application logic with specific form controls rather than the data they hold.  
    /// For a more object-oriented approach and code that can be unit-tested, consider encapsulating the desired values into their own 'model' class,
    /// making event handlers in the form manipulate these 'model' properties, then have the code that displayed the form query this encapsulated state as needed.
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
    /// </example>
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
    /// ' simple solution: embed the model in the form itself, expose a getter procedure for each desired property.
    /// ' > pros: simple to implement, silences the inspection!
    /// ' > cons: view vs model responsibilities are fuzzy, intellisense get bloated, business logic is still coupled with the UI.
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
    /// <module name="TestModel" type="Class Module">
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
    /// ' MVP solution: encapsulate the model data into its own data type.
    /// ' > pros: easily extended, cleanly separates data from presentation concerns; application logic can be tested independently of the form.
    /// ' > cons: Model-View-Presenter architecture requires more modules and can feel/be "overkill" for simpler scenarios.
    /// Option Explicit
    /// Private Type TView
    ///     Model As TestModel
    /// End Type
    /// Private This As TView
    ///
    /// '@Description "Gets or sets Model object for this instance."
    /// Public Property Get Model() As TestModel
    ///     Set Model = This.Model
    /// End Property
    ///
    /// Public Property Set Model(ByVal RHS As TestModel)
    ///     Set This.Model = RHS
    /// End Property
    ///
    /// Private Sub ExportPathBox_Change()
    ///     ' the export path has changed; update the model accordingly
    ///     Model.ExportPath = ExportPathBox.Text
    /// End Sub
    ///
    /// Private Sub FileNameBox_Change()
    ///     ' the file name has changed; update the model accordingly
    ///     Model.FileName = FileNameBox.Text
    /// End Sub
    ///
    /// '...
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