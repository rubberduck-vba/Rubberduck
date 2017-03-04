namespace Rubberduck.SmartIndenter
{
    public interface IIndenterSettings
    {
        bool IndentEntireProcedureBody { get; set; }
        bool IndentFirstCommentBlock { get; set; }
        bool IndentEnumTypeAsProcedure { get; set; }
        bool IndentFirstDeclarationBlock { get; set; }
        bool AlignCommentsWithCode { get; set; }
        bool AlignContinuations { get; set; }
        bool IgnoreOperatorsInContinuations { get; set; }
        bool IndentCase { get; set; }
        bool ForceDebugStatementsInColumn1 { get; set; }
        bool ForceCompilerDirectivesInColumn1 { get; set; }
        bool IndentCompilerDirectives { get; set; }
        bool AlignDims { get; set; }
        int AlignDimColumn { get; set; }
        EndOfLineCommentStyle EndOfLineCommentStyle { get; set; }
        int EndOfLineCommentColumnSpaceAlignment { get; set; }
        int IndentSpaces { get; set; }
        bool VerticallySpaceProcedures { get; set; }
        int LinesBetweenProcedures { get; set; }
        bool LegacySettingsExist();
        void LoadLegacyFromRegistry();
    }

    public enum EndOfLineCommentStyle
    {
        Absolute,
        SameGap,
        StandardGap,
        AlignInColumn
    }
}
