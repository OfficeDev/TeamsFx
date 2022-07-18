export class TreatmentVariables {
  public static readonly VSCodeConfig = "vscode";
  public static readonly EmbeddedSurvey = "embeddedsurvey";
  public static readonly CustomizeTreeview = "customizetreeview";
  public static readonly UseFolderSelection = "usefolderselection";
}

export class TreatmentVariableValue {
  public static isEmbeddedSurvey: boolean | undefined = undefined;
  // If this is true, users can choose default folder or a customized folder when creating a new project.
  public static useFolderSelection: boolean | undefined = undefined;
}
