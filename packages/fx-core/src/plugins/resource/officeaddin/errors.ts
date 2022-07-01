import { SystemError } from "@microsoft/teamsfx-api";

const source = "office-addin";

export function UndefinedProjectPathError(): SystemError {
  return new SystemError(
    source,
    "InvalidProjectPath",
    "Project path is undefined",
    "Project path is undefined"
  );
}
