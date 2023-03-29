import { UserError, UserErrorOptions } from "@microsoft/teamsfx-api";
import { getDefaultString, getLocalizedString } from "../common/localizeUtils";

/**
 * Failed to compile bicep into ARM template
 */
export class CompileBicepError extends UserError {
  constructor(filePath: string, error: Error) {
    const key = "error.arm.CompileBicepError";
    const errorOptions: UserErrorOptions = {
      source: "arm/deploy",
      name: "CompileBicepError",
      message: getDefaultString(key, filePath, error.message || ""),
      displayMessage: getLocalizedString(key, filePath, error.message || ""),
    };
    super(errorOptions);
  }
}

/**
 * Failed to deploy arm templates for some reason
 */
export class DeployArmError extends UserError {
  constructor(deployName: string, resourceGroup: string, error: Error) {
    const key = "error.arm.DeployArmError";
    const errorOptions: UserErrorOptions = {
      source: "arm/deploy",
      name: "DeployArmError",
      message: getDefaultString(key, deployName, resourceGroup, error.message || ""),
      displayMessage: getLocalizedString(key + ".Notification", deployName, resourceGroup),
    };
    super(errorOptions);
    if (error.stack) super.stack = error.stack;
  }
}

/**
 * Failed to deploy arm templates and get error message failed
 */
export class GetArmDeploymentError extends UserError {
  constructor(deployName: string, resourceGroup: string, deployError: Error, getError: Error) {
    const errorOptions: UserErrorOptions = {
      source: "arm/deploy",
      name: "GetArmDeploymentError",
      message: getDefaultString(
        "error.arm.GetArmDeploymentError",
        deployName,
        resourceGroup,
        deployError.message || "",
        getError.message || "",
        resourceGroup
      ),
      displayMessage: getLocalizedString(
        "error.arm.DeployArmError.Notification",
        deployName,
        resourceGroup
      ),
    };
    super(errorOptions);
  }
}

/**
 * Failed to convert ARM deployment result to action output
 */
export class ConvertArmOutputError extends UserError {
  constructor(outputKey: string) {
    const key = "error.arm.ConvertArmOutputError";
    const errorOptions: UserErrorOptions = {
      source: "arm/deploy",
      name: "ConvertArmOutputError",
      message: getDefaultString(key, outputKey),
      displayMessage: getLocalizedString(key, outputKey),
    };
    super(errorOptions);
  }
}
