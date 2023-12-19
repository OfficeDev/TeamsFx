// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
"use strict";

import { OpenAPIV3 } from "openapi-types";
import fs from "fs-extra";
import path from "path";
import { ErrorType, WarningResult } from "./interfaces";
import {
  getSafeRegistrationIdEnvName,
  isAPIKeyAuth,
  isBearerTokenAuth,
  parseApiInfo,
} from "./utils";
import { SpecParserError } from "./specParserError";
import { ConstantString } from "./constants";
import {
  IComposeExtension,
  IMessagingExtensionCommand,
  TeamsAppManifest,
} from "@microsoft/teamsfx-api";

export async function updateManifest(
  manifestPath: string,
  outputSpecPath: string,
  adaptiveCardFolder: string,
  spec: OpenAPIV3.Document,
  allowMultipleParameters: boolean,
  auth?: OpenAPIV3.SecuritySchemeObject
): Promise<[TeamsAppManifest, WarningResult[]]> {
  try {
    const originalManifest: TeamsAppManifest = await fs.readJSON(manifestPath);
    const updatedPart: any = {};
    const [commands, warnings] = await generateCommands(
      spec,
      adaptiveCardFolder,
      manifestPath,
      allowMultipleParameters
    );
    const composeExtension: IComposeExtension = {
      composeExtensionType: "apiBased",
      apiSpecificationFile: getRelativePath(manifestPath, outputSpecPath),
      commands: commands,
    };

    if (auth) {
      if (isAPIKeyAuth(auth)) {
        auth = auth as OpenAPIV3.ApiKeySecurityScheme;
        const safeApiSecretRegistrationId = getSafeRegistrationIdEnvName(
          `${auth.name}_${ConstantString.RegistrationIdPostfix}`
        );
        (composeExtension as any).authorization = {
          authType: "apiSecretServiceAuth",
          apiSecretServiceAuthConfiguration: {
            apiSecretRegistrationId: `\${{${safeApiSecretRegistrationId}}}`,
          },
        };
      } else if (isBearerTokenAuth(auth)) {
        (composeExtension as any).authorization = {
          authType: "microsoftEntra",
          microsoftEntraConfiguration: {
            supportsSingleSignOn: true,
          },
        };

        updatedPart.webApplicationInfo = {
          id: "${{AAD_APP_CLIENT_ID}}",
          resource: "api://${{DOMAIN}}/${{AAD_APP_CLIENT_ID}}",
        };
      }
    }

    updatedPart.description = {
      short: spec.info.title.slice(0, ConstantString.ShortDescriptionMaxLens),
      full: (spec.info.description ?? originalManifest.description.full)?.slice(
        0,
        ConstantString.FullDescriptionMaxLens
      ),
    };

    updatedPart.composeExtensions = [composeExtension];

    const updatedManifest = { ...originalManifest, ...updatedPart };

    return [updatedManifest, warnings];
  } catch (err) {
    throw new SpecParserError((err as Error).toString(), ErrorType.UpdateManifestFailed);
  }
}

export async function generateCommands(
  spec: OpenAPIV3.Document,
  adaptiveCardFolder: string,
  manifestPath: string,
  allowMultipleParameters: boolean
): Promise<[IMessagingExtensionCommand[], WarningResult[]]> {
  const paths = spec.paths;
  const commands: IMessagingExtensionCommand[] = [];
  const warnings: WarningResult[] = [];
  if (paths) {
    for (const pathUrl in paths) {
      const pathItem = paths[pathUrl];
      if (pathItem) {
        const operations = pathItem;

        // Currently only support GET and POST method
        for (const method in operations) {
          if (ConstantString.SupportedMethods.includes(method)) {
            const operationItem = (operations as any)[method] as OpenAPIV3.OperationObject;
            if (operationItem) {
              const [command, warning] = parseApiInfo(operationItem, allowMultipleParameters);

              const adaptiveCardPath = path.join(adaptiveCardFolder, command.id + ".json");
              command.apiResponseRenderingTemplateFile = (await fs.pathExists(adaptiveCardPath))
                ? getRelativePath(manifestPath, adaptiveCardPath)
                : "";

              if (warning) {
                warnings.push(warning);
              }

              commands.push(command);
            }
          }
        }
      }
    }
  }

  return [commands, warnings];
}

export function getRelativePath(from: string, to: string): string {
  const relativePath = path.relative(path.dirname(from), to);
  return path.normalize(relativePath).replace(/\\/g, "/");
}
