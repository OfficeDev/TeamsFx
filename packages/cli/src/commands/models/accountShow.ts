// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
import { err, ok } from "@microsoft/teamsfx-api";
import { AppStudioScopes } from "@microsoft/teamsfx-core";
import { outputAccountInfoOffline, outputAzureInfo, outputM365Info } from "../../cmds/account";
import AzureTokenProvider from "../../commonlib/azureLogin";
import { checkIsOnline } from "../../commonlib/codeFlowLogin";
import { signedIn } from "../../commonlib/common/constant";
import { logger } from "../../commonlib/logger";
import M365TokenProvider from "../../commonlib/m365Login";
import CliTelemetry from "../../telemetry/cliTelemetry";
import { TelemetryEvent } from "../../telemetry/cliTelemetryEvents";
import { CLICommand, CLIContext } from "../types";

export const accountShowCommand: CLICommand = {
  name: "show",
  description: "Display all connected cloud accounts information.",
  telemetry: {
    event: TelemetryEvent.AccountShow,
  },
  handler: async (cmd: CLIContext) => {
    const m365StatusRes = await M365TokenProvider.getStatus({ scopes: AppStudioScopes });
    if (m365StatusRes.isErr()) {
      CliTelemetry.sendTelemetryErrorEvent(TelemetryEvent.AccountShow, m365StatusRes.error);
      return err(m365StatusRes.error);
    }
    const m365Status = m365StatusRes.value;
    if (m365Status.status === signedIn) {
      (await checkIsOnline())
        ? await outputM365Info("show")
        : await outputAccountInfoOffline("Microsoft 365", (m365Status.accountInfo as any).upn);
    }

    const azureStatus = await AzureTokenProvider.getStatus();
    if (azureStatus.status === signedIn) {
      (await checkIsOnline())
        ? await outputAzureInfo("show")
        : await outputAccountInfoOffline("Azure", (azureStatus.accountInfo as any).upn);
    }

    if (m365Status.status !== signedIn && azureStatus.status !== signedIn) {
      logger.info(
        "Use `teamsfx account login azure` or `teamsfx account login m365` to log in to Azure or Microsoft 365 account."
      );
    }
    return ok(undefined);
  },
};
