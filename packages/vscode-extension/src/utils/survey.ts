import * as vscode from "vscode";
import { globalStateGet, globalStateUpdate, isValidProject } from "@microsoft/teamsfx-core";
import * as ExtTelemetry from "../telemetry/telemetry";
import { TelemetryEvent } from "../telemetry/extTelemetryEvents";
import { TreatmentVariableValue } from "../exp/treatmentVariables";
import * as globalVariables from "../globalVariables";
import { getDefaultString, localize } from "./localizeUtils";

const SURVEY_URL = "https://aka.ms/teams-toolkit-survey";

enum ExtensionSurveyStateKeys {
  DoNotShowAgain = "survey/doNotShowAgain",
  RemindMeLater = "survey/remindMeLater",
  DisableSurveyForTime = "survey/disableSurveyForTime",
}

const TIME_TO_DISABLE_SURVEY = 1000 * 60 * 60 * 24 * 7 * 12; // 12 weeks
const TIME_TO_REMIND_ME_LATER = 1000 * 60 * 60 * 24 * 7 * 2; // 2 weeks
const TIME_TO_SHOW_SURVEY = 1000 * 60 * 15; // 15 minutes
const SAMPLE_PERCENTAGE = 25; // 25 percent for public preview

export class ExtensionSurvey {
  private timeToShowSurvey: number;
  private timeToDisableSurvey: number;
  private timeToRemindMeLater: number;
  private checkSurveyInterval?: NodeJS.Timeout;
  private showSurveyTimeout?: NodeJS.Timeout;
  private needToShow = false;
  private static instance: ExtensionSurvey;

  public static getInstance(): ExtensionSurvey {
    if (!ExtensionSurvey.instance) {
      ExtensionSurvey.instance = new ExtensionSurvey();
    }

    return ExtensionSurvey.instance;
  }

  private constructor(
    timeToShowSurvey?: number,
    samplePercentage?: number,
    timeToDisableSurvey?: number,
    timeToRemindMeLater?: number
  ) {
    this.timeToShowSurvey = timeToShowSurvey ? timeToShowSurvey : TIME_TO_SHOW_SURVEY;

    const randomSample: number = Math.floor(Math.random() * 100) + 1;
    if (randomSample <= (samplePercentage ? samplePercentage : SAMPLE_PERCENTAGE)) {
      this.needToShow = true;
    }
    this.timeToDisableSurvey = timeToDisableSurvey ? timeToDisableSurvey : TIME_TO_DISABLE_SURVEY;
    this.timeToRemindMeLater = timeToRemindMeLater ? timeToRemindMeLater : TIME_TO_REMIND_ME_LATER;
  }

  public async activate(): Promise<void> {
    if (TreatmentVariableValue.isEmbeddedSurvey) {
      if (this.needToShow) {
        if (!(await this.shouldShowBanner())) {
          return;
        }

        setTimeout(() => {
          this.showSurvey();
        }, 200);
      }
    } else {
      if (this.needToShow && !this.checkSurveyInterval) {
        this.checkSurveyInterval = setInterval(async () => {
          if (!(await this.shouldShowBanner())) {
            return;
          }

          if (!this.showSurveyTimeout && isValidProject(globalVariables.workspaceUri!.fsPath)) {
            this.showSurveyTimeout = setTimeout(() => this.showSurvey(), this.timeToShowSurvey);
          }
        }, 2000);
      }
    }
  }

  private async shouldShowBanner(): Promise<boolean> {
    const doNotShowAgain = await globalStateGet(ExtensionSurveyStateKeys.DoNotShowAgain, false);
    if (doNotShowAgain) {
      return false;
    }

    const currentTime = Date.now();
    const remindMeLaterTime = await globalStateGet(ExtensionSurveyStateKeys.RemindMeLater, 0);
    if (remindMeLaterTime > currentTime) {
      return false;
    }

    const disableSurveyForTime = await globalStateGet(
      ExtensionSurveyStateKeys.DisableSurveyForTime,
      0
    );
    if (disableSurveyForTime > currentTime) {
      return false;
    }

    return true;
  }

  private async showWebviewSurvey(): Promise<void> {
    vscode.commands.executeCommand("fx-extension.openSurvey");
  }

  private async showSurvey(): Promise<void> {
    const extension = vscode.extensions.getExtension("TeamsDevApp.ms-teams-vscode-extension");
    if (!extension) {
      return;
    }

    const extensionVersion = extension.packageJSON.version || "unknown";
    const take = {
      title: localize("teamstoolkit.survey.takeSurvey.title"),
      run: async (): Promise<void> => {
        ExtTelemetry.sendTelemetryEvent(TelemetryEvent.Survey, {
          message: getDefaultString("teamstoolkit.survey.takeSurvey.message"),
        });

        if (TreatmentVariableValue.isEmbeddedSurvey) {
          this.showWebviewSurvey();
        } else {
          vscode.commands.executeCommand(
            "vscode.open",
            vscode.Uri.parse(
              `${SURVEY_URL}?o=${encodeURIComponent(process.platform)}&v=${encodeURIComponent(
                extensionVersion
              )}`
            )
          );
        }

        const disableSurveyForTime = Date.now() + this.timeToDisableSurvey;
        await globalStateUpdate(
          ExtensionSurveyStateKeys.DisableSurveyForTime,
          disableSurveyForTime
        );
      },
    };

    const remind = {
      title: localize("teamstoolkit.survey.remindMeLater.title"),
      run: async (): Promise<void> => {
        ExtTelemetry.sendTelemetryEvent(TelemetryEvent.Survey, {
          message: getDefaultString("teamstoolkit.survey.remindMeLater.message"),
        });
        const disableSurveyForTime = Date.now() + this.timeToRemindMeLater;
        await globalStateUpdate(ExtensionSurveyStateKeys.RemindMeLater, disableSurveyForTime);
      },
    };

    const never = {
      title: localize("teamstoolkit.survey.dontShowAgain.title"),
      run: async (): Promise<void> => {
        ExtTelemetry.sendTelemetryEvent(TelemetryEvent.Survey, {
          message: getDefaultString("teamstoolkit.survey.dontShowAgain.message"),
        });
        await globalStateUpdate(ExtensionSurveyStateKeys.DoNotShowAgain, true);
      },
    };

    const selection = await vscode.window.showInformationMessage(
      localize("teamstoolkit.survey.banner.title"),
      take,
      remind,
      never
    );

    if (this.showSurveyTimeout) {
      clearTimeout(this.showSurveyTimeout);
      this.showSurveyTimeout = undefined;
    }

    if (selection) {
      ExtTelemetry.sendTelemetryEvent(TelemetryEvent.Survey, {
        message: getDefaultString("teamstoolkit.survey.banner.message"),
      });
      await selection.run();
    } else {
      ExtTelemetry.sendTelemetryEvent(TelemetryEvent.Survey, {
        message: getDefaultString("teamstoolkit.survey.cancelMessage"),
      });
      const disableSurveyForTime = Date.now() + this.timeToRemindMeLater;
      await globalStateUpdate(ExtensionSurveyStateKeys.RemindMeLater, disableSurveyForTime);
    }

    // only pop out once in one session
    this.needToShow = false;
  }
}
