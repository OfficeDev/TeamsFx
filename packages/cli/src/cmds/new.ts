// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

"use strict";

import AdmZip from "adm-zip";
import fs from "fs-extra";
import path from "path";
import { Argv, Options } from "yargs";

import {
  FxError,
  err,
  ok,
  Result,
  Question,
  LogLevel,
  Stage,
  OptionItem,
} from "@microsoft/teamsfx-api";
import {
  downloadSampleHook,
  fetchCodeZip,
  sampleProvider,
  saveFilesRecursively,
} from "@microsoft/teamsfx-core";

import activate from "../activate";
import * as constants from "../constants";
import { NotFoundInputedFolder, SampleAppDownloadFailed, ProjectFolderExist } from "../error";
import { YargsCommand } from "../yargsCommand";
import { getSystemInputs, toLocaleLowerCase } from "../utils";
import CliTelemetry from "../telemetry/cliTelemetry";
import {
  TelemetryEvent,
  TelemetryProperty,
  TelemetrySuccess,
} from "../telemetry/cliTelemetryEvents";
import CLIUIInstance from "../userInteraction";
import CLILogProvider from "../commonlib/log";
import HelpParamGenerator from "../helpParamGenerator";

export default class New extends YargsCommand {
  public readonly commandHead = `new`;
  public readonly command = `${this.commandHead}`;
  public readonly description = "Create a new Teams application.";
  public params: { [_: string]: Options } = {};

  public readonly subCommands: YargsCommand[] = [new NewTemplete()];

  public builder(yargs: Argv): Argv<any> {
    this.params = HelpParamGenerator.getYargsParamForHelp(Stage.create);
    this.subCommands.forEach((cmd) => {
      yargs.command(cmd.command, cmd.description, cmd.builder.bind(cmd), cmd.handler.bind(cmd));
    });
    if (this.params) {
      yargs
        .options({
          interactive: {
            description: "Select the options interactively",
            boolean: true,
            default: true,
            global: false,
          },
        })
        .options(this.params);
    }
    return yargs.version(false);
  }

  public async runCommand(args: {
    [argName: string]: string | string[];
  }): Promise<Result<null, FxError>> {
    CliTelemetry.sendTelemetryEvent(TelemetryEvent.CreateProjectStart);

    if (!args.interactive) {
      CLIUIInstance.updatePresetAnswers(this.params, args);
    }

    const result = await activate();
    if (result.isErr()) {
      CliTelemetry.sendTelemetryErrorEvent(TelemetryEvent.CreateProject, result.error);
      return err(result.error);
    }

    const core = result.value;

    {
      const result = await core.createProject(getSystemInputs());
      if (result.isErr()) {
        CliTelemetry.sendTelemetryErrorEvent(TelemetryEvent.CreateProject, result.error);
        return err(result.error);
      }
    }

    CliTelemetry.sendTelemetryEvent(TelemetryEvent.CreateProject, {
      [TelemetryProperty.Success]: TelemetrySuccess.Yes,
    });
    return ok(null);
  }
}

class NewTemplete extends YargsCommand {
  public readonly commandHead = `template`;
  public readonly command = `${this.commandHead} <template-name>`;
  public readonly description = "Create an app from an existing template.";

  public readonly subCommands: YargsCommand[] = [new NewTempleteList()];

  public builder(yargs: Argv): Argv<any> {
    const RootFolderNodeData = constants.RootFolderNode.data as Question;
    this.subCommands.forEach((cmd) => {
      yargs.command(cmd.command, cmd.description, cmd.builder.bind(cmd), cmd.handler.bind(cmd));
    });
    const templatesNames = constants.templates.map((t) => toLocaleLowerCase(t.sampleAppName));
    yargs
      .positional("template-name", {
        description: "Enter the template name",
        type: "string",
        choices: templatesNames,
        default: templatesNames[0],
        coerce: toLocaleLowerCase,
      })
      .options(RootFolderNodeData.name, {
        type: "string",
        description: RootFolderNodeData.type != "func" ? RootFolderNodeData.title : "unknown",
        default: RootFolderNodeData.default,
      });
    return yargs;
  }

  public async runCommand(args: {
    [argName: string]: string | string[];
  }): Promise<Result<null, FxError>> {
    const folder = path.resolve((args.folder as string) || "./");
    if (!fs.pathExistsSync(folder)) {
      CliTelemetry.sendTelemetryErrorEvent(
        TelemetryEvent.DownloadSample,
        NotFoundInputedFolder(folder)
      );
      return err(NotFoundInputedFolder(folder));
    }
    CliTelemetry.sendTelemetryEvent(TelemetryEvent.DownloadSampleStart);

    const result = await activate();
    if (result.isErr()) {
      CliTelemetry.sendTelemetryErrorEvent(TelemetryEvent.CreateProject, result.error);
      return err(result.error);
    }

    const core = result.value;

    const templateName = args["template-name"] as string;
    const template = constants.templates.find(
      (t) => toLocaleLowerCase(t.sampleAppName) === templateName
    )!;

    const inputs = getSystemInputs();
    inputs["scratch"] = "no";
    const options = sampleProvider.SampleCollection.samples
      .filter((sample) => sample.id === templateName)
      .map((sample) => {
        return {
          id: sample.id,
          label: sample.title,
          description: sample.shortDescription,
          data: sample.link,
        } as OptionItem;
      });
    inputs["samples"] = options[0];
    inputs["folder"] = folder;
    const result = await core.createProject(inputs);
    if (result.isErr()) {
      CliTelemetry.sendTelemetryErrorEvent(TelemetryEvent.DownloadSample, result.error);
      return err(result.error);
    }

    const sampleAppFolder = path.resolve(folder, template.sampleAppName);
    // if ((await fs.pathExists(sampleAppFolder)) && (await fs.readdir(sampleAppFolder)).length > 0) {
    //   CliTelemetry.sendTelemetryErrorEvent(
    //     TelemetryEvent.DownloadSample,
    //     ProjectFolderExist(sampleAppFolder)
    //   );
    //   return err(ProjectFolderExist(sampleAppFolder));
    // }

    // const result = await fetchCodeZip(template.sampleAppUrl);
    // if (!result) {
    //   throw SampleAppDownloadFailed(template.sampleAppUrl, new Error());
    // }
    // await saveFilesRecursively(new AdmZip(result.data), template.sampleAppName, folder);
    // await downloadSampleHook(templateName, sampleAppFolder);
    CLILogProvider.necessaryLog(
      LogLevel.Info,
      `Downloaded the '${CLILogProvider.white(
        template.sampleAppName
      )}' sample to '${CLILogProvider.white(sampleAppFolder)}'.`
    );

    CliTelemetry.sendTelemetryEvent(TelemetryEvent.DownloadSample, {
      [TelemetryProperty.Success]: TelemetrySuccess.Yes,
      [TelemetryProperty.SampleName]: templateName,
    });
    return ok(null);
  }
}

class NewTempleteList extends YargsCommand {
  public readonly commandHead = `list`;
  public readonly command = `${this.commandHead}`;
  public readonly description = "List all templates";

  public builder(yargs: Argv): Argv<any> {
    return yargs.hide("template-name");
  }

  public async runCommand(args: {
    [argName: string]: string | string[];
  }): Promise<Result<null, FxError>> {
    CLILogProvider.necessaryLog(LogLevel.Info, `The following are sample apps:`);
    CLILogProvider.necessaryLog(
      LogLevel.Info,
      JSON.stringify(constants.templates, undefined, 4),
      true
    );
    CLILogProvider.necessaryLog(
      LogLevel.Info,
      `Use the command ${CLILogProvider.white(
        "teamsfx new template <sampleAppName>"
      )} to create an application from the sample app.`
    );
    return ok(null);
  }
}
