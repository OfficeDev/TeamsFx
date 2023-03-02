// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import "mocha";
import mockedEnv from "mocked-env";
import rewire from "rewire";
import fs from "fs-extra";
import chai from "chai";
import { stub, restore } from "sinon";
import { GeneratorChecker } from "../../../../../src/component/resource/spfx/depsChecker/generatorChecker";
import { telemetryHelper } from "../../../../../src/component/resource/spfx/utils/telemetry-helper";
import { Colors, LogLevel, LogProvider } from "@microsoft/teamsfx-api";
import { TestHelper } from "../helper";
import { cpUtils } from "../../../../../src/common/deps-checker/util/cpUtils";

const rGeneratorChecker = rewire(
  "../../../../../src/component/resource/spfx/depsChecker/generatorChecker"
);

class StubLogger implements LogProvider {
  async log(logLevel: LogLevel, message: string): Promise<boolean> {
    return true;
  }

  async trace(message: string): Promise<boolean> {
    return true;
  }

  async debug(message: string): Promise<boolean> {
    return true;
  }

  async info(message: string | Array<{ content: string; color: Colors }>): Promise<boolean> {
    return true;
  }

  async warning(message: string): Promise<boolean> {
    return true;
  }

  async error(message: string): Promise<boolean> {
    return true;
  }

  async fatal(message: string): Promise<boolean> {
    return true;
  }
}

describe("generator checker", () => {
  beforeEach(() => {
    stub(telemetryHelper, "sendSuccessEvent").callsFake(() => {
      console.log("success event");
      return;
    });
    stub(telemetryHelper, "sendErrorEvent").callsFake(() => {
      console.log("error event");
      return;
    });
  });

  afterEach(() => {
    restore();
  });

  describe("getDependencyInfo", async () => {
    it("Set SPFx version to 1.15", () => {
      const info = GeneratorChecker.getDependencyInfo();

      chai.expect(info).to.be.deep.equal({
        supportedVersion: "1.16.1",
        displayName: "@microsoft/generator-sharepoint@1.16.1",
      });
    });

    it("ensure deps - already installed", async () => {
      const generatorChecker = new GeneratorChecker(new StubLogger());
      const pluginContext = TestHelper.getFakePluginContext("test", "./", "");
      stub(generatorChecker, "isInstalled").callsFake(async () => {
        return true;
      });
      const result = await generatorChecker.ensureDependency(pluginContext);
      chai.expect(result.isOk()).is.true;
    });

    it("ensure deps - uninstalled", async () => {
      const generatorChecker = new GeneratorChecker(new StubLogger());
      const pluginContext = TestHelper.getFakePluginContext("test", "./", "");
      stub(generatorChecker, "isInstalled").callsFake(async () => {
        return false;
      });

      stub(generatorChecker, "install").throwsException(new Error());

      const result = await generatorChecker.ensureDependency(pluginContext);
      chai.expect(result.isOk()).is.false;
    });

    it("ensure deps -  install", async () => {
      const generatorChecker = new GeneratorChecker(new StubLogger());
      const pluginContext = TestHelper.getFakePluginContext("test", "./", "");
      stub(generatorChecker, "isInstalled").callsFake(async () => {
        return false;
      });

      stub(generatorChecker, "install");

      const result = await generatorChecker.ensureDependency(pluginContext);
      chai.expect(result.isOk()).is.true;
    });

    it("is installed", async () => {
      const generatorChecker = new GeneratorChecker(new StubLogger());
      stub(fs, "pathExists").callsFake(async () => {
        console.log("stub pathExists");
        return true;
      });

      stub(GeneratorChecker.prototype, <any>"queryVersion").callsFake(async () => {
        console.log("stub queryversion");
        return rGeneratorChecker.__get__("supportedVersion");
      });

      const result = await generatorChecker.isInstalled();
      chai.expect(result).is.true;
    });

    it("install", async () => {
      const generatorChecker = new GeneratorChecker(new StubLogger());
      const cleanStub = stub(GeneratorChecker.prototype, <any>"cleanup").callsFake(async () => {
        console.log("stub cleanup");
        return;
      });
      const installStub = stub(GeneratorChecker.prototype, <any>"installGenerator").callsFake(
        async () => {
          console.log("stub installyo");
          return;
        }
      );
      const validateStub = stub(GeneratorChecker.prototype, <any>"validate").callsFake(async () => {
        console.log("stub validate");
        return false;
      });

      try {
        await generatorChecker.install();
      } catch {
        chai.expect(installStub.callCount).equal(1);
        chai.expect(cleanStub.callCount).equal(2);
        chai.expect(validateStub.callCount).equal(1);
      }
    });

    it("findGloballyInstalledVersion: returns version", async () => {
      const generatorChecker = new GeneratorChecker(new StubLogger());
      stub(cpUtils, "executeCommand").resolves(
        "C:\\Roaming\\npm\n`-- @microsoft/generator-sharepoint@1.16.1\n\n"
      );

      const res = await generatorChecker.findGloballyInstalledVersion(1);
      chai.expect(res).equal("1.16.1");
    });

    it("findGloballyInstalledVersion: regex error", async () => {
      const generatorChecker = new GeneratorChecker(new StubLogger());
      stub(cpUtils, "executeCommand").resolves(
        "C:\\Roaming\\npm\n`-- @microsoft/generator-sharepoint@empty\n\n"
      );

      const res = await generatorChecker.findGloballyInstalledVersion(1);
      chai.expect(res).equal(undefined);
    });

    it("findGloballyInstalledVersion: exeute commmand error", async () => {
      const generatorChecker = new GeneratorChecker(new StubLogger());
      stub(cpUtils, "executeCommand").throws("run command error");
      let error = undefined;

      try {
        const res = await generatorChecker.findGloballyInstalledVersion(1);
      } catch (e) {
        error = e;
      }
      chai.expect(error).not.undefined;
    });

    it("findLatestVersion: returns version", async () => {
      const generatorChecker = new GeneratorChecker(new StubLogger());
      stub(cpUtils, "executeCommand").resolves("1.16.1");

      const res = await generatorChecker.findLatestVersion(1);
      chai.expect(res).equal("1.16.1");
    });

    it("findLatestVersion: regex error", async () => {
      const generatorChecker = new GeneratorChecker(new StubLogger());
      stub(cpUtils, "executeCommand").resolves("empty");

      const res = await generatorChecker.findLatestVersion(1);
      chai.expect(res).equal("latest");
    });

    it("findLatestVersion: exeute commmand error", async () => {
      const generatorChecker = new GeneratorChecker(new StubLogger());
      stub(cpUtils, "executeCommand").throws("run command error");

      const res = await generatorChecker.findLatestVersion();
      chai.expect(res).equal("latest");
    });
  });
});
