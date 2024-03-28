// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * @author yefuwang@microsoft.com
 */

import {
  Context,
  devPreview,
  err,
  Inputs,
  ManifestUtil,
  ok,
  Platform,
  SystemError,
} from "@microsoft/teamsfx-api";
import * as chai from "chai";
import * as childProcess from "child_process";
import EventEmitter from "events";
import fs from "fs";
import fse from "fs-extra";
import "mocha";
import mockfs from "mock-fs";
import mockedEnv, { RestoreFn } from "mocked-env";
import * as fetch from "node-fetch";
import { OfficeAddinManifest } from "office-addin-manifest";
import * as path from "path";
import proxyquire from "proxyquire";
import * as sinon from "sinon";
import * as unzip from "unzipper";
import * as uuid from "uuid";
import { cpUtils } from "../../../src/common/deps-checker";
import { manifestUtils } from "../../../src/component/driver/teamsApp/utils/ManifestUtils";
import { Generator } from "../../../src/component/generator/generator";
import {
  OfficeAddinGenerator,
  getHost,
} from "../../../src/component/generator/officeAddin/generator";
import {
  HelperMethods,
  unzipErrorHandler,
} from "../../../src/component/generator/officeAddin/helperMethods";
import { createContextV3 } from "../../../src/component/utils";
import { setTools } from "../../../src/core/globalVars";
import { AccessGithubError, UserCancelError } from "../../../src/error";
import { ProjectTypeOptions, QuestionNames } from "../../../src/question";
import { MockTools } from "../../core/utils";

describe("OfficeAddinGenerator for Outlook Addin", function () {
  const testFolder = path.resolve("./tmp");
  let context: Context;
  let mockedEnvRestore: RestoreFn;
  const mockedError = new SystemError("mockedSource", "mockedError", "mockedMessage");

  beforeEach(async () => {
    mockedEnvRestore = mockedEnv({ TEAMSFX_V3: "true" }, { clear: true });
    const gtools = new MockTools();
    setTools(gtools);
    context = createContextV3();

    await fse.ensureDir(testFolder);
    sinon.stub(fs, "stat").resolves();
    sinon.stub(cpUtils, "executeCommand").resolves("succeed");
    const manifestId = uuid.v4();
    sinon.stub(fs, "readFile").resolves(new Buffer(`{"id": "${manifestId}"}`));
    sinon.stub(fs, "writeFile").resolves();
    sinon.stub(fs, "rename").resolves();
    sinon.stub(fs, "copyFile").resolves();
    sinon.stub(fse, "remove").resolves();
    sinon.stub(fse, "readJson").resolves({});
    sinon.stub(fse, "ensureFile").resolves();
    sinon.stub(fse, "writeJSON").resolves();
  });

  it("should run childProcessExec command success", async function () {
    sinon.stub(childProcess, "exec").yields(`echo 'test'`, "test");
    chai.assert(await OfficeAddinGenerator.childProcessExec(`echo 'test'`), "test");
  });

  it("should throw error once command fail", async function () {
    try {
      await OfficeAddinGenerator.childProcessExec("exit -1");
    } catch (err) {
      chai.assert(err.message, "Command failed: exit -1");
    }
  });

  it("should call both doScaffolding and template generator", async function () {
    const inputs: Inputs = {
      platform: Platform.CLI,
      projectPath: testFolder,
      ProjectType: ProjectTypeOptions.outlookAddin().id,
      "app-name": "outlook-addin-test",
    };
    const doScaffoldStub = sinon
      .stub(OfficeAddinGenerator, "doScaffolding")
      .resolves(ok(undefined));
    const generateTemplateStub = sinon.stub(Generator, "generateTemplate").resolves(ok(undefined));

    const result = await OfficeAddinGenerator.generate(context, inputs, testFolder);

    chai.expect(result.isOk()).to.eq(true);
    chai.expect(doScaffoldStub.calledOnce).to.be.true;
    chai.expect(generateTemplateStub.calledOnce).to.be.true;
  });

  it("should return error if doScaffolding() returns error", async function () {
    const inputs: Inputs = {
      platform: Platform.CLI,
      projectPath: testFolder,
      ProjectType: ProjectTypeOptions.outlookAddin().id,
      "app-name": "outlook-addin-test",
    };
    sinon.stub(OfficeAddinGenerator, "doScaffolding").resolves(err(mockedError));
    sinon.stub(Generator, "generateTemplate").resolves(ok(undefined));

    const result = await OfficeAddinGenerator.generate(context, inputs, testFolder);

    chai.assert.isTrue(result.isErr() && result.error.name === "mockedError");
  });

  it("should call both doScaffolding and template generator", async function () {
    const inputs: Inputs = {
      platform: Platform.CLI,
      projectPath: testFolder,
      ProjectType: ProjectTypeOptions.outlookAddin().id,
      "app-name": "outlook-addin-test",
    };
    sinon.stub(OfficeAddinGenerator, "doScaffolding").resolves(ok(undefined));
    sinon.stub(Generator, "generateTemplate").resolves(err(mockedError));

    const result = await OfficeAddinGenerator.generate(context, inputs, testFolder);

    chai.assert.isTrue(result.isErr() && result.error.name === "mockedError");
  });

  it("should scaffold taskpane successfully on happy path", async () => {
    const inputs: Inputs = {
      platform: Platform.CLI,
      projectPath: testFolder,
      ProjectType: ProjectTypeOptions.outlookAddin().id,
      "app-name": "outlook-addin-test",
    };
    inputs[QuestionNames.ProjectType] = ProjectTypeOptions.outlookAddin().id;
    inputs[QuestionNames.Capabilities] = "json-taskpane";
    inputs[QuestionNames.OfficeAddinFolder] = undefined;
    inputs[QuestionNames.ProgrammingLanguage] = "typescript";

    sinon.stub(OfficeAddinGenerator, "childProcessExec").resolves();
    sinon.stub(HelperMethods, "downloadProjectTemplateZipFile").resolves(undefined);
    sinon.stub(OfficeAddinManifest, "modifyManifestFile").resolves({});
    const result = await OfficeAddinGenerator.doScaffolding(context, inputs, testFolder);

    chai.expect(result.isOk()).to.eq(true);
  });

  it("should scaffold taskpane failed, throw error", async () => {
    const inputs: Inputs = {
      platform: Platform.CLI,
      projectPath: testFolder,
      "app-name": "outlook-addin-test",
    };
    inputs[QuestionNames.ProjectType] = ProjectTypeOptions.outlookAddin().id;
    inputs[QuestionNames.Capabilities] = "json-taskpane";
    inputs[QuestionNames.OfficeAddinFolder] = undefined;
    inputs[QuestionNames.ProgrammingLanguage] = "typescript";

    sinon.stub(OfficeAddinGenerator, "childProcessExec").resolves();
    sinon.stub(HelperMethods, "downloadProjectTemplateZipFile").rejects(new UserCancelError());
    sinon.stub(OfficeAddinManifest, "modifyManifestFile").resolves({});
    const result = await OfficeAddinGenerator.doScaffolding(context, inputs, testFolder);

    chai.expect(result.isErr()).to.eq(true);
  });

  it("should copy addin files and updateManifest if addin folder is specified with json manifest", async () => {
    const inputs: Inputs = {
      platform: Platform.CLI,
      projectPath: testFolder,
      "app-name": "outlook-addin-test",
    };
    inputs[QuestionNames.ProjectType] = ProjectTypeOptions.outlookAddin().id;
    inputs[QuestionNames.Capabilities] = "json-taskpane";
    inputs[QuestionNames.OfficeAddinFolder] = "somepath";
    inputs[QuestionNames.ProgrammingLanguage] = "typescript";
    inputs[QuestionNames.OfficeAddinManifest] = "manifest.json";

    const copyAddinFilesStub = sinon
      .stub(HelperMethods, "copyAddinFiles")
      .callsFake((from: string, to: string) => {
        return;
      });
    const updateManifestStub = sinon
      .stub(HelperMethods, "updateManifest")
      .callsFake(async (destination: string, manifestPath: string) => {
        return;
      });

    sinon.stub<any, any>(ManifestUtil, "loadFromPath").resolves({
      extensions: [
        {
          requirements: {
            scopes: ["mail"],
          },
        },
      ],
    });

    const result = await OfficeAddinGenerator.doScaffolding(context, inputs, testFolder);

    chai.expect(result.isOk()).to.eq(true);
    chai.expect(copyAddinFilesStub.calledOnce).to.be.true;
    chai.expect(updateManifestStub.calledOnce).to.be.true;
    chai.expect(inputs[QuestionNames.OfficeAddinHost]).to.eq("Outlook");

    const hostResult = await getHost(inputs[QuestionNames.OfficeAddinFolder]);
    chai.expect(hostResult).to.equal("Outlook");
  });

  it("should copy addin files and convert manifest if addin folder is specified with xml manifest", async () => {
    const inputs: Inputs = {
      platform: Platform.CLI,
      projectPath: testFolder,
      "app-name": "outlook-addin-test",
    };
    inputs[QuestionNames.ProjectType] = ProjectTypeOptions.outlookAddin().id;
    inputs[QuestionNames.Capabilities] = "json-taskpane";
    inputs[QuestionNames.OfficeAddinFolder] = "somepath";
    inputs[QuestionNames.ProgrammingLanguage] = "typescript";
    inputs[QuestionNames.OfficeAddinManifest] = "manifest.xml";

    let progressBarStartCalled = 0;
    let progressBarNextCalled = 0;
    let progessBarEndCalled = 0;
    const createProgressBarStub = sinon.stub(context.userInteraction, "createProgressBar").returns({
      start: async () => {
        progressBarStartCalled++;
      },
      next: async () => {
        progressBarNextCalled++;
      },
      end: async () => {
        progessBarEndCalled++;
      },
    });

    const copyAddinFilesStub = sinon
      .stub(HelperMethods, "copyAddinFiles")
      .callsFake((from: string, to: string) => {
        return;
      });
    const updateManifestStub = sinon
      .stub(HelperMethods, "updateManifest")
      .callsFake(async (destination: string, manifestPath: string) => {
        return;
      });
    const convertProjectStub = sinon
      .stub()
      .callsFake(async (manifestPath?: string, backupPath?: string) => {
        return;
      });

    const generator = proxyquire("../../../src/component/generator/officeAddin/generator", {
      "office-addin-project": {
        convertProject: convertProjectStub,
      },
    });

    sinon.stub<any, any>(ManifestUtil, "loadFromPath").resolves({
      extensions: [
        {
          requirements: {
            scopes: ["mail"],
          },
        },
      ],
    });

    const result = await generator.OfficeAddinGenerator.doScaffolding(context, inputs, testFolder);

    chai.expect(result.isOk()).to.eq(true);
    chai.expect(copyAddinFilesStub.calledOnce).to.be.true;
    chai.expect(updateManifestStub.calledOnce).to.be.true;
    chai.expect(convertProjectStub.calledOnce).to.be.true;
    chai.expect(inputs[QuestionNames.OfficeAddinHost]).to.eq("Outlook");
    chai.expect(progressBarStartCalled).to.eq(1);
    chai.expect(progressBarNextCalled).to.eq(3);
    chai.expect(progessBarEndCalled).to.eq(1);

    const hostResult = await getHost(inputs[QuestionNames.OfficeAddinFolder]);
    chai.expect(hostResult).to.equal("Outlook");
  });

  afterEach(async () => {
    sinon.restore();
    mockedEnvRestore();
    if (await fse.pathExists(testFolder)) {
      await fse.rm(testFolder, { recursive: true });
    }
  });

  it(`should generate common template if language is undefined`, async () => {
    const inputs: Inputs = {
      platform: Platform.CLI,
      projectPath: testFolder,
      ProjectType: ProjectTypeOptions.outlookAddin().id,
      "app-name": "outlook-addin-test",
      "programming-language": undefined,
    };
    sinon.stub(OfficeAddinGenerator, "doScaffolding").resolves(ok(undefined));
    const stub = sinon.stub(Generator, "generateTemplate").resolves(ok(undefined));

    const result = await OfficeAddinGenerator.generate(context, inputs, testFolder);
    chai.assert.isTrue(result.isOk());
    // chai.assert.isTrue(stub.calledWith(context, testFolder, "office-addin", undefined));
  });

  it(`should generate ts template if language is "typescript"`, async () => {
    const inputs: Inputs = {
      platform: Platform.CLI,
      projectPath: testFolder,
      ProjectType: ProjectTypeOptions.outlookAddin().id,
      "app-name": "outlook-addin-test",
      "programming-language": "typescript",
    };
    sinon.stub(OfficeAddinGenerator, "doScaffolding").resolves(ok(undefined));
    const stub = sinon.stub(Generator, "generateTemplate").resolves(ok(undefined));

    const result = await OfficeAddinGenerator.generate(context, inputs, testFolder);

    chai.assert.isTrue(result.isOk() && stub.calledWith(context, testFolder, "office-addin", "ts"));
  });

  it(`should generate js template if language is "javascript"`, async () => {
    const inputs: Inputs = {
      platform: Platform.CLI,
      projectPath: testFolder,
      ProjectType: ProjectTypeOptions.outlookAddin().id,
      "app-name": "outlook-addin-test",
      "programming-language": "javascript",
    };
    sinon.stub(OfficeAddinGenerator, "doScaffolding").resolves(ok(undefined));
    const stub = sinon.stub(Generator, "generateTemplate").resolves(ok(undefined));

    const result = await OfficeAddinGenerator.generate(context, inputs, testFolder);

    chai.assert.isTrue(result.isOk() && stub.calledWith(context, testFolder, "office-addin", "js"));
  });
});

describe("helperMethods", async () => {
  describe("updateManifest", () => {
    const sandbox = sinon.createSandbox();
    const manifestPath = "manifestPath";
    const manifestTemplatePath = "manifestTemplatePath";
    let writePathResult: devPreview.DevPreviewSchema | undefined = undefined;

    beforeEach(() => {
      sandbox.stub(ManifestUtil, "loadFromPath").callsFake(async (path) => {
        if (path === manifestPath) {
          return {
            extensions: [],
            authorization: {
              permissions: {
                resourceSpecific: [],
              },
            },
          } as unknown as devPreview.DevPreviewSchema;
        } else if (path === manifestTemplatePath) {
          return {
            extensions: undefined,
            authorization: undefined,
          } as unknown as devPreview.DevPreviewSchema;
        }

        throw new Error("Invalid path");
      });

      sandbox.stub(ManifestUtil, "writeToPath").callsFake(async (path, manifest) => {
        writePathResult = manifest as devPreview.DevPreviewSchema;
        return;
      });

      sandbox.stub(manifestUtils, "getTeamsAppManifestPath").returns(manifestTemplatePath);
    });

    afterEach(() => {
      sandbox.restore();
      writePathResult = undefined;
    });

    it("should update manifest's extenstions and authorization", async () => {
      sandbox.stub(fse, "pathExists").resolves(true);
      await HelperMethods.updateManifest("", manifestPath);

      chai.assert.isDefined(writePathResult);
      chai.assert.equal(writePathResult?.extensions?.length, 0);
      chai.assert.equal(writePathResult?.authorization?.permissions?.resourceSpecific?.length, 0);
    });

    it("should early return if there's no appPackage folder", async () => {
      sandbox.stub(fse, "pathExists").resolves(false);
      await HelperMethods.updateManifest("", manifestPath);

      chai.assert.isUndefined(writePathResult, "writeToPath should not be called");
    });
  });

  describe("downloadProjectTemplateZipFile", async () => {
    const sandbox = sinon.createSandbox();
    class ResponseData extends EventEmitter {
      pipe(ws: fs.WriteStream) {
        return this;
      }
    }

    class MockedWriteStream {
      on(event: string, cb: () => void) {
        return this;
      }
    }

    afterEach(() => {
      sandbox.restore();
    });
    it("should fetch fail", async () => {
      const resp = new ResponseData();
      sandbox.stub(fetch, "default").rejects(new Error());
      const mockedStream = new MockedWriteStream();
      const unzipStub = sandbox.stub(HelperMethods, "unzipProjectTemplate").resolves();
      sandbox.stub<any, any>(fs, "createWriteStream").returns(mockedStream);
      try {
        await HelperMethods.downloadProjectTemplateZipFile("", "");
        chai.assert.fail("should not reach here");
      } catch (e) {
        chai.assert.isTrue(e instanceof AccessGithubError);
      }
    });
    it("should download project template zip file", async () => {
      const resp = new ResponseData();
      sandbox.stub(fetch, "default").resolves({ body: resp } as any);
      const mockedStream = new MockedWriteStream();
      const unzipStub = sandbox.stub(HelperMethods, "unzipProjectTemplate").resolves();
      sandbox.stub<any, any>(fs, "createWriteStream").returns(mockedStream);
      const promise = HelperMethods.downloadProjectTemplateZipFile("", "");
      // manully wait for the close event to be registered
      await new Promise((resolve) => setTimeout(resolve, 100));
      resp.emit("close");
      await promise;
      chai.assert.isTrue(unzipStub.calledOnce);
    });

    it("unzipProjectTemplate error", async () => {
      const resp = new ResponseData();
      sandbox.stub(fetch, "default").resolves({ body: resp } as any);
      const mockedStream = new MockedWriteStream();
      sandbox.stub(HelperMethods, "unzipProjectTemplate").rejects(new Error());
      sandbox.stub<any, any>(fs, "createWriteStream").returns(mockedStream);
      const promise = HelperMethods.downloadProjectTemplateZipFile("", "");
      // manully wait for the close event to be registered
      await new Promise((resolve) => setTimeout(resolve, 100));
      resp.emit("close");
      try {
        await promise;
        chai.assert.fail("should throw error");
      } catch (e) {}
    });

    it("download error", async () => {
      const resp = new ResponseData();
      sandbox.stub(fetch, "default").resolves({ body: resp } as any);
      const mockedStream = new MockedWriteStream();
      const unzipStub = sandbox.stub(HelperMethods, "unzipProjectTemplate").resolves();
      sandbox.stub<any, any>(fs, "createWriteStream").returns(mockedStream);
      const promise = HelperMethods.downloadProjectTemplateZipFile("", "");
      // manully wait for the close event to be registered
      await new Promise((resolve) => setTimeout(resolve, 100));
      resp.emit("error", new Error());
      try {
        await promise;
        chai.assert.fail("should throw error");
      } catch (e) {}
      chai.assert.isTrue(unzipStub.notCalled);
    });

    it("Response body is null.", async () => {
      sandbox.stub(fetch, "default").resolves({ body: null } as any);
      const promise = HelperMethods.downloadProjectTemplateZipFile("", "");
      try {
        await promise;
        chai.assert.fail("should throw error");
      } catch (e) {
        chai.assert.isTrue(e instanceof AccessGithubError);
      }
    });
  });

  describe("unzipProjectTemplate", () => {
    const sandbox = sinon.createSandbox();

    class MockedReadStream {
      on(event: string, cb: () => void) {
        return this;
      }

      pipe(ws: fs.WriteStream) {
        return this;
      }
    }

    afterEach(() => {
      sandbox.restore();
    });

    it("work as expected", async () => {
      sandbox.stub<any, any>(fs, "createReadStream").returns(new MockedReadStream());
      sandbox.stub<any, any>(unzip, "Extract").returns({});
      try {
        HelperMethods.unzipProjectTemplate("");
      } catch (err) {
        chai.assert.fail(err);
      } finally {
        sandbox.restore();
      }
    });

    it("unzipErrorHandler", async () => {
      let i = 0;
      const reject = () => {
        i++;
      };
      unzipErrorHandler("", reject, new Error());
      chai.assert.equal(i, 1);
    });
    it("unzipErrorHandler 2", async () => {
      let i = 0;
      const reject = () => {
        i++;
      };
      unzipErrorHandler("", reject, new Error("test"));
      chai.assert.equal(i, 1);
    });
  });

  describe("moveUnzippedFiles", () => {
    const projectRoot = "/home/user/teamsapp";

    beforeEach(() => {
      mockfs({
        "/home/user/teamsapp/project.zip": "xxx",
        "/home/user/teamsapp/project": {
          file1: "xxx",
          file2: "yyy",
        },
      });
    });

    afterEach(() => {
      mockfs.restore();
    });

    it("should remove zip file and unzipped folder and copy files", async () => {
      try {
        HelperMethods.moveUnzippedFiles(projectRoot);
        chai.assert.equal(fs.existsSync("/home/user/teamsapp/project.zip"), false);
        chai.assert.equal(fs.existsSync("/home/user/teamsapp/project"), false);
        chai.assert.equal(fs.existsSync("/home/user/teamsapp/file1"), true);
        chai.assert.equal(fs.existsSync("/home/user/teamsapp/file2"), true);
      } catch (err) {
        chai.assert.fail(err);
      }
    });
  });

  describe("copyAddinFiles", () => {
    const projectRoot = "/home/user/teamsapp";

    beforeEach(() => {
      mockfs({
        "/home/user/teamsapp/.gitignore": "xxx",
        "/home/user/teamsapp/project": {
          file1: "xxx",
          file2: "yyy",
        },
        "/home/user/teamsapp/node_modules": {
          file3: "xxx",
        },
      });
    });

    afterEach(() => {
      mockfs.restore();
    });

    it("should copy project files and .gitignore but ignore node_modules", async () => {
      try {
        const destination = "/home/user/destination";
        HelperMethods.copyAddinFiles(projectRoot, destination);
        chai.assert.equal(fs.existsSync(path.join(destination, "project", "file1")), true);
        chai.assert.equal(fs.existsSync(path.join(destination, "project", "file2")), true);
        chai.assert.equal(fs.existsSync(path.join(destination, ".gitignore")), true);
        chai.assert.equal(fs.existsSync(path.join(destination, "node_modules")), false);
      } catch (err) {
        chai.assert.fail(err);
      }
    });
  });

  describe("moveManifestLocation", () => {
    const projectRoot = "/home/user/addin";

    beforeEach(() => {
      mockfs({
        "/home/user/addin/manifest.json": "{}",
        "/home/user/addin/assets": {
          file1: "xxx",
        },
        "/home/user/addin/webpack.config.js": JSON.stringify([
          {
            from: "assets/*",
            to: "assets/[name][ext][query]",
          },
          {
            from: "manifest*.json",
            to: "[name]" + "[ext]",
          },
        ]),
        "/home/user/addin/package.json": JSON.stringify({
          scripts: {
            start: "office-addin-debugging start manifest.json",
            stop: "office-addin-debugging stop manifest.json",
            validate: "office-addin-manifest validate manifest.json",
          },
        }),
        "/home/user/addin/src/taskpane/taskpane.html": `<img width="90" height="90" src="../../assets/logo-filled.png" alt="Contoso" title="Contoso" />`,
      });
    });

    afterEach(() => {
      mockfs.restore();
    });

    it("should move manifest.json into appPackage folder", async () => {
      await HelperMethods.moveManifestLocation(projectRoot, "manifest.json");
      chai.assert.isFalse(await fse.pathExists(path.join(projectRoot, "manifest.json")));
      chai.assert.isFalse(await fse.pathExists(path.join(projectRoot, "assets")));

      chai.assert.isTrue(
        await fse.pathExists(path.join(projectRoot, "appPackage", "manifest.json"))
      );
      chai.assert.isTrue(
        await fse.pathExists(path.join(projectRoot, "appPackage", "assets", "file1"))
      );

      const webpackConfigPath = path.join(projectRoot, "webpack.config.js");
      const webpackConfigJson = JSON.parse(await fse.readFile(webpackConfigPath, "utf8"));
      chai.assert.equal(webpackConfigJson[0].from, "appPackage/assets/*");
      chai.assert.equal(webpackConfigJson[1].from, "appPackage/manifest*.json");

      const packageJsonPath = path.join(projectRoot, "package.json");
      const packageJson = JSON.parse(await fse.readFile(packageJsonPath, "utf8"));
      chai.assert.equal(
        packageJson.scripts.start,
        "office-addin-debugging start appPackage/manifest.json"
      );

      chai.assert.equal(
        packageJson.scripts.stop,
        "office-addin-debugging stop appPackage/manifest.json"
      );
      chai.assert.equal(
        packageJson.scripts.validate,
        "office-addin-manifest validate appPackage/manifest.json"
      );

      const htmlPath = path.join(projectRoot, "src", "taskpane", "taskpane.html");
      const html = await fse.readFile(htmlPath, "utf8");
      chai.assert.equal(
        html,
        `<img width="90" height="90" src="../../appPackage/assets/logo-filled.png" alt="Contoso" title="Contoso" />`
      );
    });
  });
});

describe("OfficeAddinGenerator for Office Addin", function () {
  const testFolder = path.resolve("./tmp");
  let context: Context;
  let mockedEnvRestore: RestoreFn;
  const mockedError = new SystemError("mockedSource", "mockedError", "mockedMessage");

  beforeEach(async () => {
    mockedEnvRestore = mockedEnv({ TEAMSFX_V3: "true" }, { clear: true });
    const gtools = new MockTools();
    setTools(gtools);
    context = createContextV3();

    await fse.ensureDir(testFolder);
    sinon.stub(fs, "stat").resolves();
    sinon.stub(cpUtils, "executeCommand").resolves("succeed");
    const manifestId = uuid.v4();
    sinon.stub(fs, "readFile").resolves(new Buffer(`{"id": "${manifestId}"}`));
    sinon.stub(fs, "writeFile").resolves();
    sinon.stub(fs, "rename").resolves();
    sinon.stub(fs, "copyFile").resolves();
    sinon.stub(fse, "remove").resolves();
    sinon.stub(fse, "readJson").resolves({});
    sinon.stub(fse, "ensureFile").resolves();
    sinon.stub(fse, "writeJSON").resolves();
  });

  it("should run childProcessExec command success", async function () {
    sinon.stub(childProcess, "exec").yields(`echo 'test'`, "test");
    chai.assert(await OfficeAddinGenerator.childProcessExec(`echo 'test'`), "test");
  });

  it("should throw error once command fail", async function () {
    try {
      await OfficeAddinGenerator.childProcessExec("exit -1");
    } catch (err) {
      chai.assert(err.message, "Command failed: exit -1");
    }
  });

  it("should call both doScaffolding and template generator", async function () {
    const inputs: Inputs = {
      platform: Platform.CLI,
      projectPath: testFolder,
      "project-type": ProjectTypeOptions.officeAddin().id,
      "app-name": "office-addin-test",
      "office-addin-framework-type": "default",
    };
    const doScaffoldStub = sinon
      .stub(OfficeAddinGenerator, "doScaffolding")
      .resolves(ok(undefined));
    const generateTemplateStub = sinon.stub(Generator, "generateTemplate").resolves(ok(undefined));

    const result = await OfficeAddinGenerator.generate(context, inputs, testFolder);

    chai.expect(result.isOk()).to.eq(true);
    chai.expect(doScaffoldStub.calledOnce).to.be.true;
    chai.expect(generateTemplateStub.calledOnce).to.be.true;
  });

  it("should return error if doScaffolding() returns error", async function () {
    const inputs: Inputs = {
      platform: Platform.CLI,
      projectPath: testFolder,
      "project-type": ProjectTypeOptions.officeAddin().id,
      "app-name": "office-addin-test",
      "office-addin-framework-type": "default",
    };
    sinon.stub(OfficeAddinGenerator, "doScaffolding").resolves(err(mockedError));
    sinon.stub(Generator, "generateTemplate").resolves(ok(undefined));

    const result = await OfficeAddinGenerator.generate(context, inputs, testFolder);

    chai.assert.isTrue(result.isErr() && result.error.name === "mockedError");
  });

  it("should call both doScaffolding and template generator", async function () {
    const inputs: Inputs = {
      platform: Platform.CLI,
      projectPath: testFolder,
      "project-type": ProjectTypeOptions.officeAddin().id,
      "app-name": "office-addin-test",
      "office-addin-framework-type": "default",
    };
    sinon.stub(OfficeAddinGenerator, "doScaffolding").resolves(ok(undefined));
    sinon.stub(Generator, "generateTemplate").resolves(err(mockedError));

    const result = await OfficeAddinGenerator.generate(context, inputs, testFolder);

    chai.assert.isTrue(result.isErr() && result.error.name === "mockedError");
  });

  it("should scaffold taskpane successfully on happy path", async () => {
    const inputs: Inputs = {
      platform: Platform.CLI,
      projectPath: testFolder,
      "project-type": ProjectTypeOptions.officeAddin().id,
      "app-name": "office-addin-test",
      "office-addin-framework-type": "default",
    };
    inputs[QuestionNames.Capabilities] = "json-taskpane";
    inputs[QuestionNames.OfficeAddinFolder] = undefined;
    inputs[QuestionNames.ProgrammingLanguage] = "typescript";

    sinon.stub(OfficeAddinGenerator, "childProcessExec").resolves();
    sinon.stub(HelperMethods, "downloadProjectTemplateZipFile").resolves(undefined);
    sinon.stub(OfficeAddinManifest, "modifyManifestFile").resolves({});
    const result = await OfficeAddinGenerator.doScaffolding(context, inputs, testFolder);

    chai.expect(result.isOk()).to.eq(true);
  });

  it("should scaffold taskpane failed, throw error", async () => {
    const inputs: Inputs = {
      platform: Platform.CLI,
      projectPath: testFolder,
      "app-name": "office-addin-test",
    };
    inputs[QuestionNames.ProjectType] = ProjectTypeOptions.officeAddin().id;
    inputs[QuestionNames.Capabilities] = "json-taskpane";
    inputs[QuestionNames.OfficeAddinFolder] = undefined;
    inputs[QuestionNames.ProgrammingLanguage] = "typescript";
    inputs[QuestionNames.OfficeAddinFramework] = "default";

    sinon.stub(OfficeAddinGenerator, "childProcessExec").resolves();
    sinon.stub(HelperMethods, "downloadProjectTemplateZipFile").rejects(new UserCancelError());
    sinon.stub(OfficeAddinManifest, "modifyManifestFile").resolves({});
    const result = await OfficeAddinGenerator.doScaffolding(context, inputs, testFolder);

    chai.expect(result.isErr()).to.eq(true);
  });

  const testCases = [
    { scope: "document", host: "Word" },
    { scope: "workbook", host: "Excel" },
    { scope: "presentation", host: "PowerPoint" },
  ];

  testCases.forEach((testCase) => {
    it(`should copy addin files and updateManifest if addin folder is specified with json manifest for ${testCase.host}`, async () => {
      const inputs: Inputs = {
        platform: Platform.CLI,
        projectPath: testFolder,
        "app-name": "office-addin-test",
      };
      inputs[QuestionNames.ProjectType] = ProjectTypeOptions.officeAddin().id;
      inputs[QuestionNames.Capabilities] = "json-taskpane";
      inputs[QuestionNames.OfficeAddinFolder] = "somepath";
      inputs[QuestionNames.ProgrammingLanguage] = "typescript";
      inputs[QuestionNames.OfficeAddinFramework] = "default";
      inputs[QuestionNames.OfficeAddinManifest] = "manifest.json";

      const copyAddinFilesStub = sinon
        .stub(HelperMethods, "copyAddinFiles")
        .callsFake((from: string, to: string) => {
          return;
        });
      const updateManifestStub = sinon
        .stub(HelperMethods, "updateManifest")
        .callsFake(async (destination: string, manifestPath: string) => {
          return;
        });

      sinon.stub<any, any>(ManifestUtil, "loadFromPath").resolves({
        extensions: [
          {
            requirements: {
              scopes: [testCase.scope],
            },
          },
        ],
      });

      const result = await OfficeAddinGenerator.doScaffolding(context, inputs, testFolder);

      chai.expect(result.isOk()).to.eq(true);
      chai.expect(copyAddinFilesStub.calledOnce).to.be.true;
      chai.expect(updateManifestStub.calledOnce).to.be.true;
      chai.expect(inputs[QuestionNames.OfficeAddinHost]).to.equal(testCase.host);
      const hostResult = await getHost(inputs[QuestionNames.OfficeAddinFolder]);
      chai.expect(hostResult).to.equal(testCase.host);
    });
  });

  testCases.forEach((testCase) => {
    it(`should copy addin files and convert manifest if addin folder is specified with xml manifest for ${testCase.host}`, async () => {
      const inputs: Inputs = {
        platform: Platform.CLI,
        projectPath: testFolder,
        "app-name": "office-addin-test",
        [QuestionNames.ProjectType]: ProjectTypeOptions.officeAddin().id,
        [QuestionNames.Capabilities]: "json-taskpane",
        [QuestionNames.OfficeAddinFolder]: "somepath",
        [QuestionNames.ProgrammingLanguage]: "typescript",
        [QuestionNames.OfficeAddinFramework]: "default",
        [QuestionNames.OfficeAddinManifest]: "manifest.xml",
      };

      let progressBarStartCalled = 0;
      let progressBarNextCalled = 0;
      let progessBarEndCalled = 0;
      const createProgressBarStub = sinon
        .stub(context.userInteraction, "createProgressBar")
        .returns({
          start: async () => {
            progressBarStartCalled++;
          },
          next: async () => {
            progressBarNextCalled++;
          },
          end: async () => {
            progessBarEndCalled++;
          },
        });

      const copyAddinFilesStub = sinon
        .stub(HelperMethods, "copyAddinFiles")
        .callsFake((from: string, to: string) => {
          return;
        });
      const updateManifestStub = sinon
        .stub(HelperMethods, "updateManifest")
        .callsFake(async (destination: string, manifestPath: string) => {
          return;
        });
      const convertProjectStub = sinon
        .stub()
        .callsFake(async (manifestPath?: string, backupPath?: string) => {
          return;
        });

      const generator = proxyquire("../../../src/component/generator/officeAddin/generator", {
        "office-addin-project": {
          convertProject: convertProjectStub,
        },
      });

      sinon.stub<any, any>(ManifestUtil, "loadFromPath").resolves({
        extensions: [
          {
            requirements: {
              scopes: [testCase.scope],
            },
          },
        ],
      });

      const result = await generator.OfficeAddinGenerator.doScaffolding(
        context,
        inputs,
        testFolder
      );

      chai.expect(result.isOk()).to.eq(true);
      chai.expect(copyAddinFilesStub.calledOnce).to.be.true;
      chai.expect(updateManifestStub.calledOnce).to.be.true;
      chai.expect(convertProjectStub.calledOnce).to.be.true;
      chai.expect(inputs[QuestionNames.OfficeAddinHost]).to.equal(testCase.host);
      chai.expect(progressBarStartCalled).to.eq(1);
      chai.expect(progressBarNextCalled).to.eq(3);
      chai.expect(progessBarEndCalled).to.eq(1);

      const resultHost = await getHost(inputs[QuestionNames.OfficeAddinFolder]);
      chai.expect(resultHost).to.equal(testCase.host);
    });
  });

  afterEach(async () => {
    sinon.restore();
    mockedEnvRestore();
    if (await fse.pathExists(testFolder)) {
      await fse.rm(testFolder, { recursive: true });
    }
  });

  it(`should generate common template if language is undefined`, async () => {
    const inputs: Inputs = {
      platform: Platform.CLI,
      projectPath: testFolder,
      "app-name": "office-addin-test",
    };
    inputs[QuestionNames.ProjectType] = ProjectTypeOptions.officeAddin().id;
    inputs[QuestionNames.Capabilities] = "json-taskpane";
    inputs[QuestionNames.ProgrammingLanguage] = undefined;
    inputs[QuestionNames.OfficeAddinFramework] = "default";

    sinon.stub(OfficeAddinGenerator, "doScaffolding").resolves(ok(undefined));
    const stub = sinon.stub(Generator, "generateTemplate").resolves(ok(undefined));

    const result = await OfficeAddinGenerator.generate(context, inputs, testFolder);
    chai.assert.isTrue(result.isOk());
    // chai.assert.isTrue(stub.calledWith(context, testFolder, "office-json-addin", undefined));
  });

  it(`should generate taskpane ts template if language is "typescript"`, async () => {
    const inputs: Inputs = {
      platform: Platform.CLI,
      projectPath: testFolder,
      "app-name": "office-addin-test",
    };
    inputs[QuestionNames.ProjectType] = ProjectTypeOptions.officeAddin().id;
    inputs[QuestionNames.Capabilities] = "json-taskpane";
    inputs[QuestionNames.ProgrammingLanguage] = "typescript";
    inputs[QuestionNames.OfficeAddinFramework] = "default";

    sinon.stub(OfficeAddinGenerator, "doScaffolding").resolves(ok(undefined));
    const stub = sinon.stub(Generator, "generateTemplate").resolves(ok(undefined));

    const result = await OfficeAddinGenerator.generate(context, inputs, testFolder);

    chai.assert.isTrue(result.isOk());
    chai.assert.isTrue(stub.calledWith(context, testFolder, "office-json-addin", "ts"));
  });

  it(`should generate taskpane js template if language is "JavaScript"`, async () => {
    const inputs: Inputs = {
      platform: Platform.CLI,
      projectPath: testFolder,
      "app-name": "office-addin-test",
    };
    inputs[QuestionNames.ProjectType] = ProjectTypeOptions.officeAddin().id;
    inputs[QuestionNames.Capabilities] = "json-taskpane";
    inputs[QuestionNames.ProgrammingLanguage] = "JavaScript";
    inputs[QuestionNames.OfficeAddinFramework] = "default";

    sinon.stub(OfficeAddinGenerator, "doScaffolding").resolves(ok(undefined));
    const stub = sinon.stub(Generator, "generateTemplate").resolves(ok(undefined));

    const result = await OfficeAddinGenerator.generate(context, inputs, testFolder);

    chai.assert.isTrue(
      result.isOk() && stub.calledWith(context, testFolder, "office-json-addin", "js")
    );
  });

  it(`should generate content ts template if language is "typescript"`, async () => {
    const inputs: Inputs = {
      platform: Platform.CLI,
      projectPath: testFolder,
      "app-name": "office-addin-test",
    };
    inputs[QuestionNames.ProjectType] = ProjectTypeOptions.officeAddin().id;
    inputs[QuestionNames.Capabilities] = "json-taskpane";
    inputs[QuestionNames.ProgrammingLanguage] = "typescript";
    inputs[QuestionNames.OfficeAddinFramework] = "default";

    sinon.stub(OfficeAddinGenerator, "doScaffolding").resolves(ok(undefined));
    const stub = sinon.stub(Generator, "generateTemplate").resolves(ok(undefined));

    const result = await OfficeAddinGenerator.generate(context, inputs, testFolder);

    chai.assert.isTrue(result.isOk());
    chai.assert.isTrue(stub.calledWith(context, testFolder, "office-json-addin", "ts"));
  });

  it(`should generate content js template if language is "JavaScript"`, async () => {
    const inputs: Inputs = {
      platform: Platform.CLI,
      projectPath: testFolder,
      "app-name": "office-addin-test",
    };
    inputs[QuestionNames.ProjectType] = ProjectTypeOptions.officeAddin().id;
    inputs[QuestionNames.Capabilities] = "json-taskpane";
    inputs[QuestionNames.ProgrammingLanguage] = "JavaScript";
    inputs[QuestionNames.OfficeAddinFramework] = "default";

    sinon.stub(OfficeAddinGenerator, "doScaffolding").resolves(ok(undefined));
    const stub = sinon.stub(Generator, "generateTemplate").resolves(ok(undefined));

    const result = await OfficeAddinGenerator.generate(context, inputs, testFolder);

    chai.assert.isTrue(
      result.isOk() && stub.calledWith(context, testFolder, "office-json-addin", "js")
    );
  });
});
