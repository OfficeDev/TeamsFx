/* eslint-disable @typescript-eslint/no-non-null-assertion */
import "mocha";
import * as sinon from "sinon";
import * as chai from "chai";
import fs from "fs-extra";
import * as path from "path";
import {
  getAddSPFxQuestionNode,
  getSPFxScaffoldQuestion,
  SPFxTab,
} from "../../../src/component/feature/spfx";
import {
  AutoGeneratedReadme,
  InputsWithProjectPath,
  ok,
  Platform,
  ProjectSettingsV3,
  QTreeNode,
} from "@microsoft/teamsfx-api";
import { SPFxTabCodeProvider } from "../../../src/component/code/spfxTabCode";
import mockedEnv, { RestoreFn } from "mocked-env";
import { MockTools } from "../../core/utils";
import { setTools } from "../../../src/core/globalVars";
import { createContextV3 } from "../../../src/component/utils";
import { getTemplatesFolder } from "../../../src/folder";
import { SPFXQuestionNames } from "../../../src/component/resource/spfx/utils/questions";

describe("spfx", () => {
  describe("add", () => {
    afterEach(() => {
      sinon.restore();
    });

    it("Root README file will be generated", async () => {
      const sourcePath = path.join(
        getTemplatesFolder(),
        "plugins",
        "resource",
        "SPFx",
        "solution",
        "rootREADME.md"
      );
      const targetPath = path.join("c:\\test", AutoGeneratedReadme);
      sinon.stub(SPFxTabCodeProvider.prototype, "generate").resolves(ok(undefined));
      sinon.stub(fs, "ensureDir").resolves();
      sinon.stub(fs, "writeJSON").resolves();
      sinon.stub(fs, "pathExists").callsFake(async (directory) => {
        if (directory === sourcePath) {
          return true;
        }
        if (directory === targetPath) {
          return false;
        }
        return false;
      });
      const stubCopy = sinon.stub(fs, "copy");
      const mockedEnvRestore = mockedEnv({ TEAMSFX_SPFX_MULTI_TAB: "true" });
      const tools = new MockTools();
      setTools(tools);
      const context = createContextV3();
      const projectSetting: ProjectSettingsV3 = {
        appName: "",
        projectId: "",
        programmingLanguage: "typescript",
        components: [
          {
            name: "teams-tab",
            hosting: "spfx",
            deploy: true,
            provision: true,
            build: true,
            folder: "SPFx",
          },
        ],
      };
      context.projectSetting = projectSetting;
      const inputs: InputsWithProjectPath = {
        projectPath: "c:\\test",
        platform: Platform.VSCode,
        language: "typescript",
        "app-name": "spfxtabapp",
      };

      const spfx = new SPFxTab();
      await spfx.add(context, inputs);

      chai.expect(stubCopy.calledOnce);
      chai.expect(stubCopy.calledWith(sourcePath, targetPath));
      mockedEnvRestore();
    });
  });

  describe("getAddSPFxQuestionNode", () => {
    afterEach(() => {
      sinon.restore();
    });

    it("Ask framework when .yo-rc.json not exist", async () => {
      sinon.stub(fs, "pathExists").resolves(false);

      const res = await getAddSPFxQuestionNode("c:\\testFolder");

      chai.expect(res.isOk()).equals(true);
      if (res.isOk()) {
        chai.expect(res.value!.children![0].children!.length).equals(2);
      }
    });

    it("Ask framework when template not persisted in .yo-rc.json", async () => {
      sinon.stub(fs, "pathExists").resolves(true);
      sinon.stub(fs, "readJson").resolves({
        "@microsoft/generator-sharepoint": {
          componentType: "webpart",
        },
      });

      const res = await getAddSPFxQuestionNode("c:\\testFolder");

      chai.expect(res.isOk()).equals(true);
      if (res.isOk()) {
        chai.expect(res.value!.children![0].children!.length).equals(2);
      }
    });

    it("Don't ask framework when template persisted in .yo-rc.json", async () => {
      sinon.stub(fs, "pathExists").resolves(true);
      sinon.stub(fs, "readJson").resolves({
        "@microsoft/generator-sharepoint": {
          componentType: "webpart",
          template: "none",
        },
      });

      const res = await getAddSPFxQuestionNode("c:\\testFolder");

      chai.expect(res.isOk()).equals(true);
      if (res.isOk()) {
        chai.expect(res.value!.children![0].children!.length).equals(1);
      }
    });
  });

  describe("getSPFxScaffoldQuestion: isSpfxDecoupleEnabled", () => {
    const sandbox = sinon.createSandbox();
    let mockedEnvRestore: RestoreFn | undefined;

    afterEach(() => {
      sandbox.restore();
      if (mockedEnvRestore) {
        mockedEnvRestore();
      }
    });

    it("questions: SPFx decouple enabled", () => {
      mockedEnvRestore = mockedEnv({
        TEAMSFX_SPFX_DECOUPLE: "true",
      });

      const node: QTreeNode = getSPFxScaffoldQuestion();

      chai.expect(node.children![0].data.name).equal(SPFXQuestionNames.load_package_version);
      chai
        .expect(node.children![0].children![0].data.name)
        .equal(SPFXQuestionNames.use_global_package_or_install_local);
    });
  });
});
