// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import "mocha";
import { Duplex } from "stream";
import * as chai from "chai";
import chaiAsPromised from "chai-as-promised";
import sinon from "sinon";
import ServerAzureAccountProvider from "../../../src/providers/token/azure";
import { createMessageConnection } from "vscode-jsonrpc";
import { err, ok } from "@microsoft/teamsfx-api";
import { expect } from "chai";
import { TokenCredentialsBase } from "@azure/ms-rest-nodeauth";

chai.use(chaiAsPromised);

class TestStream extends Duplex {
  _write(chunk: string, _encoding: string, done: () => void) {
    this.emit("data", chunk);
    done();
  }

  _read(_size: number) {}
}

describe("azure", () => {
  const sandbox = sinon.createSandbox();
  const up = new TestStream();
  const down = new TestStream();
  const msgConn = createMessageConnection(up as any, down as any);

  after(() => {
    sandbox.restore();
  });

  it("constructor", () => {
    const azure = new ServerAzureAccountProvider(msgConn);
    chai.assert.equal(azure["connection"], msgConn);
  });

  describe("getAccountCredentialAsync", () => {
    const azure = new ServerAzureAccountProvider(msgConn);

    it("getAccountCredentialAsync: error", async () => {
      const e = new Error("test");
      const promise = Promise.resolve(err(e));
      const stub = sandbox.stub(msgConn, "sendRequest").returns(promise);
      await chai.expect(azure.getAccountCredentialAsync()).to.be.rejected;
      stub.restore();
    });

    it("getAccountCredentialAsync: ok", () => {
      const token =
        "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJzdWIiOiIxMjM0NTY3ODkwIiwibmFtZSI6IkpvaG4gRG9lIiwiaWF0IjoxNTE2MjM5MDIyfQ.SflKxwRJSMeKKF2QT4fwpMeJf36POk6yJV_adQssw5c";
      const jsonstr = JSON.stringify(
        (function ConvertTokenToJson(token: string) {
          const array = token!.split(".");
          const buff = Buffer.from(array[1], "base64");
          return JSON.parse(buff.toString("utf8"));
        })(token)
      );
      const r = {
        accessToken: token,
        tokenJsonString: jsonstr,
      };
      const promise = Promise.resolve(ok(r));
      const stub = sandbox.stub(msgConn, "sendRequest").returns(promise);
      const res = azure.getAccountCredentialAsync();
      res.then((data) => {
        expect(data instanceof TokenCredentialsBase).is.true;
      });
      stub.restore();
    });
  });

  it("getIdentityCredentialAsync", () => {
    const azure = new ServerAzureAccountProvider(msgConn);
    const res = azure.getIdentityCredentialAsync();
    res.then((data) => {
      chai.assert.isUndefined(data);
    });
  });

  it("signout", async () => {
    const azure = new ServerAzureAccountProvider(msgConn);
    await chai.expect(azure.signout()).to.be.rejected;
  });

  it("setStatusChangeMap", async () => {
    const azure = new ServerAzureAccountProvider(msgConn);
    await chai.expect(azure.setStatusChangeMap("test", sandbox.fake())).to.be.rejected;
  });

  it("removeStatusChangeMap", async () => {
    const azure = new ServerAzureAccountProvider(msgConn);
    await chai.expect(azure.removeStatusChangeMap("test")).to.be.rejected;
  });

  describe("getJsonObject", () => {
    const azure = new ServerAzureAccountProvider(msgConn);

    it("getJsonObject: err", async () => {
      const promise = Promise.resolve(err(new Error("test")));
      const stub = sandbox.stub(msgConn, "sendRequest").returns(promise);
      await chai.expect(azure.getJsonObject()).to.be.rejected;
      stub.restore();
    });

    it("getJsonObject: ok", () => {
      const promise = Promise.resolve(ok("test"));
      const stub = sandbox.stub(msgConn, "sendRequest").returns(promise);
      const res = azure.getJsonObject();
      res.then((data) => {
        chai.expect(data).equal("test");
      });
      stub.restore();
    });
  });

  describe("listSubscriptions", () => {
    const azure = new ServerAzureAccountProvider(msgConn);

    it("listSubscriptions: err", async () => {
      const promise = Promise.resolve(err(new Error("test")));
      const stub = sandbox.stub(msgConn, "sendRequest").returns(promise);
      await chai.expect(azure.listSubscriptions()).to.be.rejected;
      stub.restore();
    });

    it("listSubscriptions: ok", () => {
      const promise = Promise.resolve(ok("test"));
      const stub = sandbox.stub(msgConn, "sendRequest").returns(promise);
      const res = azure.listSubscriptions();
      res.then((data) => {
        chai.expect(data).equal("test");
      });
      stub.restore();
    });
  });

  describe("setSubscription", () => {
    const azure = new ServerAzureAccountProvider(msgConn);

    it("setSubscription: err", async () => {
      const promise = Promise.resolve(err(new Error("test")));
      const stub = sandbox.stub(msgConn, "sendRequest").returns(promise);
      await chai.expect(azure.setSubscription("test")).to.be.rejected;
      stub.restore();
    });

    it("setSubscription: ok", () => {
      const promise = Promise.resolve(ok("test"));
      const stub = sandbox.stub(msgConn, "sendRequest").returns(promise);
      const res = azure.setSubscription("test");
      res.then((data) => {
        chai.expect(data).equal("test");
      });
      stub.restore();
    });
  });

  it("getAccountInfo", () => {
    const azure = new ServerAzureAccountProvider(msgConn);
    chai.expect(() => azure.getAccountInfo()).to.throw();
  });

  describe("getSelectedSubscription", () => {
    const azure = new ServerAzureAccountProvider(msgConn);

    it("getSelectedSubscription: err", async () => {
      const promise = Promise.resolve(err(new Error("test")));
      const stub = sandbox.stub(msgConn, "sendRequest").returns(promise);
      await chai.expect(azure.getSelectedSubscription()).to.be.rejected;
      stub.restore();
    });

    it("getSelectedSubscription: ok", () => {
      const promise = Promise.resolve(ok("test"));
      const stub = sandbox.stub(msgConn, "sendRequest").returns(promise);
      const res = azure.getSelectedSubscription();
      res.then((data) => {
        chai.expect(data).equal("test");
      });
      stub.restore();
    });
  });
});
