import { TeamsUserCredential } from "@microsoft/teamsfx";

let instance;

class TeamsUserCredentialContext {
  credential;
  constructor() {
    if (instance) {
      throw new Error("FxContext is a singleton class, use getInstance() instead.");
    }
    instance = this;
  }

  setCredential(credential) {
    this.credential = credential;
  }

  getCredential() {
    if (!this.credential) {
      this.credential =  new TeamsUserCredential({
        initiateLoginEndpoint: process.env.REACT_APP_START_LOGIN_PAGE_URL,
        clientId: process.env.REACT_APP_CLIENT_ID,
      });
    }
    return this.credential;
  }
}

let TeamsUserCredentialContextInstance = Object.freeze(new TeamsUserCredentialContext());

export default TeamsUserCredentialContextInstance;