// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { UserError } from "@microsoft/teamsfx-api";
import { getDefaultString, getLocalizedString } from "../../../../common/localizeUtils";
import { maxSecretPerApiKey } from "../utility/constants";

const errorCode = "ApiKeyClientSecretInvalid";
const messageKey = "driver.apiKey.error.clientSecretInvalid";

export class ApiKeyClientSecretInvalidError extends UserError {
  constructor(actionName: string) {
    super({
      source: actionName,
      name: errorCode,
      message: getDefaultString(messageKey, maxSecretPerApiKey),
      displayMessage: getLocalizedString(messageKey, maxSecretPerApiKey),
    });
  }
}
