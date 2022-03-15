// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
"use strict";

import { returnUserError, UserError } from "@microsoft/teamsfx-api";

export function MissingStep(operation: string, requiredStep: string): UserError {
  return returnUserError(
    new Error(
      `Step "${requiredStep}" is required before ${operation}. Run the required step first.`
    ),
    "localdebug-plugin",
    "MissingStep"
  );
}
