// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { AzureFunctionHosting } from "./azureFunctionHosting";
import { AzureHosting } from "./azureHosting";
import { BotServiceHosting } from "./botServiceHosting";
import { ServiceType } from "./interfaces";

const HostingMap: { [key: string]: () => AzureHosting } = {
  [ServiceType.Functions]: () => new AzureFunctionHosting(),
  [ServiceType.BotService]: () => new BotServiceHosting(),
};

export class AzureHostingFactory {
  static createHosting(serviceType: ServiceType): AzureHosting {
    if (HostingMap[serviceType] !== undefined) {
      return HostingMap[serviceType]();
    }

    throw new Error(`Host type '${serviceType}' is not supported.`);
  }
}
