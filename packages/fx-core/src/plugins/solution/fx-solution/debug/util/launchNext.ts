// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
"use strict";

import * as os from "os";
import { isOfficeAddinEnabled } from "../../../../../common";
import { HubName, LaunchBrowser, LaunchUrl } from "../constants";

export function generateConfigurations(
  includeFrontend: boolean,
  includeBackend: boolean,
  includeBot: boolean,
  includeOfficeAddin: boolean
): Record<string, unknown>[] {
  let edgeOrder = 2,
    chromeOrder = 1;
  if (os.type() === "Windows_NT") {
    edgeOrder = 1;
    chromeOrder = 2;
  }

  const launchConfigurations: Record<string, unknown>[] = [
    launchRemote(LaunchBrowser.edge, "Edge", edgeOrder),
    launchRemote(LaunchBrowser.chrome, "Chrome", chromeOrder),
  ];

  if (includeFrontend) {
    launchConfigurations.push(
      attachToFrontend(LaunchBrowser.edge, "Edge", includeBackend, includeBot)
    );
    launchConfigurations.push(
      attachToFrontend(LaunchBrowser.chrome, "Chrome", includeBackend, includeBot)
    );
  } else if (includeBot) {
    launchConfigurations.push(launchBot(LaunchBrowser.edge, "Edge", includeBackend));
    launchConfigurations.push(launchBot(LaunchBrowser.chrome, "Chrome", includeBackend));
  }

  if (includeBot) {
    launchConfigurations.push(attachToBot());
  }

  if (includeBackend) {
    launchConfigurations.push(attachToBackend());
  }

  if (isOfficeAddinEnabled() && includeOfficeAddin) {
    // add an entry to the launchConfigurations
    launchConfigurations.push({
      name: "Office Add-in sample",
      type: "pwa-msedge",
      url: "https://www.bing.com",
      presentation: {
        hidden: false,
        group: "office-addin",
        order: 1,
      },
    });
  }

  return launchConfigurations;
}

export function generateCompounds(
  includeFrontend: boolean,
  includeBackend: boolean,
  includeBot: boolean,
  includeOfficeAddin: boolean
): Record<string, unknown>[] {
  const launchCompounds: Record<string, unknown>[] = [];
  let edgeOrder = 2,
    chromeOrder = 1;
  if (os.type() === "Windows_NT") {
    edgeOrder = 1;
    chromeOrder = 2;
  }

  launchCompounds.push(debug(includeFrontend, includeBackend, includeBot, "Edge", edgeOrder));
  launchCompounds.push(debug(includeFrontend, includeBackend, includeBot, "Chrome", chromeOrder));

  if (isOfficeAddinEnabled() && includeOfficeAddin) {
    // add an entry to the launchCompounds
  }

  return launchCompounds;
}

export function generateM365Configurations(
  includeFrontend: boolean,
  includeBackend: boolean,
  includeBot: boolean
): Record<string, unknown>[] {
  const launchConfigurations: Record<string, unknown>[] = [];

  if (includeFrontend) {
    launchConfigurations.push(
      attachToFrontendM365(HubName.teams, LaunchBrowser.edge, "Edge", includeBackend, includeBot)
    );
    launchConfigurations.push(
      attachToFrontendM365(
        HubName.teams,
        LaunchBrowser.chrome,
        "Chrome",
        includeBackend,
        includeBot
      )
    );
    launchConfigurations.push(
      attachToFrontendM365(HubName.outlook, LaunchBrowser.edge, "Edge", includeBackend, includeBot)
    );
    launchConfigurations.push(
      attachToFrontendM365(
        HubName.outlook,
        LaunchBrowser.chrome,
        "Chrome",
        includeBackend,
        includeBot
      )
    );
    launchConfigurations.push(
      attachToFrontendM365(HubName.office, LaunchBrowser.edge, "Edge", includeBackend, includeBot)
    );
    launchConfigurations.push(
      attachToFrontendM365(
        HubName.office,
        LaunchBrowser.chrome,
        "Chrome",
        includeBackend,
        includeBot
      )
    );
  } else if (includeBot) {
    launchConfigurations.push(
      launchBotM365(HubName.teams, LaunchBrowser.edge, "Edge", includeBackend)
    );
    launchConfigurations.push(
      launchBotM365(HubName.teams, LaunchBrowser.chrome, "Chrome", includeBackend)
    );
    launchConfigurations.push(
      launchBotM365(HubName.outlook, LaunchBrowser.edge, "Edge", includeBackend)
    );
    launchConfigurations.push(
      launchBotM365(HubName.outlook, LaunchBrowser.chrome, "Chrome", includeBackend)
    );
  }

  if (includeBot) {
    launchConfigurations.push(attachToBot());
  }

  if (includeBackend) {
    launchConfigurations.push(attachToBackend());
  }

  return launchConfigurations;
}

export function generateM365Compounds(
  includeFrontend: boolean,
  includeBackend: boolean,
  includeBot: boolean
): Record<string, unknown>[] {
  const launchCompounds: Record<string, unknown>[] = [];
  let edgeOrder = 2,
    chromeOrder = 1;
  if (os.type() === "Windows_NT") {
    edgeOrder = 1;
    chromeOrder = 2;
  }

  launchCompounds.push(
    debugM365(includeFrontend, includeBackend, includeBot, HubName.teams, 1, "Edge", edgeOrder)
  );
  launchCompounds.push(
    debugM365(includeFrontend, includeBackend, includeBot, HubName.teams, 1, "Chrome", chromeOrder)
  );
  launchCompounds.push(
    debugM365(includeFrontend, includeBackend, includeBot, HubName.outlook, 2, "Edge", edgeOrder)
  );
  launchCompounds.push(
    debugM365(
      includeFrontend,
      includeBackend,
      includeBot,
      HubName.outlook,
      2,
      "Chrome",
      chromeOrder
    )
  );
  if (includeFrontend) {
    launchCompounds.push(
      debugM365(includeFrontend, includeBackend, includeBot, HubName.office, 3, "Edge", edgeOrder)
    );
    launchCompounds.push(
      debugM365(
        includeFrontend,
        includeBackend,
        includeBot,
        HubName.office,
        3,
        "Chrome",
        chromeOrder
      )
    );
  }

  return launchCompounds;
}

export function mergeLaunches(
  existingData: Record<string, unknown>,
  newData: Record<string, unknown>
): Record<string, unknown> {
  const mergedData = {} as Record<string, unknown>;
  Object.assign(mergedData, existingData);

  if (mergedData.version === undefined) {
    mergedData.version = "0.2.0";
  }

  if (mergedData.configurations === undefined) {
    mergedData.configurations = newData.configurations;
  } else {
    const existingConfigurations = mergedData.configurations as Record<string, unknown>[];
    const newConfigurations = (newData.configurations ?? []) as Record<string, unknown>[];
    const keptConfigurations = [];
    for (const existingConfiguration of existingConfigurations) {
      if (
        !newConfigurations.some(
          (newConfiguration) =>
            existingConfiguration.name === newConfiguration.name &&
            existingConfiguration.type === newConfiguration.type &&
            existingConfiguration.request === newConfiguration.request
        )
      ) {
        keptConfigurations.push(existingConfiguration);
      }
    }
    mergedData.configurations = [...keptConfigurations, ...newConfigurations];
  }

  if (mergedData.compounds === undefined) {
    mergedData.compounds = newData.compounds;
  } else {
    const existingCompounds = mergedData.compounds as Record<string, unknown>[];
    const newCompounds = (newData.compounds ?? []) as Record<string, unknown>[];
    const keptCompounds = [];
    for (const existingCompound of existingCompounds) {
      if (!newCompounds.some((newCompound) => existingCompound.name === newCompound.name)) {
        keptCompounds.push(existingCompound);
      }
    }
    mergedData.compounds = [...keptCompounds, ...newCompounds];
  }

  return mergedData;
}

function launchRemote(
  browserType: string,
  browserName: string,
  order: number
): Record<string, unknown> {
  return {
    name: `Launch Remote (${browserName})`,
    type: browserType,
    request: "launch",
    url: LaunchUrl.teamsRemote,
    presentation: {
      group: "remote",
      order: order,
    },
  };
}

function attachToFrontend(
  browserType: string,
  browserName: string,
  includeBackend: boolean,
  includeBot: boolean
): Record<string, unknown> {
  // NOTE: if no cascadeTerminateToConfigurations, closing browser will not stop
  // "Attach to Backend" and "Attach to Bot" even though stopAll in compound is true
  const cascadeTerminateToConfigurations = [];
  if (includeBackend) {
    cascadeTerminateToConfigurations.push("Attach to Backend");
  }
  if (includeBot) {
    cascadeTerminateToConfigurations.push("Attach to Bot");
  }
  return {
    name: `Attach to Frontend (${browserName})`,
    type: browserType,
    request: "launch",
    url: LaunchUrl.teamsLocal,
    cascadeTerminateToConfigurations:
      cascadeTerminateToConfigurations.length > 0 ? cascadeTerminateToConfigurations : undefined,
    presentation: {
      group: "all",
      hidden: true,
    },
  };
}

function attachToFrontendM365(
  hubName: string,
  browserType: string,
  browserName: string,
  includeBackend: boolean,
  includeBot: boolean
): Record<string, unknown> {
  const cascadeTerminateToConfigurations = [];
  if (includeBackend) {
    cascadeTerminateToConfigurations.push("Attach to Backend");
  }
  if (includeBot) {
    cascadeTerminateToConfigurations.push("Attach to Bot");
  }
  return {
    name: `Attach to Frontend in ${hubName} (${browserName})`,
    type: browserType,
    request: "launch",
    url: getFrontendLaunchUrl(true, hubName),
    cascadeTerminateToConfigurations:
      cascadeTerminateToConfigurations.length > 0 ? cascadeTerminateToConfigurations : undefined,
    presentation: {
      group: "all",
      hidden: true,
    },
  };
}

function attachToBot() {
  return {
    name: "Attach to Bot",
    type: "pwa-node",
    request: "attach",
    port: 9239,
    restart: true,
    presentation: {
      group: "all",
      hidden: true,
    },
  };
}

function attachToBackend(): Record<string, unknown> {
  return {
    name: "Attach to Backend",
    type: "pwa-node",
    request: "attach",
    port: 9229,
    restart: true,
    presentation: {
      group: "all",
      hidden: true,
    },
    internalConsoleOptions: "neverOpen",
  };
}

function launchBot(
  browserType: string,
  browserName: string,
  includeBackend: boolean
): Record<string, unknown> {
  const cascadeTerminateToConfigurations = ["Attach to Bot"];
  if (includeBackend) {
    cascadeTerminateToConfigurations.push("Attach to Backend");
  }
  return {
    name: `Launch Bot (${browserName})`,
    type: browserType,
    request: "launch",
    url: LaunchUrl.teamsLocal,
    cascadeTerminateToConfigurations,
    presentation: {
      group: "all",
      hidden: true,
    },
  };
}

function launchBotM365(
  hubName: string,
  browserType: string,
  browserName: string,
  includeBackend: boolean
): Record<string, unknown> {
  const cascadeTerminateToConfigurations = ["Attach to Bot"];
  if (includeBackend) {
    cascadeTerminateToConfigurations.push("Attach to Backend");
  }
  return {
    name: `Launch Bot in ${hubName} (${browserName})`,
    type: browserType,
    request: "launch",
    url: getBotLaunchUrl(true, hubName),
    cascadeTerminateToConfigurations,
    presentation: {
      group: "all",
      hidden: true,
    },
  };
}

function debug(
  includeFrontend: boolean,
  includeBackend: boolean,
  includeBot: boolean,
  browserName: string,
  order: number
): Record<string, unknown> {
  const configurations: string[] = [];
  if (includeFrontend) {
    configurations.push(`Attach to Frontend (${browserName})`);
  } else if (includeBot) {
    configurations.push(`Launch Bot (${browserName})`);
  }
  if (includeBot) {
    configurations.push("Attach to Bot");
  }
  if (includeBackend) {
    configurations.push("Attach to Backend");
  }
  return {
    name: `Debug (${browserName})`,
    configurations,
    preLaunchTask: "Pre Debug Check & Start All",
    presentation: {
      group: "all",
      order: order,
    },
    stopAll: true,
  };
}

function debugM365(
  includeFrontend: boolean,
  includeBackend: boolean,
  includeBot: boolean,
  hubName: string,
  hubOrder: number,
  browserName: string,
  order: number
): Record<string, unknown> {
  const configurations: string[] = [];
  if (includeFrontend) {
    configurations.push(`Attach to Frontend in ${hubName} (${browserName})`);
  } else if (includeBot) {
    configurations.push(`Launch Bot in ${hubName} (${browserName})`);
  }
  if (includeBot) {
    configurations.push("Attach to Bot");
  }
  if (includeBackend) {
    configurations.push("Attach to Backend");
  }
  return {
    name: `Debug in ${hubName} (${browserName})`,
    configurations,
    preLaunchTask:
      hubName === HubName.teams
        ? "Pre Debug Check & Start All"
        : "Pre Debug Check & Start All & Install App",
    presentation: {
      group: `group ${hubOrder}: ${hubName}`,
      order: order,
    },
    stopAll: true,
  };
}

function getFrontendLaunchUrl(isLocal: boolean, hubName: string) {
  if (hubName === HubName.teams) {
    return isLocal ? LaunchUrl.teamsLocal : LaunchUrl.teamsRemote;
  } else if (hubName === HubName.outlook) {
    return isLocal ? LaunchUrl.outlookLocalTab : LaunchUrl.outlookRemoteTab;
  } else if (hubName === HubName.office) {
    return isLocal ? LaunchUrl.officeLocalTab : LaunchUrl.officeRemoteTab;
  }
  return "";
}

function getBotLaunchUrl(isLocal: boolean, hubName: string): string {
  if (hubName === HubName.teams) {
    return isLocal ? LaunchUrl.teamsLocal : LaunchUrl.teamsRemote;
  } else if (hubName === HubName.outlook) {
    return isLocal ? LaunchUrl.outlookLocalBot : LaunchUrl.outlookRemoteBot;
  }
  return "";
}
