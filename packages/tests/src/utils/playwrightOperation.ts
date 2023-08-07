// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
import { BrowserContext, Page, Frame } from "playwright";
import { assert } from "chai";
import { Timeout, ValidationContent, TemplateProject } from "./constants";
import { RetryHandler } from "./retryHandler";
import { getPlaywrightScreenshotPath } from "./nameUtil";
import axios from "axios";
import { SampledebugContext } from "../ui-test/samples/sampledebugContext";
import path from "path";
import fs from "fs";
import { dotenvUtil } from "./envUtil";
import { startDebugging } from "./vscodeOperation";
import { editDotEnvFile } from "./commonUtils";
import { AzSqlHelper } from "./azureCliHelper";
import { expect } from "chai";
import * as uuid from "uuid";

export const debugInitMap: Record<TemplateProject, () => Promise<void>> = {
  [TemplateProject.AdaptiveCard]: async () => {
    await startDebugging();
  },
  [TemplateProject.AssistDashboard]: async () => {
    await startDebugging("Debug in Teams (Chrome)");
  },
  [TemplateProject.ContactExporter]: async () => {
    await startDebugging();
  },
  [TemplateProject.Dashboard]: async () => {
    await startDebugging();
  },
  [TemplateProject.GraphConnector]: async () => {
    await startDebugging();
  },
  [TemplateProject.OutlookTab]: async () => {
    await startDebugging("Debug in Teams (Chrome)");
  },
  [TemplateProject.HelloWorldTabBackEnd]: async () => {
    await startDebugging();
  },
  [TemplateProject.MyFirstMetting]: async () => {
    await startDebugging();
  },
  [TemplateProject.HelloWorldBotSSO]: async () => {
    await startDebugging();
  },
  [TemplateProject.IncomingWebhook]: async () => {
    await startDebugging("Attach to Incoming Webhook");
  },
  [TemplateProject.NpmSearch]: async () => {
    await startDebugging("Debug in Teams (Chrome)");
  },
  [TemplateProject.OneProductivityHub]: async () => {
    await startDebugging();
  },
  [TemplateProject.ProactiveMessaging]: async () => {
    await startDebugging();
  },
  [TemplateProject.QueryOrg]: async () => {
    await startDebugging();
  },
  [TemplateProject.ShareNow]: async () => {
    await startDebugging();
  },
  [TemplateProject.StockUpdate]: async () => {
    await startDebugging();
  },
  [TemplateProject.TodoListBackend]: async () => {
    await startDebugging();
  },
  [TemplateProject.TodoListM365]: async () => {
    await startDebugging("Debug in Teams (Chrome)");
  },
  [TemplateProject.TodoListSpfx]: async () => {
    await startDebugging("Teams workbench (Chrome)");
  },
  [TemplateProject.Deeplinking]: async () => {
    await startDebugging();
  },
  [TemplateProject.DiceRoller]: async () => {
    await startDebugging();
  },
  [TemplateProject.OutlookSignature]: async () => {
    await startDebugging();
  },
  [TemplateProject.ChefBot]: async () => {
    await startDebugging();
  },
};

export const middleWareMap: Record<
  TemplateProject,
  (
    sampledebugContext: SampledebugContext,
    env: "local" | "dev",
    azSqlHelper?: AzSqlHelper,
    steps?: {
      before?: boolean;
      after?: boolean;
      afterCreate?: boolean;
      afterdeploy?: boolean;
    }
  ) => Promise<void | AzSqlHelper>
> = {
  [TemplateProject.AdaptiveCard]: () => Promise.resolve(),
  [TemplateProject.AssistDashboard]: async (
    sampledebugContext: SampledebugContext,
    env: "local" | "dev",
    azSqlHelper?: AzSqlHelper,
    steps?: {
      afterCreate?: boolean;
    }
  ) => {
    assistantDashboardMiddleWare(sampledebugContext, env, azSqlHelper, {
      afterCreate: steps?.afterCreate,
    });
  },
  [TemplateProject.ContactExporter]: () => Promise.resolve(),
  [TemplateProject.Dashboard]: () => Promise.resolve(),
  [TemplateProject.GraphConnector]: () => Promise.resolve(),
  [TemplateProject.OutlookTab]: () => Promise.resolve(),
  [TemplateProject.HelloWorldTabBackEnd]: () => Promise.resolve(),
  [TemplateProject.MyFirstMetting]: () => Promise.resolve(),
  [TemplateProject.HelloWorldBotSSO]: () => Promise.resolve(),
  [TemplateProject.IncomingWebhook]: async (
    sampledebugContext: SampledebugContext,
    env: "local" | "dev",
    azSqlHelper?: AzSqlHelper,
    steps?: {
      afterCreate?: boolean;
    }
  ) => {
    await incomingWebhookMiddleWare(sampledebugContext, env, azSqlHelper, {
      afterCreate: steps?.afterCreate,
    });
  },
  [TemplateProject.NpmSearch]: () => Promise.resolve(),
  [TemplateProject.OneProductivityHub]: () => Promise.resolve(),
  [TemplateProject.ProactiveMessaging]: () => Promise.resolve(),
  [TemplateProject.QueryOrg]: () => Promise.resolve(),
  [TemplateProject.ShareNow]: async (
    sampledebugContext: SampledebugContext,
    env: "local" | "dev",
    azSqlHelper?: AzSqlHelper,
    steps?: {
      before?: boolean;
      afterCreate?: boolean;
    }
  ) => {
    return await shareNowMiddleWare(sampledebugContext, env, azSqlHelper, {
      before: steps?.before,
      afterCreate: steps?.afterCreate,
    });
  },
  [TemplateProject.StockUpdate]: async (
    sampledebugContext: SampledebugContext,
    env: "local" | "dev",
    azSqlHelper?: AzSqlHelper,
    steps?: {
      afterCreate?: boolean;
    }
  ) => {
    await stockUpdateMiddleWare(sampledebugContext, env, azSqlHelper, {
      afterCreate: steps?.afterCreate,
    });
  },
  [TemplateProject.TodoListBackend]: async (
    sampledebugContext: SampledebugContext,
    env: "local" | "dev",
    azSqlHelper?: AzSqlHelper,
    steps?: {
      before?: boolean;
      afterCreate?: boolean;
    }
  ) => {
    return await todoListSqlMiddleWare(sampledebugContext, env, azSqlHelper, {
      before: steps?.before,
      afterCreate: steps?.afterCreate,
    });
  },
  [TemplateProject.TodoListM365]: async (
    sampledebugContext: SampledebugContext,
    env: "local" | "dev",
    azSqlHelper?: AzSqlHelper,
    steps?: {
      afterCreate?: boolean;
    }
  ) => {
    await TodoListM365MiddleWare(sampledebugContext, env, azSqlHelper, {
      afterCreate: steps?.afterCreate,
    });
  },
  [TemplateProject.TodoListSpfx]: () => Promise.resolve(),
  [TemplateProject.Deeplinking]: () => Promise.resolve(),
  [TemplateProject.DiceRoller]: () => Promise.resolve(),
  [TemplateProject.OutlookSignature]: () => Promise.resolve(),
  [TemplateProject.ChefBot]: () => Promise.resolve(),
};

export async function initPage(
  context: BrowserContext,
  teamsAppId: string,
  username: string,
  password: string,
  options?: {
    teamsAppName?: string;
    dashboardFlag?: boolean;
  }
): Promise<Page> {
  let page = await context.newPage();
  page.setDefaultTimeout(Timeout.playwrightDefaultTimeout);

  // open teams app page
  // https://github.com/puppeteer/puppeteer/issues/3338
  await Promise.all([
    page.goto(
      `https://teams.microsoft.com/_#/l/app/${teamsAppId}?installAppPackage=true`
    ),
    page.waitForNavigation(),
  ]);

  // input username
  await RetryHandler.retry(async () => {
    await page.fill("input.input[type='email']", username);
    console.log(`fill in username ${username}`);

    // next
    await Promise.all([
      page.click("input.button[type='submit']"),
      page.waitForNavigation(),
    ]);
  });

  // input password
  console.log(`fill in password`);
  await page.fill("input.input[type='password'][name='passwd']", password);

  // sign in
  await Promise.all([
    page.click("input.button[type='submit']"),
    page.waitForNavigation(),
  ]);

  // stay signed in confirm page
  console.log(`stay signed confirm`);
  await Promise.all([
    page.click("input.button[type='submit'][value='Yes']"),
    page.waitForNavigation(),
  ]);
  await page.waitForTimeout(Timeout.shortTimeLoading);

  // add app
  await RetryHandler.retry(async (retries: number) => {
    if (retries > 0) {
      console.log(`Retried to run adding app for ${retries} times.`);
    }
    await page.close();
    console.log(`open teams page`);
    page = await context.newPage();
    await Promise.all([
      page.goto(
        `https://teams.microsoft.com/_#/l/app/${teamsAppId}?installAppPackage=true`
      ),
      page.waitForNavigation(),
    ]);
    await page.waitForTimeout(Timeout.longTimeWait);
    console.log("click add button");

    const frameElementHandle = await page.waitForSelector(
      "iframe.embedded-page-content"
    );
    const frame = await frameElementHandle?.contentFrame();
    const addBtn = await frame?.waitForSelector("button span:has-text('Add')");

    // dashboard template will have a popup
    if (options?.dashboardFlag) {
      console.log("Before popup");
      const [popup] = await Promise.all([
        page
          .waitForEvent("popup")
          .then((popup) =>
            popup
              .waitForEvent("close", {
                timeout: Timeout.playwrightConsentPopupPage,
              })
              .catch(() => popup)
          )
          .catch(() => {}),
        addBtn?.click(),
      ]);
      console.log("after popup");

      if (popup && !popup?.isClosed()) {
        // input password
        console.log(`fill in password`);
        await popup.fill(
          "input.input[type='password'][name='passwd']",
          password
        );
        // sign in
        await Promise.all([
          popup.click("input.button[type='submit'][value='Sign in']"),
          popup.waitForNavigation(),
        ]);
        await popup.click("input.button[type='submit'][value='Accept']");
      }
    } else {
      await addBtn?.click();
    }
    await page.waitForTimeout(Timeout.shortTimeLoading);
    // verify add page is closed
    await frame?.waitForSelector("button span:has-text('Add')", {
      state: "detached",
    });
    try {
      try {
        await page?.waitForSelector(".team-information span:has-text('About')");
      } catch (error) {
        await page?.waitForSelector(
          ".ts-messages-header span:has-text('About')"
        );
      }
      console.log("[success] app loaded");
    } catch (error) {
      await page.screenshot({
        path: getPlaywrightScreenshotPath("error"),
        fullPage: true,
      });
      assert.fail("[Error] add app failed");
    }
    await page.waitForTimeout(Timeout.shortTimeLoading);
  });

  return page;
}

export async function initTeamsPage(
  context: BrowserContext,
  teamsAppId: string,
  username: string,
  password: string,
  options?: {
    teamsAppName?: string;
    dashboardFlag?: boolean;
    type?: string;
  }
): Promise<Page> {
  let page = await context.newPage();
  try {
    page.setDefaultTimeout(Timeout.playwrightDefaultTimeout);

    // open teams app page
    // https://github.com/puppeteer/puppeteer/issues/3338
    await Promise.all([
      page.goto(
        `https://teams.microsoft.com/_#/l/app/${teamsAppId}?installAppPackage=true`
      ),
      page.waitForNavigation(),
    ]);

    // input username
    await RetryHandler.retry(async () => {
      await page.fill("input.input[type='email']", username);
      console.log(`fill in username ${username}`);

      // next
      await Promise.all([
        page.click("input.button[type='submit']"),
        page.waitForNavigation(),
      ]);
    });

    // input password
    console.log(`fill in password`);
    await page.fill("input.input[type='password'][name='passwd']", password);

    // sign in
    await Promise.all([
      page.click("input.button[type='submit']"),
      page.waitForNavigation(),
    ]);

    // stay signed in confirm page
    console.log(`stay signed confirm`);
    await Promise.all([
      page.click("input.button[type='submit'][value='Yes']"),
      page.waitForNavigation(),
    ]);

    // add app
    await RetryHandler.retry(async (retries: number) => {
      if (retries > 0) {
        console.log(`Retried to run adding app for ${retries} times.`);
      }
      await page.close();
      console.log(`open teams page`);
      page = await context.newPage();
      await Promise.all([
        page.goto(
          `https://teams.microsoft.com/_#/l/app/${teamsAppId}?installAppPackage=true`
        ),
        page.waitForNavigation(),
      ]);
      await page.waitForTimeout(Timeout.longTimeWait);
      console.log("click add button");
      const frameElementHandle = await page.waitForSelector(
        "iframe.embedded-page-content"
      );
      const frame = await frameElementHandle?.contentFrame();

      try {
        console.log("dismiss message");
        await page.click('button:has-text("Dismiss")');
      } catch (error) {
        console.log("no message to dismiss");
      }
      // default
      const addBtn = await frame?.waitForSelector(
        "button span:has-text('Add')"
      );
      await addBtn?.click();
      await page.waitForTimeout(Timeout.shortTimeLoading);

      if (options?.type === "meeting") {
        // verify add page is closed
        const frameElementHandle = await page.waitForSelector(
          "iframe.embedded-page-content"
        );
        const frame = await frameElementHandle?.contentFrame();
        await frame?.waitForSelector(
          `h1:has-text('Add ${options?.teamsAppName} to a team')`
        );
        // TODO: need to add more logic
        console.log("successful to add teams app!!!");
        return;
      }

      // verify add page is closed
      await frame?.waitForSelector(
        `h1:has-text('Add ${options?.teamsAppName} to a team')`
      );

      try {
        const frameElementHandle = await page.waitForSelector(
          "iframe.embedded-page-content"
        );
        const frame = await frameElementHandle?.contentFrame();

        try {
          const items = await frame?.waitForSelector("li.ui-dropdown__item");
          await items?.click();
        } catch (error) {
          const searchBtn = await frame?.waitForSelector(
            "div.ui-dropdown__toggle-indicator"
          );
          await searchBtn?.click();
          await page.waitForTimeout(Timeout.shortTimeLoading);
          const items = await frame?.waitForSelector("li.ui-dropdown__item");
          await items?.click();
        }

        const setUpBtn = await frame?.waitForSelector(
          'button span:has-text("Set up a tab")'
        );
        await setUpBtn?.click();
        await page.waitForTimeout(Timeout.shortTimeLoading);
      } catch (error) {
        await page.screenshot({
          path: getPlaywrightScreenshotPath("error"),
          fullPage: true,
        });
        throw error;
      }
      {
        const frameElementHandle = await page.waitForSelector(
          "iframe.embedded-iframe"
        );
        const frame = await frameElementHandle?.contentFrame();
        if (options?.type === "spfx") {
          try {
            console.log("Load debug scripts");
            await frame?.click('button:has-text("Load debug scripts")');
            console.log("Debug scripts loaded");
          } catch (error) {
            console.log("No debug scripts to load");
          }
        }
        try {
          const saveBtn = await page.waitForSelector(`button:has-text("Save")`);
          await saveBtn?.click();
          await page.waitForSelector(`button:has-text("Save")`, {
            state: "detached",
          });
        } catch (error) {
          console.log("No save button to click");
        }
      }
      await page.waitForTimeout(Timeout.shortTimeLoading);
      console.log("successful to add teams app!!!");
    });

    return page;
  } catch (error) {
    await page.screenshot({
      path: getPlaywrightScreenshotPath("error"),
      fullPage: true,
    });
    throw error;
  }
}

export async function validateOneProducitvity(
  page: Page,
  options?: { displayName?: string }
) {
  try {
    console.log("start to verify One Productivity Hub");
    const frameElementHandle = await page.waitForSelector(
      "iframe.embedded-iframe"
    );
    const frame = await frameElementHandle?.contentFrame();
    try {
      console.log("dismiss message");
      await page
        .click('button:has-text("Dismiss")', {
          timeout: Timeout.playwrightDefaultTimeout,
        })
        .catch(() => {});
    } catch (error) {
      console.log("no message to dismiss");
    }
    try {
      const startBtn = await frame?.waitForSelector(
        'button:has-text("Start One Productivity Hub")'
      );
      console.log("click Start One Productivity Hub button");
      await RetryHandler.retry(async () => {
        console.log("Before popup");
        const [popup] = await Promise.all([
          page
            .waitForEvent("popup")
            .then((popup) =>
              popup
                .waitForEvent("close", {
                  timeout: Timeout.playwrightConsentPopupPage,
                })
                .catch(() => popup)
            )
            .catch(() => {}),
          startBtn?.click(),
        ]);
        console.log("after popup");

        if (popup && !popup?.isClosed()) {
          await popup
            .click('button:has-text("Reload")', {
              timeout: Timeout.playwrightConsentPageReload,
            })
            .catch(() => {});
          await popup.click("input.button[type='submit'][value='Accept']");
        }
        await frame?.waitForSelector(`div:has-text("${options?.displayName}")`);
        // TODO: need to add more logic
      });
    } catch (e: any) {
      console.log(`[Command not executed successfully] ${e.message}`);
      await page.screenshot({
        path: getPlaywrightScreenshotPath("error"),
        fullPage: true,
      });
      throw e;
    }

    await page.waitForTimeout(Timeout.shortTimeLoading);
  } catch (error) {
    await page.screenshot({
      path: getPlaywrightScreenshotPath("error"),
      fullPage: true,
    });
    throw error;
  }
}

export async function validateTab(
  page: Page,
  options?: { displayName?: string; includeFunction?: boolean }
) {
  try {
    const frameElementHandle = await page.waitForSelector(
      "iframe.embedded-iframe"
    );
    const frame = await frameElementHandle?.contentFrame();

    await RetryHandler.retry(async () => {
      console.log("Before popup");
      const [popup] = await Promise.all([
        page
          .waitForEvent("popup")
          .then((popup) =>
            popup
              .waitForEvent("close", {
                timeout: Timeout.playwrightConsentPopupPage,
              })
              .catch(() => popup)
          )
          .catch(() => {}),
        frame?.click('button:has-text("Authorize")', {
          timeout: Timeout.playwrightAddAppButton,
          force: true,
          noWaitAfter: true,
          clickCount: 2,
          delay: 10000,
        }),
      ]);
      console.log("after popup");

      if (popup && !popup?.isClosed()) {
        await popup
          .click('button:has-text("Reload")', {
            timeout: Timeout.playwrightConsentPageReload,
          })
          .catch(() => {});
        await popup.click("input.button[type='submit'][value='Accept']");
      }

      await frame?.waitForSelector(`b:has-text("${options?.displayName}")`);
    });

    if (options?.includeFunction) {
      await RetryHandler.retry(async () => {
        console.log("verify function info");
        const authorizeButton = await frame?.waitForSelector(
          'button:has-text("Call Azure Function")'
        );
        await authorizeButton?.click();
        const backendElement = await frame?.waitForSelector(
          'pre:has-text("receivedHTTPRequestBody")'
        );
        const content = await backendElement?.innerText();
        if (!content?.includes("User display name is"))
          assert.fail("User display name is not found in the response");
        console.log("verify function info success");
      });
    }
  } catch (error) {
    await page.screenshot({
      path: getPlaywrightScreenshotPath("error"),
      fullPage: true,
    });
    throw error;
  }
}

export async function validateReactTab(
  page: Page,
  displayName: string,
  includeFunction?: boolean
) {
  try {
    const frameElementHandle = await page.waitForSelector(
      "iframe.embedded-iframe"
    );
    const frame = await frameElementHandle?.contentFrame();
    if (includeFunction) {
      await RetryHandler.retry(async () => {
        console.log("Before popup");
        const [popup] = await Promise.all([
          page
            .waitForEvent("popup")
            .then((popup) =>
              popup
                .waitForEvent("close", {
                  timeout: Timeout.playwrightConsentPopupPage,
                })
                .catch(() => popup)
            )
            .catch(() => {}),
          frame?.click('button:has-text("Call Azure Function")', {
            timeout: Timeout.playwrightAddAppButton,
            force: true,
            noWaitAfter: true,
            clickCount: 2,
            delay: 10000,
          }),
        ]);
        console.log("after popup");

        if (popup && !popup?.isClosed()) {
          await popup
            .click('button:has-text("Reload")', {
              timeout: Timeout.playwrightConsentPageReload,
            })
            .catch(() => {});
          await popup.click("input.button[type='submit'][value='Accept']");
        }
      });

      console.log("verify function info");
      const backendElement = await frame?.waitForSelector(
        'pre:has-text("receivedHTTPRequestBody")'
      );
      const content = await backendElement?.innerText();
      if (!content?.includes("User display name is"))
        assert.fail("User display name is not found in the response");
      console.log("verify function info success");
    }

    await frame?.waitForSelector(`b:has-text("${displayName}")`);
  } catch (error) {
    await page.screenshot({
      path: getPlaywrightScreenshotPath("error"),
      fullPage: true,
    });
    throw error;
  }
}

export async function validateReactOutlookTab(
  page: Page,
  displayName: string,
  includeFunction?: boolean
) {
  try {
    await page.waitForTimeout(Timeout.longTimeWait);
    const frameElementHandle = await page.waitForSelector(
      'iframe[data-tid="app-host-iframe"]'
    );
    const frame = await frameElementHandle?.contentFrame();
    if (includeFunction) {
      await RetryHandler.retry(async () => {
        console.log("Before popup");
        const [popup] = await Promise.all([
          page
            .waitForEvent("popup")
            .then((popup) =>
              popup
                .waitForEvent("close", {
                  timeout: Timeout.playwrightConsentPopupPage,
                })
                .catch(() => popup)
            )
            .catch(() => {}),
          frame?.click('button:has-text("Call Azure Function")', {
            timeout: Timeout.playwrightAddAppButton,
            force: true,
            noWaitAfter: true,
            clickCount: 2,
            delay: 10000,
          }),
        ]);
        console.log("after popup");

        if (popup && !popup?.isClosed()) {
          await popup
            .click('button:has-text("Reload")', {
              timeout: Timeout.playwrightConsentPageReload,
            })
            .catch(() => {});
          await popup.click("input.button[type='submit'][value='Accept']");
        }
      });

      console.log("verify function info");
      const backendElement = await frame?.waitForSelector(
        'pre:has-text("receivedHTTPRequestBody")'
      );
      const content = await backendElement?.innerText();
      if (!content?.includes("User display name is"))
        assert.fail("User display name is not found in the response");
      console.log("verify function info success");
    }

    await frame?.waitForSelector(`b:has-text("${displayName}")`);
  } catch (error) {
    await page.screenshot({
      path: getPlaywrightScreenshotPath("error"),
      fullPage: true,
    });
    throw error;
  }
}

export async function validateBasicTab(
  page: Page,
  content = "Hello, World",
  hubState = "Teams"
) {
  try {
    const frameElementHandle = await page.waitForSelector(
      "iframe.embedded-iframe"
    );
    const frame = await frameElementHandle?.contentFrame();
    console.log(`Check if ${content} showed`);
    await frame?.waitForSelector(`h1:has-text("${content}")`);
    console.log(`Check if ${hubState} showed`);
    await frame?.waitForSelector(`#hubState:has-text("${hubState}")`);
    console.log(`${hubState} showed`);
  } catch (error) {
    await page.screenshot({
      path: getPlaywrightScreenshotPath("error"),
      fullPage: true,
    });
    throw error;
  }
}

export async function validateTabNoneSSO(
  page: Page,
  content = "Congratulations",
  content2 = "Add Single Sign On feature to retrieve user profile"
) {
  try {
    const frameElementHandle = await page.waitForSelector(
      "iframe.embedded-iframe"
    );
    const frame = await frameElementHandle?.contentFrame();
    console.log(`Check if ${content} showed`);
    await frame?.waitForSelector(`h1:has-text("${content}")`);
    console.log(`Check if ${content2} showed`);
    await frame?.waitForSelector(`h2:has-text("${content2}")`);
    console.log(`${content2} showed`);
  } catch (error) {
    await page.screenshot({
      path: getPlaywrightScreenshotPath("error"),
      fullPage: true,
    });
    throw error;
  }
}

export async function validatePersonalTab(page: Page) {
  try {
    const frameElementHandle = await page.waitForSelector(
      "iframe.embedded-iframe"
    );
    const frame = await frameElementHandle?.contentFrame();
    console.log(`Check if Congratulations showed`);
    await frame?.waitForSelector(`h1:has-text("Congratulations!")`);
    console.log(`Check tab 1 content`);
    await frame?.waitForSelector(`h2:has-text("Change this code")`);
    console.log(`Check tab 2 content`);
    const tab1 = await frame?.waitForSelector(
      `span:has-text("2. Provision and Deploy to the Cloud")`
    );
    await tab1?.click();
    {
      const frameElementHandle = await page.waitForSelector(
        "iframe.embedded-iframe"
      );
      const frame = await frameElementHandle?.contentFrame();
      await frame?.waitForSelector(`h2:has-text("Deploy to the Cloud")`);
    }
    console.log(`debug finish!`);
  } catch (error) {
    await page.screenshot({
      path: getPlaywrightScreenshotPath("error"),
      fullPage: true,
    });
    throw error;
  }
}

export async function validateOutlookTab(
  page: Page,
  displayName: string,
  includeFunction?: boolean
) {
  try {
    const frameElementHandle = await page.waitForSelector(
      'iframe[data-tid="app-host-iframe"]'
    );
    const frame = await frameElementHandle?.contentFrame();

    console.log("Before popup");
    const [popup] = await Promise.all([
      page
        .waitForEvent("popup")
        .then((popup) =>
          popup
            .waitForEvent("close", {
              timeout: Timeout.playwrightConsentPopupPage,
            })
            .catch(() => popup)
        )
        .catch(() => {}),
      frame?.click('button:has-text("Authorize")', {
        timeout: Timeout.playwrightAddAppButton,
        force: true,
        noWaitAfter: true,
        clickCount: 2,
        delay: 10000,
      }),
    ]);
    console.log("after popup");

    if (popup && !popup?.isClosed()) {
      await popup
        .click('button:has-text("Reload")', {
          timeout: Timeout.playwrightConsentPageReload,
        })
        .catch(() => {});
      await popup.click("input.button[type='submit'][value='Accept']");
    }

    await frame?.waitForSelector(`span:has-text("${displayName}")`);

    if (includeFunction) {
      await RetryHandler.retry(async () => {
        const authorizeButton = await frame?.waitForSelector(
          'button:has-text("Call Azure Function")'
        );
        await authorizeButton?.click();
        const backendElement = await frame?.waitForSelector(
          'pre:has-text("receivedHTTPRequestBody")'
        );
        const content = await backendElement?.innerText();
        // TODO validate content
      });
    }
  } catch (error) {
    await page.screenshot({
      path: getPlaywrightScreenshotPath("error"),
      fullPage: true,
    });
    throw error;
  }
}

export async function validateBot(
  page: Page,
  options?: { botCommand?: string; expected?: ValidationContent }
) {
  try {
    console.log("start to verify bot");
    await page.waitForTimeout(Timeout.shortTimeLoading);
    const frameElementHandle = await page.waitForSelector(
      "iframe.embedded-page-content"
    );
    const frame = await frameElementHandle?.contentFrame();
    try {
      console.log("dismiss message");
      await frame?.waitForSelector("div.ui-box");
      await page
        .click('button:has-text("Dismiss")', {
          timeout: Timeout.playwrightDefaultTimeout,
        })
        .catch(() => {});
    } catch (error) {
      console.log("no message to dismiss");
    }
    try {
      console.log("sending message ", options?.botCommand || "helloWorld");
      await executeBotSuggestionCommand(
        page,
        frame,
        options?.botCommand || "helloWorld"
      );
      await frame?.click('button[name="send"]');
    } catch (e: any) {
      console.log(
        `[Command "${options?.botCommand}" not executed successfully] ${e.message}`
      );
    }
    if (options?.botCommand === "show") {
      await RetryHandler.retry(async () => {
        // wait for alert message to show
        const btn = await frame?.waitForSelector(
          `div.ui-box button:has-text("Continue")`
        );
        await btn?.click();
        // wait for new tab to show
        const popup = await page
          .waitForEvent("popup")
          .then((popup) =>
            popup
              .waitForEvent("close", {
                timeout: Timeout.playwrightConsentPopupPage,
              })
              .catch(() => popup)
          )
          .catch(() => {});
        if (popup && !popup?.isClosed()) {
          await popup
            .click('button:has-text("Reload")', {
              timeout: Timeout.playwrightConsentPageReload,
            })
            .catch(() => {});
          await popup.click("input.button[type='submit'][value='Accept']");
        }
        await RetryHandler.retry(async () => {
          await frame?.waitForSelector(`p:has-text("${options?.expected}")`);
          console.log("verify bot successfully!!!");
        }, 2);
        console.log(`${options?.expected}`);
      }, 2);
      console.log(`${options?.expected}`);
    } else {
      await RetryHandler.retry(async () => {
        await frame?.waitForSelector(
          `p:has-text("${options?.expected || ValidationContent.Bot}")`
        );
        console.log("verify bot successfully!!!");
      }, 2);
      console.log(`${options?.expected}`);
    }
    await page.waitForTimeout(Timeout.shortTimeLoading);
  } catch (error) {
    await page.screenshot({
      path: getPlaywrightScreenshotPath("error"),
      fullPage: true,
    });
    throw error;
  }
}

export async function validateNpm(page: Page, options?: { npmName?: string }) {
  try {
    console.log("start to verify npm search");
    await page.waitForTimeout(Timeout.shortTimeLoading);
    const frameElementHandle = await page.waitForSelector(
      "iframe.embedded-page-content"
    );
    const frame = await frameElementHandle?.contentFrame();
    try {
      console.log("dismiss message");
      await frame?.waitForSelector("div.ui-box");
      await page
        .click('button:has-text("Dismiss")', {
          timeout: Timeout.playwrightDefaultTimeout,
        })
        .catch(() => {});
    } catch (error) {
      console.log("no message to dismiss");
    }
    console.log("search npm ", options?.npmName);
    const input = await frame?.waitForSelector("div.ui-box input.ui-box");
    await input?.type(options?.npmName || "axios");
    try {
      const targetItem = await frame?.waitForSelector(
        `span:has-text("${options?.npmName}")`
      );
      await targetItem?.click();
      await frame?.waitForSelector(`card span:has-text("${options?.npmName}")`);
      console.log("verify npm search successfully!!!");
      await page.waitForTimeout(Timeout.shortTimeLoading);
    } catch (error) {
      await frame?.waitForSelector(
        'div.ui-box span:has-text("Unable to reach app. Please try again.")'
      );
      assert.fail("Unable to reach app. Please try again.");
    }
  } catch (error) {
    await page.screenshot({
      path: getPlaywrightScreenshotPath("error"),
      fullPage: true,
    });
    throw error;
  }
}

export async function validateDeeplinking(page: Page, displayName: string) {
  try {
    console.log("start to verify deeplinking");
    await page.waitForTimeout(Timeout.shortTimeLoading);
    try {
      console.log("dismiss message");
      await page
        ?.click('div:has-text("Dismiss")', {
          timeout: Timeout.playwrightDefaultTimeout,
        })
        .catch(() => {});
    } catch (error) {
      console.log("no message to dismiss");
    }

    const frameElementHandle = await page.waitForSelector(
      "iframe.embedded-page-content"
    );
    const frame = await frameElementHandle?.contentFrame();
    await frame?.waitForSelector('h1:has-text("Congratulations!")');

    // verify tab navigate within app tab
    await page.waitForTimeout(Timeout.shortTimeLoading);
    try {
      const navigateBtn = await page?.waitForSelector(
        'li div a span:has-text("Navigate within app")'
      );
      await navigateBtn?.click();
      await page.waitForTimeout(Timeout.shortTimeLoading);
      const frameElementHandle = await page.waitForSelector(
        "iframe.embedded-page-content"
      );
      const frame = await frameElementHandle?.contentFrame();
      await frame?.waitForSelector(
        'div.welcome div.main-section div#navigate-within-app h2:has-text("2. Navigate within the app")'
      );
      console.log("navigate within app tab found");
    } catch (error) {
      console.log("navigate within app tab verify failed");
      await page.screenshot({
        path: getPlaywrightScreenshotPath("error"),
        fullPage: true,
      });
      throw error;
    }

    // verify details tab
    await page.waitForTimeout(Timeout.shortTimeLoading);
    try {
      const detailsBtn = await page?.waitForSelector(
        'li div a span:has-text("Details Tab")'
      );
      await RetryHandler.retry(async () => {
        await detailsBtn?.click();
        await page.waitForTimeout(Timeout.shortTimeLoading);
        const frameElementHandle = await page.waitForSelector(
          "iframe.embedded-page-content"
        );
        const frame = await frameElementHandle?.contentFrame();
        await frame?.waitForSelector('li a span:has-text("Tab 1")');
        console.log("details tab found");
      });
    } catch (error) {
      console.log("details tab verify failed");
      await page.screenshot({
        path: getPlaywrightScreenshotPath("error"),
        fullPage: true,
      });
      throw error;
    }

    // verify navigate within hub tab
    await page.waitForTimeout(Timeout.shortTimeLoading);
    try {
      const navigateHubBtn = await page?.waitForSelector(
        'li div a span:has-text("Navigate within hub")'
      );
      await RetryHandler.retry(async () => {
        await navigateHubBtn?.click();
        await page.waitForTimeout(Timeout.shortTimeLoading);
        const frameElementHandle = await page.waitForSelector(
          "iframe.embedded-page-content"
        );
        const frame = await frameElementHandle?.contentFrame();
        await frame?.waitForSelector(
          'h1.center:has-text("Chat functionality")'
        );
        console.log("navigate within hub tab found");
      });
      // TODO: add person
    } catch (error) {
      console.log("navigate within hub tab verify failed");
      await page.screenshot({
        path: getPlaywrightScreenshotPath("error"),
        fullPage: true,
      });
      throw error;
    }

    // verify generate deeplink tab
    try {
      const shareBtn = await page?.waitForSelector(
        'li div a span:has-text("Generate Share URL")'
      );
      await RetryHandler.retry(async () => {
        await shareBtn?.click();
        await page.waitForTimeout(Timeout.shortTimeLoading);
        const frameElementHandle = await page.waitForSelector(
          "iframe.embedded-page-content"
        );
        const frame = await frameElementHandle?.contentFrame();
        await frame?.waitForSelector('span:has-text("Copy a link to ")');
        console.log("popup message found");
        const closeBtn = await frame?.waitForSelector(
          "dev.close-container button.icons-close"
        );
        await closeBtn?.click();
      });
    } catch (error) {
      await page.screenshot({
        path: getPlaywrightScreenshotPath("error"),
        fullPage: true,
      });
      throw error;
    }
    console.log("verify deeplinking successfully!!!");
  } catch (error) {
    await page.screenshot({
      path: getPlaywrightScreenshotPath("error"),
      fullPage: true,
    });
    throw error;
  }
}

export async function validateQueryOrg(
  page: Page,
  options?: { displayName?: string }
) {
  try {
    console.log("start to verify query org");
    await page.waitForTimeout(Timeout.shortTimeLoading);
    const frameElementHandle = await page.waitForSelector(
      "iframe.embedded-page-content"
    );
    const frame = await frameElementHandle?.contentFrame();
    try {
      console.log("dismiss message");
      await frame?.waitForSelector("div.ui-box");
      await page
        .click('button:has-text("Dismiss")', {
          timeout: Timeout.playwrightDefaultTimeout,
        })
        .catch(() => {});
    } catch (error) {
      console.log("no message to dismiss");
    }
    const inputBar = await frame?.waitForSelector(
      "div.ui-popup__content input.ui-box"
    );
    await inputBar?.fill(options?.displayName || "");
    await page.waitForTimeout(Timeout.shortTimeLoading);
    const loginBtn = await frame?.waitForSelector(
      'div.ui-popup__content a:has-text("sign in")'
    );
    // todo add more verify
    // await RetryHandler.retry(async () => {
    //   console.log("Before popup");
    //   const [popup] = await Promise.all([
    //     page
    //       .waitForEvent("popup")
    //       .then((popup) =>
    //         popup
    //           .waitForEvent("close", {
    //             timeout: Timeout.playwrightConsentPopupPage,
    //           })
    //           .catch(() => popup)
    //       )
    //       .catch(() => {}),
    //     loginBtn?.click(),
    //   ]);
    //   console.log("after popup");

    //   if (popup && !popup?.isClosed()) {
    //     await popup.click('span:has-text("Continue")')
    //     await popup.click("input.button[type='submit'][value='Accept']");
    //   }
    // });
    // console.log("search ", displayName);
    // const input = await frame?.waitForSelector("div.ui-box input.ui-box");
    // await input?.type(displayName);

    console.log("verify query org successfully!!!");
  } catch (error) {
    await page.screenshot({
      path: getPlaywrightScreenshotPath("error"),
      fullPage: true,
    });
    throw error;
  }
}

export async function validateShareNow(page: Page) {
  try {
    console.log("start to verify share now");
    await page.waitForTimeout(Timeout.shortTimeLoading);
    const frameElementHandle = await page.waitForSelector(
      "iframe.embedded-page-content"
    );
    const frame = await frameElementHandle?.contentFrame();
    try {
      console.log("dismiss message");
      await frame?.waitForSelector("div.ui-box");
      await page
        .click('button:has-text("Dismiss")', {
          timeout: Timeout.playwrightDefaultTimeout,
        })
        .catch(() => {});
    } catch (error) {
      console.log("no message to dismiss");
    }

    await page.waitForTimeout(Timeout.shortTimeLoading);
    // click Suggest content
    console.log("click Suggest content");
    const startBtn = await frame?.waitForSelector(
      'span:has-text("Suggest content")'
    );
    await startBtn?.click();
    await page.waitForTimeout(Timeout.shortTimeLoading);
    // select content type
    console.log("select content type");
    const popupModal = await frame?.waitForSelector(
      ".ui-dialog .dialog-provider-wrapper"
    );
    const typeSelector = await popupModal?.waitForSelector(
      'button:has-text("Select content type")'
    );
    await typeSelector?.click();
    await page.waitForTimeout(Timeout.shortTimeLoading);
    const item = await popupModal?.waitForSelector(
      'ul li:has-text("Article / blog")'
    );
    await item?.click();
    await page.waitForTimeout(Timeout.shortTimeLoading);
    // fill in title
    console.log("fill in title");
    const titleInput = await popupModal?.waitForSelector(
      'input[placeholder="Type a title for your post"]'
    );
    await titleInput?.fill("test title");
    await page.waitForTimeout(Timeout.shortTimeLoading);
    // fill in description
    console.log("fill in description");
    const descriptionInput = await popupModal?.waitForSelector(
      'textarea[placeholder="Describe why you\'re suggesting this content"]'
    );
    await descriptionInput?.fill("test description for content suggestion");
    await page.waitForTimeout(Timeout.shortTimeLoading);
    // fill in link
    console.log("fill in link");
    const linkInput = await popupModal?.waitForSelector(
      'input[placeholder="Type or paste a link"]'
    );
    await linkInput?.fill("https://www.test.com");
    await page.waitForTimeout(Timeout.shortTimeLoading);
    // submit
    const submitBtn = await frame?.waitForSelector('span:has-text("Submit")');
    console.log("submit");
    await submitBtn?.click();
    await page.waitForTimeout(Timeout.shortTimeLoading);
    // verify
    await frame?.waitForSelector('span:has-text("test title")');
  } catch (error) {
    await page.screenshot({
      path: getPlaywrightScreenshotPath("error"),
      fullPage: true,
    });
    throw error;
  }
}

export async function validateWorkFlowBot(page: Page) {
  try {
    const frameElementHandle = await page.waitForSelector(
      "iframe.embedded-page-content"
    );
    const frame = await frameElementHandle?.contentFrame();
    await frame
      ?.click('button:has-text("DoStuff")', {
        timeout: Timeout.playwrightDefaultTimeout,
      })
      .catch(() => {});
    await frame?.waitForSelector(`p:has-text("[ACK] Hello World Bot")`);
  } catch (error) {
    await page.screenshot({
      path: getPlaywrightScreenshotPath("error"),
      fullPage: true,
    });
    throw error;
  }
}

export async function validateNotificationBot(
  page: Page,
  notificationEndpoint = "http://127.0.0.1:3978/api/notification"
) {
  try {
    const frameElementHandle = await page.waitForSelector(
      "iframe.embedded-page-content"
    );
    const frame = await frameElementHandle?.contentFrame();
    await frame?.waitForSelector("div.ui-box");
    await page
      .click('button:has-text("Dismiss")', {
        timeout: Timeout.playwrightDefaultTimeout,
      })
      .catch(() => {});
    await RetryHandler.retry(async () => {
      try {
        const result = await axios.post(notificationEndpoint);
        if (result.status !== 200) {
          throw new Error(
            `POST /api/notification failed: status code: '${result.status}', body: '${result.data}'`
          );
        }
        console.log("Successfully sent notification");
      } catch (e: any) {
        console.log(
          `[Command "welcome" not executed successfully] ${e.message}`
        );
      }
      await frame?.waitForSelector(
        'p:has-text("This is a sample http-triggered notification to Person")'
      );
    }, 2);
    console.log("User received notification");
  } catch (error) {
    await page.screenshot({
      path: getPlaywrightScreenshotPath("error"),
      fullPage: true,
    });
    throw error;
  }
}

export async function validateStockUpdate(page: Page) {
  try {
    console.log("start to verify stock update");
    await page.waitForTimeout(Timeout.shortTimeLoading);
    const frameElementHandle = await page.waitForSelector(
      "iframe.embedded-page-content"
    );
    const frame = await frameElementHandle?.contentFrame();
    try {
      console.log("click stock update");
      await frame?.waitForSelector('p:has-text("Microsoft Corporation")');
      console.log("verify stock update successfully!!!");
      await page.waitForTimeout(Timeout.shortTimeLoading);
    } catch (e: any) {
      console.log(`[Command not executed successfully] ${e.message}`);
      await page.screenshot({
        path: getPlaywrightScreenshotPath("error"),
        fullPage: true,
      });
      throw e;
    }
  } catch (error) {
    await page.screenshot({
      path: getPlaywrightScreenshotPath("error"),
      fullPage: true,
    });
    throw error;
  }
}

export async function validateTodoList(
  page: Page,
  options?: { displayName?: string }
) {
  try {
    console.log("start to verify todo list");
    try {
      const tabs = await page.$$("button[role='tab']");
      const tab = tabs.find(async (tab) => {
        const text = await tab.innerText();
        return text?.includes("Todo List");
      });
      await tab?.click();
      await page.waitForTimeout(Timeout.shortTimeLoading);
      const frameElementHandle = await page.waitForSelector(
        "iframe.embedded-iframe"
      );
      const frame = await frameElementHandle?.contentFrame();
      const startBtn = await frame?.waitForSelector('button:has-text("Start")');
      console.log("click Start button");
      await RetryHandler.retry(async () => {
        console.log("Before popup");
        const [popup] = await Promise.all([
          page
            .waitForEvent("popup")
            .then((popup) =>
              popup
                .waitForEvent("close", {
                  timeout: Timeout.playwrightConsentPopupPage,
                })
                .catch(() => popup)
            )
            .catch(() => {}),
          startBtn?.click(),
        ]);
        console.log("after popup");

        if (popup && !popup?.isClosed()) {
          await popup
            .click('button:has-text("Reload")', {
              timeout: Timeout.playwrightConsentPageReload,
            })
            .catch(() => {});
          await popup.click("input.button[type='submit'][value='Accept']");
        }
        const addBtn = await frame?.waitForSelector(
          'button:has-text("Add task")'
        );
        await addBtn?.click();
        //TODO: verify add task

        // clean tab, right click
        await tab?.click({ button: "right" });
        await page.waitForTimeout(Timeout.shortTimeLoading);
        const contextMenu = await page.waitForSelector("ul[role='menu']");
      });
    } catch (e: any) {
      console.log(`[Command not executed successfully] ${e.message}`);
      await page.screenshot({
        path: getPlaywrightScreenshotPath("error"),
        fullPage: true,
      });
      throw e;
    }

    await page.waitForTimeout(Timeout.shortTimeLoading);
  } catch (error) {
    await page.screenshot({
      path: getPlaywrightScreenshotPath("error"),
      fullPage: true,
    });
    throw error;
  }
}

export async function validateProactiveMessaging(page: Page): Promise<void> {
  console.log(`validating proactive messaging`);
  await page.waitForTimeout(Timeout.shortTimeLoading);
  const frameElementHandle = await page.waitForSelector(
    "iframe.embedded-page-content"
  );
  const frame = await frameElementHandle?.contentFrame();
  try {
    console.log("dismiss message");
    await frame?.waitForSelector("div.ui-box");
    await page
      .click('button:has-text("Dismiss")', {
        timeout: Timeout.playwrightDefaultTimeout,
      })
      .catch(() => {});
  } catch (error) {
    console.log("no message to dismiss");
  }
  try {
    console.log("sending message ", "welcome");
    await executeBotSuggestionCommand(page, frame, "welcome");
    await frame?.click('button[name="send"]');
  } catch (e: any) {
    await page.screenshot({
      path: getPlaywrightScreenshotPath("error"),
      fullPage: true,
    });
    console.log(`[Command 'learn' not executed successfully] ${e.message}`);
    throw e;
  }
}

async function executeBotSuggestionCommand(
  page: Page,
  frame: null | Frame,
  command: string
) {
  try {
    await frame?.click(`div.ui-list__itemheader:has-text("${command}")`);
  } catch (e: any) {
    console.log("can't find quickly select, try another way");
    await page.click('div[role="presentation"]:has-text("Chat")');
    console.log("open quick select");
    await page.click('div[role="presentation"]:has-text("Chat")');
    await frame?.click('div.cke_textarea_inline[role="textbox"]');
    console.log("select: ", command);
    await frame?.click(`div.ui-list__itemheader:has-text("${command}")`);
  }
}

export async function validateTeamsWorkbench(page: Page, displayName: string) {
  try {
    console.log("Load debug scripts");
    const frameElementHandle = await page.waitForSelector(
      "iframe.embedded-iframe"
    );
    const frame = await frameElementHandle?.contentFrame();
    await frame?.click('button:has-text("Load debug scripts")');
    console.log("Debug scripts loaded");
  } catch (error) {
    await page.screenshot({
      path: getPlaywrightScreenshotPath("error"),
      fullPage: true,
    });
    throw error;
  }
}

export async function validateSpfx(
  page: Page,
  options?: { displayName?: string }
) {
  try {
    const frameElementHandle = await page.waitForSelector(
      "iframe.embedded-page-content"
    );
    const frame = await frameElementHandle?.contentFrame();
    await frame?.waitForSelector(`text=${options?.displayName}`);
  } catch (error) {
    await page.screenshot({
      path: getPlaywrightScreenshotPath("error"),
      fullPage: true,
    });
    throw error;
  }
}

export async function switchToTab(page: Page) {
  try {
    await page.click('a:has-text("Personal Tab")');
  } catch (error) {
    await page.screenshot({
      path: getPlaywrightScreenshotPath("error"),
      fullPage: true,
    });
    throw error;
  }
}

export async function validateContact(
  page: Page,
  options?: { displayName?: string }
) {
  try {
    console.log("start to verify contact");
    await page.waitForTimeout(Timeout.shortTimeLoading);
    const frameElementHandle = await page.waitForSelector(
      "iframe.embedded-page-content"
    );
    const frame = await frameElementHandle?.contentFrame();
    try {
      console.log("dismiss message");
      await page
        .click('button:has-text("Dismiss")', {
          timeout: Timeout.playwrightDefaultTimeout,
        })
        .catch(() => {});
    } catch (error) {
      console.log("no message to dismiss");
    }
    try {
      const startBtn = await frame?.waitForSelector('button:has-text("Start")');
      await RetryHandler.retry(async () => {
        console.log("Before popup");
        const [popup] = await Promise.all([
          page
            .waitForEvent("popup")
            .then((popup) =>
              popup
                .waitForEvent("close", {
                  timeout: Timeout.playwrightConsentPopupPage,
                })
                .catch(() => popup)
            )
            .catch(() => {}),
          startBtn?.click(),
        ]);
        console.log("after popup");

        if (popup && !popup?.isClosed()) {
          await popup
            .click('button:has-text("Reload")', {
              timeout: Timeout.playwrightConsentPageReload,
            })
            .catch(() => {});
          await popup.click("input.button[type='submit'][value='Accept']");
        }

        await frame?.waitForSelector(`div:has-text("${options?.displayName}")`);
      });
      page.waitForTimeout(1000);

      // verify add person
      await addPerson(frame, options?.displayName || "");
      // verify delete person
      await delPerson(frame, options?.displayName || "");
    } catch (e: any) {
      console.log(`[Command not executed successfully] ${e.message}`);
      await page.screenshot({
        path: getPlaywrightScreenshotPath("error"),
        fullPage: true,
      });
      throw e;
    }

    await RetryHandler.retry(async () => {}, 2);

    await page.waitForTimeout(Timeout.shortTimeLoading);
  } catch (error) {
    await page.screenshot({
      path: getPlaywrightScreenshotPath("error"),
      fullPage: true,
    });
    throw error;
  }
}

export async function validateGraphConnector(
  page: Page,
  options?: { displayName?: string }
) {
  try {
    console.log("start to verify contact");
    await page.waitForTimeout(Timeout.shortTimeLoading);
    const frameElementHandle = await page.waitForSelector(
      "iframe.embedded-page-content"
    );
    const frame = await frameElementHandle?.contentFrame();
    try {
      const startBtn = await frame?.waitForSelector('button:has-text("Start")');
      await RetryHandler.retry(async () => {
        console.log("Before popup");
        const [popup] = await Promise.all([
          page
            .waitForEvent("popup")
            .then((popup) =>
              popup
                .waitForEvent("close", {
                  timeout: Timeout.playwrightConsentPopupPage,
                })
                .catch(() => popup)
            )
            .catch(() => {}),
          startBtn?.click(),
        ]);
        console.log("after popup");

        if (popup && !popup?.isClosed()) {
          await popup
            .click('button:has-text("Reload")', {
              timeout: Timeout.playwrightConsentPageReload,
            })
            .catch(() => {});
          await popup.click("input.button[type='submit'][value='Accept']");
        }

        await frame?.waitForSelector(`div:has-text("${options?.displayName}")`);
      });
      page.waitForTimeout(1000);
    } catch (e: any) {
      console.log(`[Command not executed successfully] ${e.message}`);
    }

    await page.waitForTimeout(Timeout.shortTimeLoading);
  } catch (error) {
    await page.screenshot({
      path: getPlaywrightScreenshotPath("error"),
      fullPage: true,
    });
    throw error;
  }
}

export async function validateMsg(page: Page) {
  try {
    const frameElementHandle = await page.waitForSelector(
      "iframe.embedded-page-content"
    );
    const frame = await frameElementHandle?.contentFrame();
    await frame?.waitForSelector("div.ui-box");
    console.log("start to validate msg");
    try {
      await frame?.waitForSelector('input[aria-label="Your search query"]');
    } catch (error) {
      console.log("no search box to validate msg.");
      return;
    }
    //input keyword
    const searchKeyword = "teamsfx";
    //check
    await frame?.fill('input[aria-label="Your search query"]', searchKeyword);
    console.log("Check if npm list showed");
    await frame?.waitForSelector('ul[datatid="app-picker-list"]');
    console.log("[search for npm packages success]");
  } catch (error) {
    console.log("[search for npm packages faild,Unable to reach app]");
    await page.screenshot({
      path: getPlaywrightScreenshotPath("error"),
      fullPage: true,
    });
    throw error;
  }
}

export async function validateBasicDashboardTab(page: Page) {
  try {
    console.log("start to verify dashboard tab");
    const frameElementHandle = await page.waitForSelector(
      "iframe.embedded-iframe"
    );
    const frame = await frameElementHandle?.contentFrame();
    await frame?.waitForSelector("span:has-text('Your List')");
    await frame?.waitForSelector("span:has-text('Your chart')");
    await frame?.waitForSelector("button:has-text('View Details')");
    console.log("Dashboard tab loaded successfully");
  } catch (error) {
    await page.screenshot({
      path: getPlaywrightScreenshotPath("error"),
      fullPage: true,
    });
    throw error;
  }
}

export async function validateDashboardTab(page: Page) {
  try {
    console.log("start to verify dashboard tab");
    const frameElementHandle = await page.waitForSelector(
      "iframe.embedded-iframe"
    );
    const frame = await frameElementHandle?.contentFrame();
    await frame?.waitForSelector("span:has-text('Area chart')");
    await frame?.waitForSelector("span:has-text('Your upcoming events')");
    await frame?.waitForSelector("span:has-text('Your tasks')");
    await frame?.waitForSelector("span:has-text('Team collaborations')");
    await frame?.waitForSelector("span:has-text('Your documents')");
    console.log("Dashboard tab loaded successfully");
  } catch (error) {
    await page.screenshot({
      path: getPlaywrightScreenshotPath("error"),
      fullPage: true,
    });
    throw error;
  }
}

export async function validateNotificationTimeBot(page: Page) {
  try {
    const frameElementHandle = await page.waitForSelector(
      "iframe.embedded-page-content"
    );
    const frame = await frameElementHandle?.contentFrame();
    await frame?.waitForSelector("div.ui-box");
    await RetryHandler.retry(async () => {
      await frame?.waitForSelector(
        `p:has-text("This is a sample time-triggered notification")`
      );
      console.log("verify noti time-trigger bot successfully!!!");
    }, 2);
    console.log("User received notification");
  } catch (error) {
    await page.screenshot({
      path: getPlaywrightScreenshotPath("error"),
      fullPage: true,
    });
    throw error;
  }
}

export async function validateAdaptiveCard(
  page: Page,
  options?: { context?: SampledebugContext; env?: "local" | "dev" }
) {
  try {
    const frameElementHandle = await page.waitForSelector(
      "iframe.embedded-page-content"
    );
    const frame = await frameElementHandle?.contentFrame();
    await frame?.waitForSelector("div.ui-box");
    await page
      .click('button:has-text("Dismiss")', {
        timeout: Timeout.playwrightDefaultTimeout,
      })
      .catch(() => {});
    await RetryHandler.retry(async () => {
      try {
        // send post request to bot
        console.log("Post request sent to bot");
        let url: string;
        if (options?.env === "dev") {
          const endpointFilePath = path.join(
            options?.context?.projectPath ?? "",
            "env",
            ".env.dev"
          );
          // read env file
          const endpoint = fs.readFileSync(endpointFilePath, "utf8");
          const devEnv = dotenvUtil.deserialize(endpoint);
          url =
            devEnv.obj["BOT_FUNCTION_ENDPOINT"] + "/api/default-notification";
        } else {
          url = "http://127.0.0.1:3978/api/default-notification";
        }
        console.log(url);
        await axios.post(url);
        await frame?.waitForSelector('p:has-text("New Event Occurred!")');
        console.log("Successfully sent notification");
      } catch (e: any) {
        console.log(`[ Not receive response! ] ${e.message}`);
        await page.screenshot({
          path: getPlaywrightScreenshotPath("error"),
          fullPage: true,
        });
        throw e;
      }
    }, 2);
  } catch (error) {
    await page.screenshot({
      path: getPlaywrightScreenshotPath("error"),
      fullPage: true,
    });
    throw error;
  }
}

export async function addPerson(
  frame: Frame | null,
  displayName: string
): Promise<void> {
  console.log(`add person: ${displayName}`);
  const input = await frame?.waitForSelector("input#people-picker-input");
  await input?.click();
  await input?.type(displayName);
  const item = await frame?.waitForSelector(`span:has-text("${displayName}")`);
  await item?.click();
  await frame?.waitForSelector(
    `div.table-area div.line1:has-text("${displayName}")`
  );
}

export async function delPerson(
  frame: Frame | null,
  displayName: string
): Promise<void> {
  console.log(`delete person: ${displayName}`);
  await frame?.waitForSelector(
    `li div.details.small div:has-text("${displayName}")`
  );

  const closeBtn = await frame?.waitForSelector('li div[role="button"]');
  await closeBtn?.click();
  await frame?.waitForSelector(
    `div.table-area div.line1:has-text("${displayName}")`,
    { state: "detached" }
  );
}

const assistantDashboardMiddleWare = (
  sampledebugContext: SampledebugContext,
  env: "local" | "dev",
  azSqlHelper?: AzSqlHelper,
  steps?: {
    afterCreate?: boolean;
  }
) => {
  if (steps?.afterCreate) {
    const envFilePath = path.resolve(
      sampledebugContext.projectPath,
      "env",
      `.env.${env}.user`
    );
    const envString =
      'PLANNER_GROUP_ID=YOUR_PLANNER_GROUP_ID\nDEVOPS_ORGANIZATION_NAME=msazure\nDEVOPS_PROJECT_NAME="Microsoft Teams Extensibility"\nGITHUB_REPO_NAME=test002\nGITHUB_REPO_OWNER=hellyzh\nPLANNER_PLAN_ID=YOUR_PLAN_ID\nPLANNER_BUCKET_ID=YOUR_BUCKET_ID\nSECRET_DEVOPS_ACCESS_TOKEN=YOUR_DEVOPS_ACCESS_TOKEN\nSECRET_GITHUB_ACCESS_TOKEN=YOUR_GITHUB_ACCESS_TOKEN';
    fs.writeFileSync(envFilePath, envString);
  }
};

const incomingWebhookMiddleWare = async (
  sampledebugContext: SampledebugContext,
  env: "local" | "dev",
  azSqlHelper?: AzSqlHelper,
  steps?: {
    afterCreate?: boolean;
  }
) => {
  if (steps?.afterCreate) {
    // replace "<webhook-url>" to "https://test.com"
    console.log("replace webhook url");
    const targetFile = path.resolve(
      sampledebugContext.projectPath,
      "src",
      "index.ts"
    );
    let data = fs.readFileSync(targetFile, "utf-8");
    data = data.replace(/<webhook-url>/g, "https://test.com");
    fs.writeFileSync(targetFile, data);
    console.log("replace webhook url finish!");
  }
};

const shareNowMiddleWare = async (
  sampledebugContext: SampledebugContext,
  env: "local" | "dev",
  azSqlHelper?: AzSqlHelper,
  steps?: {
    before?: boolean;
    afterCreate?: boolean;
    afterDeploy?: boolean;
  }
) => {
  const sqlUserName = "Abc123321";
  const sqlPassword = "Cab232332" + uuid.v4().substring(0, 6);
  if (steps?.afterCreate) {
    if (env === "dev") {
      const envFilePath = path.resolve(
        sampledebugContext.projectPath,
        "env",
        ".env.dev.user"
      );
      editDotEnvFile(envFilePath, "SQL_USER_NAME", sqlUserName);
      editDotEnvFile(envFilePath, "SQL_PASSWORD", sqlPassword);
    } else {
      const res = await azSqlHelper?.createSql();
      expect(res).to.be.true;
      const envFilePath = path.resolve(
        sampledebugContext.projectPath,
        "env",
        ".env.local.user"
      );
      editDotEnvFile(envFilePath, "SQL_USER_NAME", azSqlHelper?.sqlAdmin ?? "");
      editDotEnvFile(
        envFilePath,
        "SQL_PASSWORD",
        azSqlHelper?.sqlPassword ?? ""
      );
      editDotEnvFile(
        envFilePath,
        "SQL_ENDPOINT",
        azSqlHelper?.sqlEndpoint ?? ""
      );
      editDotEnvFile(
        envFilePath,
        "SQL_DATABASE_NAME",
        azSqlHelper?.sqlDatabaseName ?? ""
      );
    }
  }
  if (steps?.afterDeploy) {
    if (env === "local") return;
    const devEnvFilePath = path.resolve(
      sampledebugContext.projectPath,
      "env",
      ".env.dev"
    );
    // read database from devEnvFilePath
    const sqlDatabaseNameLine = fs
      .readFileSync(devEnvFilePath, "utf-8")
      .split("\n")
      .find((line: string) =>
        line.startsWith("PROVISIONOUTPUT__AZURESQLOUTPUT__DATABASENAME")
      );

    const sqlDatabaseName = sqlDatabaseNameLine
      ? sqlDatabaseNameLine.split("=")[1]
      : undefined;

    const sqlEndpointLine = fs
      .readFileSync(devEnvFilePath, "utf-8")
      .split("\n")
      .find((line: string) =>
        line.startsWith("PROVISIONOUTPUT__AZURESQLOUTPUT__SQLENDPOINT")
      );

    const sqlEndpoint = sqlEndpointLine
      ? sqlEndpointLine.split("=")[1]
      : undefined;

    const sqlCommands = [
      `CREATE TABLE [TeamPostEntity](
          [PostID] [int] PRIMARY KEY IDENTITY,
          [ContentUrl] [nvarchar](400) NOT NULL,
          [CreatedByName] [nvarchar](50) NOT NULL,
          [CreatedDate] [datetime] NOT NULL,
          [Description] [nvarchar](500) NOT NULL,
          [IsRemoved] [bit] NOT NULL,
          [Tags] [nvarchar](100) NULL,
          [Title] [nvarchar](100) NOT NULL,
          [TotalVotes] [int] NOT NULL,
          [Type] [int] NOT NULL,
          [UpdatedDate] [datetime] NOT NULL,
          [UserID] [uniqueidentifier] NOT NULL
       );`,
      `CREATE TABLE [UserVoteEntity](
        [VoteID] [int] PRIMARY KEY IDENTITY,
        [PostID] [int] NOT NULL,
        [UserID] [uniqueidentifier] NOT NULL
      );`,
    ];
    const sqlHelper = new AzSqlHelper(
      `${sampledebugContext.appName}-dev-rg`,
      sqlCommands,
      sqlDatabaseName,
      sqlDatabaseName,
      sqlUserName,
      sqlPassword
    );
    await sqlHelper.createTable(sqlEndpoint ?? "");
  }
};

const stockUpdateMiddleWare = async (
  sampledebugContext: SampledebugContext,
  env: "local" | "dev",
  azSqlHelper?: AzSqlHelper,
  steps?: {
    afterCreate?: boolean;
  }
) => {
  if (steps?.afterCreate) {
    const envFile = path.resolve(
      sampledebugContext.projectPath,
      "env",
      `.env.${env}`
    );
    let ENDPOINT = fs.readFileSync(envFile, "utf-8");
    ENDPOINT +=
      "\nTEAMSFX_API_ALPHAVANTAGE_ENDPOINT=https://www.alphavantage.co";
    fs.writeFileSync(envFile, ENDPOINT);
    console.log(`add endpoint ${ENDPOINT} to .env.${env} file`);
    const userFile = path.resolve(
      sampledebugContext.projectPath,
      "env",
      `.env.${env}.user`
    );
    let KEY = fs.readFileSync(userFile, "utf-8");
    KEY += "\nTEAMSFX_API_ALPHAVANTAGE_API_KEY=demo";
    fs.writeFileSync(userFile, KEY);
    console.log(`add key ${KEY} to .env.${env}.user file`);
  }
};

const todoListSqlMiddleWare = async (
  sampledebugContext: SampledebugContext,
  env: "local" | "dev",
  azSqlHelper?: AzSqlHelper,
  steps?: {
    before?: boolean;
    afterCreate?: boolean;
    afterdeploy?: boolean;
  }
) => {
  const sqlUserName = "Abc123321";
  const sqlPassword = "Cab232332" + uuid.v4().substring(0, 6);
  if (steps?.before) {
    // create sql db server
    const rgName = `${sampledebugContext.appName}-dev-rg`;
    const sqlCommands = [
      `CREATE TABLE Todo
         (
             id INT IDENTITY PRIMARY KEY,
             description NVARCHAR(128) NOT NULL,
             objectId NVARCHAR(36),
             channelOrChatId NVARCHAR(128),
             isCompleted TinyInt NOT NULL default 0,
         )`,
    ];
    azSqlHelper = new AzSqlHelper(rgName, sqlCommands);
    return azSqlHelper;
  }
  if (steps?.afterCreate) {
    const envFilePath = path.resolve(
      sampledebugContext.projectPath,
      "env",
      `.env.${env}.user`
    );
    if (env === "dev") {
      editDotEnvFile(envFilePath, "SQL_USER_NAME", sqlUserName);
      editDotEnvFile(envFilePath, "SQL_PASSWORD", sqlPassword);
    } else {
      const res = await azSqlHelper?.createSql();
      expect(res).to.be.true;
      editDotEnvFile(envFilePath, "SQL_USER_NAME", azSqlHelper?.sqlAdmin ?? "");
      editDotEnvFile(
        envFilePath,
        "SQL_PASSWORD",
        azSqlHelper?.sqlPassword ?? ""
      );
      editDotEnvFile(
        envFilePath,
        "SQL_ENDPOINT",
        azSqlHelper?.sqlEndpoint ?? ""
      );
      editDotEnvFile(
        envFilePath,
        "SQL_DATABASE_NAME",
        azSqlHelper?.sqlDatabaseName ?? ""
      );
    }
  }
  if (steps?.afterdeploy) {
    if (env === "local") return;
    // read database from devEnvFilePath
    const devEnvFilePath = path.resolve(
      sampledebugContext.projectPath,
      "env",
      ".env.dev"
    );
    const sqlDatabaseName = fs
      .readFileSync(devEnvFilePath, "utf-8")
      .split("\n")
      .find((line) =>
        line.startsWith("PROVISIONOUTPUT__AZURESQLOUTPUT__DATABASENAME")
      )
      ?.split("=")[1];
    const sqlEndpoint = fs
      .readFileSync(devEnvFilePath, "utf-8")
      .split("\n")
      .find((line) =>
        line.startsWith("PROVISIONOUTPUT__AZURESQLOUTPUT__SQLENDPOINT")
      )
      ?.split("=")[1];

    const sqlCommands = [
      `CREATE TABLE Todo
            (
                id INT IDENTITY PRIMARY KEY,
                description NVARCHAR(128) NOT NULL,
                objectId NVARCHAR(36),
                channelOrChatId NVARCHAR(128),
                isCompleted TinyInt NOT NULL default 0,
            )`,
    ];
    const sqlHelper = new AzSqlHelper(
      `${sampledebugContext.appName}-dev-rg`,
      sqlCommands,
      sqlDatabaseName,
      sqlDatabaseName,
      sqlUserName,
      sqlPassword
    );
    await sqlHelper.createTable(sqlEndpoint ?? "");
  }
};

const TodoListM365MiddleWare = async (
  sampledebugContext: SampledebugContext,
  env: "local" | "dev",
  azSqlHelper?: AzSqlHelper,
  steps?: {
    afterCreate?: boolean;
  }
) => {
  if (steps?.afterCreate) {
    if (env === "dev") {
      const envFilePath = path.resolve(
        sampledebugContext.projectPath,
        "env",
        ".env.dev.user"
      );
      editDotEnvFile(envFilePath, "SQL_USER_NAME", "Abc123321");
      editDotEnvFile(
        envFilePath,
        "SQL_PASSWORD",
        "Cab232332" + uuid.v4().substring(0, 6)
      );
    }
    const targetPath = path.resolve(sampledebugContext.projectPath, "tabs");
    const data = "src/";
    // create .eslintignore
    fs.writeFileSync(targetPath + "/.eslintignore", data);
  }
};
