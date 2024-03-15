// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
import { BrowserContext, Page, Frame } from "playwright";
import { assert, expect } from "chai";
import { Timeout, ValidationContent, TemplateProject } from "./constants";
import { RetryHandler } from "./retryHandler";
import { getPlaywrightScreenshotPath } from "./nameUtil";
import axios from "axios";
import { SampledebugContext } from "../ui-test/samples/sampledebugContext";
import path from "path";
import fs from "fs";
import { dotenvUtil } from "./envUtil";
import { startDebugging, startDebuggingAzure } from "./vscodeOperation";

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
    await startDebugging("Debug (Chrome)");
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
  [TemplateProject.GraphConnectorBot]: async () => {
    await startDebugging();
  },
  [TemplateProject.SpfxProductivity]: async () => {
    await startDebugging("Teams workbench (Chrome)");
  },
  [TemplateProject.RetailDashboard]: async () => {
    await startDebugging("Teams workbench (Chrome)");
  },
  [TemplateProject.TabSSOApimProxy]: async () => {
    await startDebuggingAzure(
      "Debug in Teams (Chrome)",
      "local",
      `TabSSOApimProxy`
    );
  },
  [TemplateProject.LargeScaleBot]: async () => {
    await startDebugging();
  },
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
    const addBtn = await frame?.waitForSelector("button>span:has-text('Add')");

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
        try {
          await popup?.close();
        } catch (error) {
          console.log("popup is closed");
        }
      }
    } else {
      await addBtn?.click();
    }
    await page.waitForTimeout(Timeout.shortTimeLoading);
    // verify add page is closed
    await frame?.waitForSelector("button>span:has-text('Add')", {
      state: "detached",
    });
    try {
      try {
        await page?.waitForSelector(".team-information span:has-text('About')");
      } catch (error) {
        try {
          await page?.waitForSelector(
            ".ts-messages-header span:has-text('About')"
          );
        } catch (error) {
          try {
            await page?.waitForSelector(
              ".team-information span:has-text('Chat')"
            );
          } catch (error) {
            await page?.waitForSelector(
              ".ts-messages-header span:has-text('Chat')"
            );
          }
        }
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

export async function reopenPage(
  context: BrowserContext,
  teamsAppId: string,
  username: string,
  password: string,
  options?: {
    teamsAppName?: string;
    dashboardFlag?: boolean;
  },
  addApp = true,
  inputPassword = false
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

  if (inputPassword) {
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
  }

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

    await page.screenshot({
      path: getPlaywrightScreenshotPath("reopen_page"),
      fullPage: true,
    });
    const frameElementHandle = await page.waitForSelector(
      "iframe.embedded-page-content"
    );
    const frame = await frameElementHandle?.contentFrame();
    await page.waitForTimeout(Timeout.shortTimeLoading);
    if (addApp) {
      console.log("click add button");
      const addBtn = await frame?.waitForSelector(
        "button>span:has-text('Add')"
      );

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
          try {
            await popup?.close();
          } catch (error) {
            console.log("popup is closed");
          }
        }
      } else {
        await addBtn?.click();
      }
      await page.waitForTimeout(Timeout.shortTimeLoading);
      // verify add page is closed
      await frame?.waitForSelector("button>span:has-text('Add')", {
        state: "detached",
      });
    }
    try {
      try {
        await page?.waitForSelector(".team-information span:has-text('About')");
      } catch (error) {
        try {
          await page?.waitForSelector(
            ".ts-messages-header span:has-text('About')"
          );
        } catch (error) {
          try {
            await page?.waitForSelector(
              ".team-information span:has-text('Chat')"
            );
          } catch (error) {
            await page?.waitForSelector(
              ".ts-messages-header span:has-text('Chat')"
            );
          }
        }
      }
      console.log("[success] app loaded");
    } catch (error) {
      await page.screenshot({
        path: getPlaywrightScreenshotPath("add_error"),
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
        "button>span:has-text('Add')"
      );
      await addBtn?.click();
      await page.waitForTimeout(Timeout.shortTimeLoading);

      if (options?.type === "meeting") {
        // verify add page is closed
        const frameElementHandle = await page.waitForSelector(
          "iframe.embedded-page-content"
        );
        const frame = await frameElementHandle?.contentFrame();
        try {
          await frame?.waitForSelector(
            `h1:has-text('Add ${options?.teamsAppName} to a team')`
          );
        } catch (error) {
          await frame?.waitForSelector(
            `h1:has-text('Add ${options?.teamsAppName} to a meeting')`
          );
        }
        // TODO: need to add more logic
        console.log("successful to add teams app!!!");
        return;
      }

      try {
        // verify add page is closed
        await frame?.waitForSelector(`h1:has-text('to a team')`);
        try {
          const frameElementHandle = await page.waitForSelector(
            "iframe.embedded-page-content"
          );
          const frame = await frameElementHandle?.contentFrame();

          try {
            const items = await frame?.waitForSelector("li.ui-dropdown__item");
            await items?.click();
            console.log("selected a team.");
          } catch (error) {
            const searchBtn = await frame?.waitForSelector(
              "div.ui-dropdown__toggle-indicator"
            );
            await searchBtn?.click();
            await page.waitForTimeout(Timeout.shortTimeLoading);

            const items = await frame?.waitForSelector("li.ui-dropdown__item");
            await items?.click();
            console.log("[catch] selected a team.");
          }

          const setUpBtn = await frame?.waitForSelector(
            'button span:has-text("Set up a tab")'
          );
          await setUpBtn?.click();
          console.log("click 'set up a tab' button");
          await page.waitForTimeout(Timeout.shortTimeLoading);
          await frame?.waitForSelector('button span:has-text("Set up a tab")', {
            state: "detached",
          });
        } catch (error) {
          console.log(error);
          await page.screenshot({
            path: getPlaywrightScreenshotPath("error"),
            fullPage: true,
          });
          throw error;
        }
      } catch (error) {
        console.log("no need to add to a team step");
      }

      {
        console.log('[start] click "save" button');
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
          console.log('[success] click "save" button');
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

export async function reopenTeamsPage(
  context: BrowserContext,
  teamsAppId: string,
  username: string,
  password: string,
  options?: {
    teamsAppName?: string;
    dashboardFlag?: boolean;
    type?: string;
  },
  addApp = true
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
      if (addApp) {
        await page.waitForTimeout(Timeout.longTimeWait);
        console.log("click add button");
        // default
        const addBtn = await frame?.waitForSelector(
          "button>span:has-text('Add')"
        );
        await addBtn?.click();
      }
      await page.waitForTimeout(Timeout.shortTimeLoading);

      if (options?.type === "meeting") {
        // verify add page is closed
        const frameElementHandle = await page.waitForSelector(
          "iframe.embedded-page-content"
        );
        const frame = await frameElementHandle?.contentFrame();
        try {
          await frame?.waitForSelector(
            `h1:has-text('Add ${options?.teamsAppName} to a team')`
          );
        } catch (error) {
          await frame?.waitForSelector(
            `h1:has-text('Add ${options?.teamsAppName} to a meeting')`
          );
        }
        // TODO: need to add more logic
        console.log("successful to add teams app!!!");
        return;
      }

      try {
        // verify add page is closed
        await frame?.waitForSelector(`h1:has-text('to a team')`);
        try {
          const frameElementHandle = await page.waitForSelector(
            "iframe.embedded-page-content"
          );
          const frame = await frameElementHandle?.contentFrame();

          try {
            const items = await frame?.waitForSelector("li.ui-dropdown__item");
            await items?.click();
            console.log("selected a team.");
          } catch (error) {
            const searchBtn = await frame?.waitForSelector(
              "div.ui-dropdown__toggle-indicator"
            );
            await searchBtn?.click();
            await page.waitForTimeout(Timeout.shortTimeLoading);

            const items = await frame?.waitForSelector("li.ui-dropdown__item");
            await items?.click();
            console.log("[catch] selected a team.");
          }

          const setUpBtn = await frame?.waitForSelector(
            'button span:has-text("Set up a tab")'
          );
          await setUpBtn?.click();
          console.log("click 'set up a tab' button");
          await page.waitForTimeout(Timeout.shortTimeLoading);
          await frame?.waitForSelector('button span:has-text("Set up a tab")', {
            state: "detached",
          });
        } catch (error) {
          console.log(error);
          await page.screenshot({
            path: getPlaywrightScreenshotPath("error"),
            fullPage: true,
          });
          throw error;
        }
      } catch (error) {
        console.log("no need to add to a team step");
      }

      {
        console.log('[start] click "save" button');
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
          console.log('[success] click "save" button');
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

export async function initNoAddappPage(
  context: BrowserContext,
  teamsAppId: string,
  username: string,
  password: string,
  options?: {
    teamsAppName?: string;
    dashboardFlag?: boolean;
  }
): Promise<Page> {
  const page = await context.newPage();
  page.setDefaultTimeout(Timeout.playwrightDefaultTimeout);
  // open teams app page
  // https://github.com/puppeteer/puppeteer/issues/3338
  await Promise.all([
    page.goto(
      // `https://teams.microsoft.com/_#/l/app/${teamsAppId}?installAppPackage=true`
      `https://teams.microsoft.com`
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
  const chatTab = await page?.waitForSelector(
    ".app-bar-items span:has-text('Chat')"
  );
  await chatTab?.click();
  try {
    console.log("close dialog");
    await page?.click("button[data-tid='closeModelDialogBtn']");
  } catch (error) {
    console.log("no dialog to close");
  }
  try {
    console.log("dismiss message");
    await page.click('button:has-text("Dismiss")');
  } catch (error) {
    console.log("no message to dismiss");
  }
  return page;
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

export async function validateEchoBot(
  page: Page,
  options: { botCommand?: string; expected?: ValidationContent } = {
    botCommand: "helloWorld",
    expected: ValidationContent.BotWelcomeInstruction,
  }
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

    await RetryHandler.retry(async () => {
      await frame?.waitForSelector(
        `p:has-text("${
          options?.expected || ValidationContent.BotWelcomeInstruction
        }")`
      );
      console.log(options?.expected || ValidationContent.BotWelcomeInstruction);
      console.log("verified bot that it has sent welcome!!!");
    }, 2);

    await RetryHandler.retry(async () => {
      console.log("sending message ", options?.botCommand);
      await frame?.fill(
        'div.ck-content[role="textbox"]',
        options?.botCommand || "helloWorld"
      );
      await frame?.click('button[name="send"]');
      const expectedContent = options?.botCommand
        ? `Echo: ${options?.botCommand}`
        : `Echo: helloWorld`;
      await frame?.waitForSelector(`p:has-text("${expectedContent}")`);
      console.log(`verify bot successfully with content ${expectedContent}!!!`);
    }, 2);
    await page.waitForTimeout(Timeout.shortTimeLoading);
  } catch (error) {
    await page.screenshot({
      path: getPlaywrightScreenshotPath("error"),
      fullPage: true,
    });
    throw error;
  }
}

export async function validateWelcomeAndReplyBot(
  page: Page,
  options: {
    hasWelcomeMessage?: boolean;
    hasCommandReplyValidation: boolean;
    botCommand?: string;
    expectedWelcomeMessage?: string;
    expectedReplyMessage?: string;
  } = {
    hasWelcomeMessage: true,
    hasCommandReplyValidation: true,
    botCommand: "helloWorld",
    expectedWelcomeMessage: ValidationContent.AiChatBotWelcomeInstruction,
    expectedReplyMessage: ValidationContent.AiBotErrorMessage,
  }
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

    if (options.hasWelcomeMessage) {
      await RetryHandler.retry(async () => {
        await frame?.waitForSelector(
          `p:has-text("${
            options?.expectedWelcomeMessage ||
            ValidationContent.AiChatBotWelcomeInstruction
          }")`
        );
        console.log(
          options?.expectedWelcomeMessage ||
            ValidationContent.AiChatBotWelcomeInstruction
        );
        console.log("verified bot that it has sent welcome!!!");
      }, 2);
    }

    if (options.hasCommandReplyValidation) {
      await RetryHandler.retry(async () => {
        console.log("sending message ", options?.botCommand || "helloWorld");
        await frame?.fill(
          'div.ck-content[role="textbox"]',
          options?.botCommand || "helloWorld"
        );
        await frame?.click('button[name="send"]');
        await frame?.waitForSelector(
          `p:has-text("${options?.expectedReplyMessage}")`
        );
        console.log(
          `verify bot successfully with content ${options?.expectedReplyMessage}!!!`
        );
      }, 2);
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

export async function validateBot(
  page: Page,
  options: { botCommand?: string; expected?: ValidationContent } = {
    botCommand: "welcome",
    expected: ValidationContent.Bot,
  }
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

    if (options?.botCommand === "show") {
      try {
        console.log("sending message ", options?.botCommand);
        await executeBotSuggestionCommand(page, frame, options?.botCommand);
        await frame?.click('button[name="send"]');
      } catch (e: any) {
        console.log(
          `[Command "${options?.botCommand}" not executed successfully] ${e.message}`
        );
      }
      await RetryHandler.retry(async () => {
        try {
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
            await popup
              .click('button:has-text("Continue")', {
                timeout: Timeout.playwrightConsentPageReload,
              })
              .catch(() => {});
            await popup.click("input.button[type='submit'][value='Accept']");
          }
        } catch (error) {
          console.log("reopen skip step");
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
        console.log("sending message ", options?.botCommand);
        await executeBotSuggestionCommand(
          page,
          frame,
          options?.botCommand || "welcome"
        );
        await frame?.click('button[name="send"]');
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
    const searchPack = options?.npmName || "axios";
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
    console.log("search npm ", searchPack);
    const input = await frame?.waitForSelector("div.ui-box input.ui-box");
    await input?.type(searchPack);
    try {
      const targetItem = await frame?.waitForSelector(
        `span:has-text("${searchPack}")`
      );
      await targetItem?.click();
      await frame?.waitForSelector(`card span:has-text("${searchPack}")`);
      await frame?.click('button[name="send"]');
      console.log("verify npm search successfully!!!");
      await page.waitForTimeout(Timeout.shortTimeLoading);
    } catch (error) {
      await frame?.waitForSelector(
        'div.ui-box span:has-text("Unable to reach app. Please try again.")'
      );
      await page.screenshot({
        path: getPlaywrightScreenshotPath("verify_error"),
        fullPage: true,
      });
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
      try {
        await frame?.waitForSelector(
          'p:has-text("This is a sample http-triggered notification to Person")'
        );
      } catch (e) {
        console.log("sending any message ", "helloWorld");
        await frame?.fill('div.ck-content[role="textbox"]', "helloWorld");
        await frame?.click('button[name="send"]');
        throw e;
      }
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
      await page.waitForTimeout(Timeout.shortTimeLoading);
      const frameElementHandle = await page.waitForSelector(
        "iframe.embedded-iframe"
      );
      const frame = await frameElementHandle?.contentFrame();
      const childFrame = frame?.childFrames()[0];
      const startBtn = await childFrame?.waitForSelector(
        'button:has-text("Start")'
      );
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
      });
      // add task
      console.log("click add task button");
      const addBtn = await childFrame?.waitForSelector(
        'button:has-text("Add task")'
      );
      await addBtn?.click();
      const inputBox = await childFrame?.waitForSelector(
        "div.item.add input[type='text']"
      );
      console.log("type hello world");
      await inputBox?.type("Hello World");
      await addBtn?.click();
      console.log("check result");
      await childFrame?.waitForSelector(
        `div.item .creator .name:has-text("${options?.displayName}")`
      );
      console.log("debug finish!!!");
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

export async function validateProactiveMessaging(
  page: Page,
  options?: { env: "local" | "dev"; context?: SampledebugContext }
): Promise<void> {
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
    // verify command
    const expectedContent = "You sent 'welcome '.";
    await frame?.waitForSelector(`p:has-text("${expectedContent}")`);
    console.log(`verify bot successfully with content ${expectedContent}!!!`);
    // send post request to bot
    console.log("Post request sent to bot");
    const endpointFilePath = path.join(
      options?.context?.projectPath ?? "",
      "env",
      `.env.${options?.env}`
    );
    // read env file
    const endpoint = fs.readFileSync(endpointFilePath, "utf8");
    const devEnv = dotenvUtil.deserialize(endpoint);
    const url =
      devEnv.obj["PROVISIONOUTPUT__BOTOUTPUT__SITEENDPOINT"] + "/api/notify";
    console.log(url);
    await axios.get(url);
    await frame?.waitForSelector('p:has-text("proactive hello")');
    console.log("Successfully sent notification");
    await page.waitForTimeout(Timeout.shortTimeLoading);
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
    await frame?.click(`div.autocompleteItem__header:has-text("${command}")`);
  } catch (e: any) {
    try {
      console.log("can't find quickly select, try another way");
      await frame?.click('div.ui-flex[role="main"]');
      console.log("open quick select");
      await frame?.click('div.ui-flex[role="main"]');
      await frame?.click('div.ck-content[role="textbox"]');
      console.log("select: ", command);
      await frame?.click(`div.autocompleteItem__header:has-text("${command}")`);
    } catch (e: any) {
      console.log(
        `[Command ${command} not executed successfully] ${e.message}`
      );
    }
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
    console.log(`Found: "${options?.displayName}"`);
  } catch (error) {
    await page.screenshot({
      path: getPlaywrightScreenshotPath("error"),
      fullPage: true,
    });
    throw error;
  }
}

export async function switchToTab(page: Page, tabName = "Personal Tab") {
  try {
    await page.click(`a:has-text("${tabName}")`);
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
  options?: { displayName?: string },
  rerun = false
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
      if (!rerun) {
        const startBtn = await frame?.waitForSelector(
          'button:has-text("Start")'
        );
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

          await frame?.waitForSelector(
            `div:has-text("${options?.displayName}")`
          );
        });
      }
      await page.waitForTimeout(10000);

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
  const input = await frame?.waitForSelector("input#control");
  await input?.click();
  await input?.type(displayName);
  const item = await frame?.waitForSelector(
    `ul#suggestions-list div:has-text("${displayName}")`
  );
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
    `li.selected-list-item div:has-text("${displayName}")`
  );

  const closeBtn = await frame?.waitForSelector('li div[role="button"]');
  await closeBtn?.click();
  await frame?.waitForSelector(
    `div.table-area div.line1:has-text("${displayName}")`,
    { state: "detached" }
  );
}

export async function validateCreatedCard(page: Page) {
  try {
    const frameElementHandle = await page.waitForSelector(
      "iframe.embedded-page-content"
    );
    const frame = await frameElementHandle?.contentFrame();
    console.log("start to created card");
    try {
      await frame
        ?.waitForSelector('div.ui-box button:has-text("Submit")', {
          timeout: Timeout.playwrightDefaultTimeout,
        })
        .catch(() => {});
    } catch (error) {
      console.log("no created card window");
    }
    const submitBtn = await frame?.waitForSelector(
      'div.ui-box button:has-text("Submit")'
    );
    await submitBtn?.click();
    try {
      await page.waitForTimeout(Timeout.shortTimeLoading);
      await frame?.waitForSelector("card div.card__react-wrapper");
      console.log("verify created card successfully!");
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

export async function validateUnfurlCard(page: Page) {
  try {
    const frameElementHandle = await page.waitForSelector(
      "iframe.embedded-page-content"
    );
    const frame = await frameElementHandle?.contentFrame();
    console.log("start to validate unfurl an adaptive card");
    const unfurlurl = "https://www.botframework.com/";
    await frame?.press("div.ui-box input.ui-box", "Escape");
    const msgTxtbox = await frame?.waitForSelector("div[data-tid='ckeditor']");
    await msgTxtbox?.focus();
    await msgTxtbox?.fill(unfurlurl);
    await msgTxtbox?.press("Space");
    await page.waitForTimeout(Timeout.shortTimeLoading);
    await frame?.waitForSelector('p:has-text("Link Unfurling card")');
    console.log("verify unfurl card successfully!");
  } catch (error) {
    await page.screenshot({
      path: getPlaywrightScreenshotPath("error"),
      fullPage: true,
    });
    throw error;
  }
}

export async function validateTabApim(
  page: Page,
  options?: { displayName?: string }
) {
  try {
    const frameElementHandle = await page.waitForSelector(
      "iframe.embedded-iframe"
    );
    const frame = await frameElementHandle?.contentFrame();

    const startBtn = await frame?.waitForSelector(
      'button:has-text("Consent and log in")'
    );

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
    });

    await frame?.waitForSelector(`div:has-text("${options?.displayName}")`);
  } catch (error) {
    await page.screenshot({
      path: getPlaywrightScreenshotPath("error"),
      fullPage: true,
    });
    throw error;
  }
}

export async function validateSearchCmdResult(
  page: Page,
  teamsAppName: string,
  envName: string
) {
  try {
    const frameElementHandle = await page.waitForSelector(
      "iframe.embedded-page-content"
    );
    const frame = await frameElementHandle?.contentFrame();
    console.log("start to validate search command");
    await frame?.click('button[name="message-extension-flyout-command"]');
    const input = await frame?.waitForSelector("div.ui-box input.ui-box");
    const appName = teamsAppName + envName;
    await input?.type(appName);
    await frame?.click(`span:has-text("${appName}")`);
    const searchcmdInput = await frame?.waitForSelector(
      "div.ui-box input.ui-box"
    );
    await searchcmdInput?.type("Karin");
    try {
      await frame?.waitForSelector('ul[datatid="app-picker-list"]');
      console.log("verify search successfully!!!");
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

export async function validateLargeNotificationBot(
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
        console.log("status code: ", result.status);
        if (result.status !== 202) {
          throw new Error(
            `POST /api/notification failed: status code: '${result.status}', body: '${result.data}'`
          );
        }
        console.log("Successfully sent notification");
      } catch (e: any) {
        console.log(e);
      }
      try {
        await frame?.waitForSelector('p:has-text("Hello World")');
      } catch (e) {
        throw e;
      }
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

export async function validateTodoListSpfx(page: Page) {
  try {
    console.log("start to verify todo list spfx");
    try {
      console.log("check result...");
      const frameElementHandle = await page.waitForSelector(
        "iframe.embedded-page-content"
      );
      const frame = await frameElementHandle?.contentFrame();
      const spfxFrame = frame?.childFrames()[0];
      // title
      console.log("check title");
      const title = await spfxFrame?.waitForSelector(
        "h2:has-text('To Do List')"
      );
      const titleContext = await title?.innerText();
      expect(titleContext).to.equal("To Do List");
      // task check
      console.log("check task");
      const task = await spfxFrame?.waitForSelector(
        "div.item .name:has-text('ttktest')"
      );
      console.log(await task?.inputValue());
      expect(task).to.not.be.undefined;

      console.log("verify finish!!!");
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
