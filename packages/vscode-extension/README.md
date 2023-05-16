# Microsoft Teams Toolkit for Visual Studio Code
## What is Teams Toolkit?

The Teams Toolkit helps developers create and deploy Teams apps with integrated Identity, access to cloud storage, data from [Microsoft Graph](https://docs.microsoft.com/en-us/graph/teams-concept-overview), and other services in [Azure](https://docs.microsoft.com/en-us/microsoftteams/platform/build-your-first-app/build-bot) and [Microsoft 365](https://docs.microsoft.com/en-us/microsoftteams/platform/concepts/build-and-test/prepare-your-o365-tenant) with a "zero-configuration" approach to the developer experience.

<img src="https://raw.githubusercontent.com/OfficeDev/TeamsFx/main/packages/vscode-extension/img/landingPage_GetStart.gif">

## Getting started

Open Teams Toolkit to create a new Teams app, start coding!

<img src="https://raw.githubusercontent.com/OfficeDev/TeamsFx/main/packages/vscode-extension/img/create_project.gif">

follow the [Get Started](https://aka.ms/teamsfx-build-first-app) instructions in our documentation to start with new project.

In the Teams Toolkit for Visual Studio Code, you can easily discover all applicable commands in the sidebar and Command Palette with the keyword "Teams". It also supports [Command Line Interface (CLI)](https://www.npmjs.com/package/@microsoft/teamsfx-cli) to increase efficiency.

## Prerequisites

Verify you have the right prerequisites for building Teams apps and install some recommended development tools. [Read more details](https://docs.microsoft.com/en-us/microsoftteams/platform/build-your-first-app/build-first-app-overview).

<table>
    <tr>
        <td><img src="https://raw.githubusercontent.com/OfficeDev/TeamsFx/main/packages/vscode-extension/img/landingPage_nodejs.png"></td>
        <td><h3>Node.js</h3>As a fundamental runtime context for Teams app, Node.js v14.x, v16.x or v18.x is required. If you develop SPFx Tab app, please install v16.x.</td>
    </tr>
    <tr>
        <td><img src="https://raw.githubusercontent.com/OfficeDev/TeamsFx/main/packages/vscode-extension/img/landingPage_m365.png"></td>
        <td><h3>Microsoft 365</h3>The Teams Toolkit requires a Microsoft 365 organizational account where Teams is running and has been registered.</td>
    </tr>
    <tr>
        <td><img src="https://raw.githubusercontent.com/OfficeDev/TeamsFx/main/packages/vscode-extension/img/landingPage_azure.png"></td>
        <td><h3>Azure</h3> The Teams Toolkit may require an Azure account and subscription to deploy the Azure resources for your project.</td>
    </tr>
</table>

> Don’t have a Microsoft 365 account to experience building Teams app? Sign up for [Microsoft Developer Program](https://developer.microsoft.com/en-us/microsoft-365/dev-program), which allows you to have a testing tenant with preconfigured permissions.

### Create your project

Use the Teams Toolkit for Visual Studio Code to set up your first app project. Create your tab app project using the following steps:

- Ensure you've installed the Microsoft Teams Toolkit for Visual Studio Code
- Select the Teams Toolkit icon in the Visual Studio Code sidebar.
- Select `Create a new Teams app` button and at next prompt to also select `Create a new Teams app`.
- Selected the [capability](https://aka.ms/teamsfx-capabilities-overview) that you want to build in your app. Select OK.
- Select JavaScript as the programming language.
- Choose a location where your new application will be created in a new folder.
- Type a name for your project and hit Enter.

Read more on [creating new project](https://docs.microsoft.com/en-us/microsoftteams/platform/toolkit/create-new-project).

### Configure your app

At its core, the Teams app embraces three components:

- The Microsoft Teams client (web, desktop or mobile) where users interact with your app.
- A server that responds to requests for content that will be displayed in Teams, e.g., HTML tab content or a bot adaptive card .
- A Teams app package consisting of three files:

  ✔️ The manifest.source.json.

  ✔️ A color icon for your app to display in the public or organization app catalog.

  ✔️ An outline icon for display on the Teams activity bar.

When an app is installed, the Teams client parses the manifest file to determine needed information like the name of your app and the URL where the services are located.

- To configure your app, navigate to the `templates/appPackage/manifest.template.json` file or from UI to select `DEVELOPMENT`-`Edit manifest file` and edit the manifest.
- The toolkit will automatically update the app registration data accordingly during app side-loading and publish.

Read more on [Preview and customize manifest file](https://docs.microsoft.com/en-us/microsoftteams/platform/toolkit/teamsfx-preview-and-customize-app-manifest).

### Preview your app on your local/remote dev environment

Simply press F5 to run your first Teams app or navigate to the Debug tab in the activity bar and select Run icon to display the Run and Debug view. As a default, the toolkit will automatically help you to setup local environment and load the app in Teams.

<img src="https://raw.githubusercontent.com/OfficeDev/TeamsFx/main/packages/vscode-extension/img/debug.gif">

If you want to have a better estimation of how the app will behave in the cloud environment, you can deploy your resources to the cloud and preview your app with the backend running in the cloud (remote).

Read more on [local debug](https://docs.microsoft.com/en-us/microsoftteams/platform/toolkit/debug-local).

### Deploy your application to Azure

Deployment happens in two steps: Provisioning and Deployment. Provisioning creates all the necessary Azure resources that your application uses. It is typically done once. To provision your backend to an active Azure subscription:


- In the Visual Studio Code Command Palette, enter "teams" and choose the command `Teams: Provision in the Cloud`.
- You will be asked to sign into your Azure account. This is the account where Azure resources will be provisioned. Typically this is different from the Microsoft 365 account you used to sign in earlier.
- You will be asked to select a subscription to use from the Azure account.
- Once provisioning is completed, Visual Studio Code will popup the notification with the message "'app name' successfully provisioned in the cloud.".

Deploy copies of your application to the provisioned Azure resources. It is typically done after every change to your application. To deploy your application to the provisioned resources in an active Azure subscription:


- In the Visual Studio Code Command Palette, enter "teams" and choose the command `Teams: Deploy to the Cloud`.
- Select Tab app and Backend to deploy.
- Once deploy is finished, go to the Visual Studio Code Debug Panel (Ctrl+Shift+D / ⌘⇧-D or View -> Run) and select Launch Remote (Edge).
- Press the start button (green arrow) to launch your app - now running remotely on Azure!

Read more on [provision](https://docs.microsoft.com/en-us/microsoftteams/platform/toolkit/provision) and [deploy](https://docs.microsoft.com/en-us/microsoftteams/platform/toolkit/deploy) to Azure.

### Publish your application to Teams

When your application resources and infrastructure are deployed successfully, you can publish and register your app to Teams app catalog to share with others in your organization.


- In the Visual Studio Code Command Palette, enter "teams" and choose the command `Teams: Publish to the Teams`.
- Depending on your permission, you can send your app to the admin portal directly, or manually submit the app package file to your admin to check.
- Once your app is approved by your admin, you can see it under "Built for your org" section in Teams Apps.

Read more on [publishing](https://docs.microsoft.com/en-us/microsoftteams/platform/toolkit/publish) Teams apps.

## Explore Code Samples

Explore our [samples](https://github.com/OfficeDev/TeamsFx-Samples) to help you quickly get started with the basic Teams app concepts and code structures.

## Contributing

There are many ways in which you can participate in the project, for example:

- [Download our latest builds](https://github.com/OfficeDev/TeamsFx/releases).
- [Submit bugs and feature requests](https://github.com/OfficeDev/TeamsFx/issues), and help us verify as they are checked in
- Review [source code changes](https://github.com/OfficeDev/TeamsFx/pulls)
- Review the [documentation](CONTRIBUTING.md) and make pull requests for anything from typos to new content

## Reporting security issues

Give security researchers information on how to privately report security vulnerabilities found in your open source project. See more details [Reporting security issues](https://docs.opensource.microsoft.com/content/releasing/security.html).

## Telemetry

The software may collect information about you and your use of the software and send it to Microsoft. Microsoft may use this information to provide services and improve our products and services. You may turn off the telemetry as described in the repository. There are also some features in the software that may enable you and Microsoft to collect data from users of your applications. If you use these features, you must comply with` applicable law, including providing appropriate notices to users of your applications together with a copy of Microsoft's privacy statement. Our privacy statement is located at [Microsoft Privacy Statement](https://go.microsoft.com/fwlink/?LinkID=824704). You can learn more about data collection and use in the help documentation and our privacy statement. Your use of the software operates as your consent to these practices.

### Telemetry Configuration

Telemetry collection is on by default. To opt out, please set the `telemetry.enableTelemetry` setting to `false`. Learn more in our [FAQ](https://code.visualstudio.com/docs/supporting/faq#_how-to-disable-telemetry-reporting).

## Code of conduct

See [Microsoft Open Source code of conduct](https://opensource.microsoft.com/codeofconduct).

## Trademark

This project may contain trademarks or logos for projects, products, or services. Authorized use of Microsoft trademarks or logos is subject to and must follow [Microsoft's Trademark & Brand Guidelines](https://www.microsoft.com/en-us/legal/intellectualproperty/trademarks/usage/general). Use of Microsoft trademarks or logos in modified versions of this project must not cause confusion or imply Microsoft sponsorship. Any use of third-party trademarks or logos are subject to those third-party's policies.

## License

Copyright (c) Microsoft Corporation. All rights reserved.

Licensed under the [MIT](LICENSE) license.
