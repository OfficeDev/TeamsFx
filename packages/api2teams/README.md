# Welcome to `api2teams`
> The `api2teams` and its generated project template is currently under active development. Report any issues to us [here](https://github.com/OfficeDev/TeamsFx/issues/new/choose)

`api2teams` is a command line tool to generate a complete conversational style command and response [Teams application](https://learn.microsoft.com/microsoftteams/platform/bots/how-to/conversations/command-bot-in-teams) based on your Open API specification file and represent the API response in the form of [Adaptive Cards](https://learn.microsoft.com/microsoftteams/platform/task-modules-and-cards/cards/cards-reference#adaptive-card).

`api2teams` is the best way to start integrating your APIs with Teams conversational experience.

## Quick start

- Install `api2teams` with npm: `npm install @microsoft/api2teams@latest -g`
- Prepare the Open API specification. If you don't currently have one, start with a sample we provided by saving a copy of the [sample-open-api-spec.yml](https://raw.githubusercontent.com/OfficeDev/TeamsFx/api2teams/packages/api2teams/sample-spec/sample-open-api-spec.yml) to your local disk.
- Convert the Open API spec to a Teams app, assuming you are using the `sample-open-api-spec.yml`: `api2teams sample-open-api-spec.yml`

## Available commands and options

The CLI name is `api2teams`. Usage is as below:

```
Usage: api2teams [options] <yaml>

Convert open api spec file to Teams APP project, only for GET operation

Arguments:
  yaml                   yaml file path to convert

Options:
  -o, --output [string]  output folder for teams app (default: "./generated-teams-app")
  -f, --force            force overwrite the output folder
  -v, --version          output the current version
  -h, --help             display help for command
```

You can input below command to generate Teams App to default or specific folder:

```bash
api2teams sample-open-api-spec.yml # generate teams app to default folder ./generated-teams-app
api2teams sample-open-api-spec.yml -o ./my-app # generate teams app to ./my-app folder
api2teams sample-open-api-spec.yml -o ./my-app -f # generate teams app to ./my-app folder, and force overwrite output folder
api2teams -h # show help message
api2teams -v # show version information
```

## Getting started with the generated Teams app

- Open the generated project in [Visual Studio Code](https://code.visualstudio.com/) and make sure you have the latest [Teams Toolkit](https://marketplace.visualstudio.com/items?itemName=TeamsDevApp.ms-teams-vscode-extension) (version 5.0.0 or higher) installed.

- Follow the instruction provided in the `README.md` for the generated project to get started. For the Teams app converted by the given sample Open API spec, you will be able to run a `GET /pets/1` command in Teams and a bot will return an Adaptive Card as response.

    ![response](https://github.com/OfficeDev/TeamsFx/wiki/api2teams/workflow1.png)
    
## Current limitations
1. The `api2teams` doesn't support Open API schema version < 3.0.0.
1. The `api2teams` doesn't support Authorization property in Open API specification.
1. The `api2teams` doesn't support `webhooks` property and it would be ignored during convert.
1. The `api2teams` doesn't support `oneOf`, `anyOf`, `not`keyword (It only support `allOf` keyword currently).
1. The `api2teams` doesn't support `POST`, `PUT`, `PATCH` or `DELETE` operations (It only supports `GET` operation currently).
1. The generated Adaptive Card doesn't support array type. 
1. The generated Adaptive Card doesn't support file upload.
1. The generated Teams app can only contain up to 10 items in the command menu.

## Further reading
- [Teams Toolkit](https://learn.microsoft.com/microsoftteams/platform/toolkit/teams-toolkit-fundamentals)
- [Teams Platform Developer Documentation](https://learn.microsoft.com/microsoftteams/platform/mstdd-landing)
- [Adaptive Card Designer](https://adaptivecards.io/designer)
- [Swagger Parser](https://github.com/APIDevTools/swagger-parser)
- [Swagger Samples](https://github.com/OAI/OpenAPI-Specification)