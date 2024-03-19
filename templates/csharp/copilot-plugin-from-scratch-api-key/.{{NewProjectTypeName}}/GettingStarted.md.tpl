# Welcome to Teams Toolkit!

## Quick Start

> **Prerequisites**
>
> To run this app template in your local dev machine, you will need:
>
> - [Visual Studio 2022](https://aka.ms/vs) 17.9 or higher and [install Teams Toolkit](https://aka.ms/install-teams-toolkit-vs)
> - A [Microsoft 365 account for development](https://docs.microsoft.com/microsoftteams/platform/toolkit/accounts).

### Add your own API Key

1. Open PowerShell, change the current working directory to this project root and run command `./GenerateApiKey.ps1`
    ```
    > ./GenerateApiKey.ps1
    ```

2. The above command will output something like "Generated a new API Key: xxx...".
3. Fill in API Key into `env/.env.*.user`.
    ```
    SECRET_API_KEY=<your-api-key>
    ```

### Debug app in Teams Web Client

1. If you haven't added your own API Key, please follow the above steps to add your own API Key.
2. In the debug dropdown menu, select `Dev Tunnels > Create a Tunnel` (set authentication type to Public) or select an existing public dev tunnel
</br>![image](https://raw.githubusercontent.com/OfficeDev/TeamsFx/dev/docs/images/visualstudio/debug/create-devtunnel-button.png).
3. Right-click your `{{NewProjectTypeName}}` project and select `Teams Toolkit > Prepare Teams App Dependencies`.
4. If prompted, sign in with a Microsoft 365 account for the Teams organization you want to install the app to.
5. Press F5, or select the `Debug > Start Debugging` menu in Visual Studio
</br>![image](https://raw.githubusercontent.com/OfficeDev/TeamsFx/dev/docs/images/visualstudio/debug/debug-button.png)
6. When Teams launches in the browser, you can navigate to a chat message and [trigger your search commands from compose message area](https://learn.microsoft.com/microsoftteams/platform/messaging-extensions/what-are-messaging-extensions?tabs=dotnet#search-commands).


## Learn more

- [Extend Teams platform with APIs](https://aka.ms/teamsfx-api-plugin)

## Report an issue

Select Visual Studio > Help > Send Feedback > Report a Problem.
Or, you can create an issue directly in our GitHub repository:
https://github.com/OfficeDev/TeamsFx/issues
