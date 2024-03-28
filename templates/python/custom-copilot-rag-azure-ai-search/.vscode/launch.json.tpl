{
    "version": "0.2.0",
    "configurations": [
        {
            "name": "Launch Remote in Teams (Edge)",
            "type": "msedge",
            "request": "launch",
            "url": "https://teams.microsoft.com/l/app/${{TEAMS_APP_ID}}?installAppPackage=true&webjoin=true&${account-hint}",
            "presentation": {
                "group": "group 1: Teams",
                "order": 3
            },
            "internalConsoleOptions": "neverOpen"
        },
        {
            "name": "Launch Remote in Teams (Chrome)",
            "type": "chrome",
            "request": "launch",
            "url": "https://teams.microsoft.com/l/app/${{TEAMS_APP_ID}}?installAppPackage=true&webjoin=true&${account-hint}",
            "presentation": {
                "group": "group 1: Teams",
                "order": 3
            },
            "internalConsoleOptions": "neverOpen"
        },
        {
            "name": "Launch App (Edge)",
            "type": "msedge",
            "request": "launch",
            "url": "https://teams.microsoft.com/l/app/${{local:TEAMS_APP_ID}}?installAppPackage=true&webjoin=true&${account-hint}",
            "presentation": {
                "group": "all",
                "hidden": true
            },
            "internalConsoleOptions": "neverOpen"
        },
        {
            "name": "Launch App (Chrome)",
            "type": "chrome",
            "request": "launch",
            "url": "https://teams.microsoft.com/l/app/${{local:TEAMS_APP_ID}}?installAppPackage=true&webjoin=true&${account-hint}",
            "presentation": {
                "group": "all",
                "hidden": true
            },
            "internalConsoleOptions": "neverOpen"
        },
        {
            "name": "Start Python",
            "type": "debugpy",
            "program": "${workspaceFolder}/src/app.py",
            "request": "launch",
            "cwd": "${workspaceFolder}",
            "console": "integratedTerminal",
        },
        {
            "name": "Start Test Tool",
            "type": "node",
            "request": "launch",
            "program": "${workspaceFolder}/devTools/teamsapptester/node_modules/@microsoft/teams-app-test-tool/cli.js",
            "args": [
                "start",
            ],
            "cwd": "${workspaceFolder}",
            "console": "integratedTerminal",
            "internalConsoleOptions": "neverOpen"
        }
    ],
    "compounds": [
        {
            "name": "Debug (Edge)",
            "configurations": [
                "Launch App (Edge)",
                "Start Python"
            ],
            "cascadeTerminateToConfigurations": [
                "Start Python"
            ],
            "preLaunchTask": "Start Teams App Locally",
            "presentation": {
{{#enableTestToolByDefault}}
                "group": "2-local",
{{/enableTestToolByDefault}}
{{^enableTestToolByDefault}}
                "group": "1-local",
{{/enableTestToolByDefault}}
                "order": 1
            },
            "stopAll": true
        },
        {
            "name": "Debug (Chrome)",
            "configurations": [
                "Launch App (Chrome)",
                "Start Python"
            ],
            "cascadeTerminateToConfigurations": [
                "Start Python"
            ],
            "preLaunchTask": "Start Teams App Locally",
            "presentation": {
{{#enableTestToolByDefault}}
                "group": "2-local",
{{/enableTestToolByDefault}}
{{^enableTestToolByDefault}}
                "group": "1-local",
{{/enableTestToolByDefault}}
                "order": 2
            },
            "stopAll": true
        },
        {
            "name": "Debug in Test Tool (Preview)",
            "configurations": [
                "Start Python",
                "Start Test Tool",
            ],
            "cascadeTerminateToConfigurations": [
                "Start Test Tool"
            ],
            "preLaunchTask": "Deploy (Test Tool)",
            "presentation": {
{{#enableTestToolByDefault}}
                "group": "1-local",
{{/enableTestToolByDefault}}
{{^enableTestToolByDefault}}
                "group": "2-local",
{{/enableTestToolByDefault}}
                "order": 1
            },
            "stopAll": true
        }
    ]
}