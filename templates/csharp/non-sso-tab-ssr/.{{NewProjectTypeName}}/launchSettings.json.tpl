{
  "profiles": {
    // Debug project within Teams
    "Microsoft Teams (browser)": {
      "commandName": "Project",
      "launchUrl": "https://teams.microsoft.com/l/app/${{TEAMS_APP_ID}}?installAppPackage=true&webjoin=true&appTenantId=${{TEAMS_APP_TENANT_ID}}&login_hint=${{TEAMSFX_M365_USER_NAME}}",
    },
    // Debug project within Microsoft 365
    "Microsoft 365 app (browser)": {
      "commandName": "Project",
      "launchUrl": "https://www.office.com/m365apps/${{M365_APP_ID}}?auth=2&login_hint=${{TEAMSFX_M365_USER_NAME}}",
    },
    // Debug project within Outlook
    "Outlook (browser)": {
      "commandName": "Project",
      "launchUrl": "https://outlook.office.com/host/${{M365_APP_ID}}?login_hint=${{TEAMSFX_M365_USER_NAME}}",
    }
  }
}