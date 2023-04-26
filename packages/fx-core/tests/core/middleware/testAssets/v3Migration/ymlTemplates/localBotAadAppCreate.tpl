  # Create or reuse an existing Azure Active Directory application for bot.
  - uses: botAadApp/create
    with:
      # The Azure Active Directory application's display name
      name: ${{CONFIG__MANIFEST__APPNAME__SHORT}}-bot
    writeToEnvironmentFile:
      # The Azure Active Directory application's client id created for bot.
      botId: BOT_ID
      # The Azure Active Directory application's client secret created for bot.
      botPassword: SECRET_BOT_PASSWORD