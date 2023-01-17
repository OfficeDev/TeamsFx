{
  "name": "{%appName%}",
  "version": "1.0.0",
  "msteams": {
    "teamsAppId": null
  },
  "description": "Microsoft Teams Toolkit message extension Bot sample",
  "engines": {
    "node": ">=14 <=16"
  },
  "author": "Microsoft",
  "license": "MIT",
  "main": "index.js",
  "scripts": {
    "dev:teamsfx": "node script/run.js . env/.env.local",
    "dev": "nodemon --inspect=9239 --signal SIGINT ./index.js",
    "start": "node ./index.js",
    "watch": "nodemon ./index.js",
    "test": "echo \"Error: no test specified\" && exit 1"
  },
  "dependencies": {
    "@microsoft/adaptivecards-tools": "^1.0.0",
    "botbuilder": "^4.18.0",
    "isomorphic-fetch": "^3.0.0",
    "restify": "^10.0.0"
  },
  "devDependencies": {
    "@microsoft/teamsfx-run-utils": "alpha",
    "nodemon": "^2.0.7"
  }
}
