import { AdaptiveCards } from "@microsoft/adaptivecards-tools";
import * as restify from "restify";
import notificationTemplate from "./adaptiveCards/notification-default.json";
import { bot, server } from "./internal/initialize";
import { CardData } from "./cardModels";

// HTTP trigger to send notification. You need to add authentication / authorization for this API. Refer https://aka.ms/teamsfx-notification for more details.
server.post(
  "/api/notification",
  restify.plugins.queryParser(),
  restify.plugins.bodyParser(), // Add more parsers if needed
  async (req, res) => {
    for (const target of await bot.notification.installations()) {
      await target.sendAdaptiveCard(
        AdaptiveCards.declare<CardData>(notificationTemplate).render({
          title: "New Event Occurred!",
          appName: "Contoso App Notification",
          description: `This is a sample http-triggered notification to ${target.type}`,
          notificationUrl: "https://www.adaptivecards.io/",
        })
      );
    }

    /****** To distinguish different target types ******/
    /** "Channel" means this bot is installed to a Team (default to notify General channel)
    if (target.type === "Channel") {
      // Directly notify the Team (to the default General channel)
      await target.sendAdaptiveCard(...);

      // List all channels in the Team then notify each channel
      const channels = await target.channels();
      for (const channel of channels) {
        await channel.sendAdaptiveCard(...);
      }

      // List all members in the Team then notify each member
      const members = await target.members();
      for (const member of members) {
        await member.sendAdaptiveCard(...);
      }
    }
    **/

    /** "Group" means this bot is installed to a Group Chat
    if (target.type === "Group") {
      // Directly notify the Group Chat
      await target.sendAdaptiveCard(...);

      // List all members in the Group Chat then notify each member
      const members = await target.members();
      for (const member of members) {
        await member.sendAdaptiveCard(...);
      }
    }
    **/

    /** "Person" means this bot is installed as a Personal app
    if (target.type === "Person") {
      // Directly notify the individual person
      await target.sendAdaptiveCard(...);
    }
    **/

    res.json({});
  }
);

// Bot Framework message handler.
server.post("/api/messages", async (req, res) => {
  await bot.requestHandler(req, res);
});
