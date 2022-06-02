import { MessagingExtensionCommand } from "./messagingExtensionCommand";
import { MessagingExtensionMessageHandler } from "./messagingExtensionMessageHandler";

export interface MessagingExtension {
  objectId?: string;
  botId?: string;
  canUpdateConfiguration: boolean;
  commands: MessagingExtensionCommand[];
  messageHandlers: MessagingExtensionMessageHandler[];
}
