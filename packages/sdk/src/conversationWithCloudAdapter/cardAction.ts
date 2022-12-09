import { CloudAdapter } from "botbuilder";
import { CardActionMiddleware } from "./middlewares/cardActionMiddleware";
import { CardActionOptions, TeamsFxAdaptiveCardActionHandler } from "./interface";

/**
 * A card action bot to respond to adaptive card universal actions.
 */
export class CardActionBot {
  private readonly adapter: CloudAdapter;
  private middleware: CardActionMiddleware;

  /**
   * Creates a new instance of the `CardActionBot`.
   *
   * @param adapter The bound `CloudAdapter`.
   * @param options - initialize options
   */
  constructor(adapter: CloudAdapter, options?: CardActionOptions) {
    this.middleware = new CardActionMiddleware(options?.actions);
    this.adapter = adapter.use(this.middleware);
  }

  /**
   * Registers a card action handler to the bot.
   * @param actionHandler A card action handler to be registered.
   */
  registerHandler(actionHandler: TeamsFxAdaptiveCardActionHandler) {
    if (actionHandler) {
      this.middleware.actionHandlers.push(actionHandler);
    }
  }

  /**
   * Registers card action handlers to the bot.
   * @param actionHandlers A set of card action handlers to be registered.
   */
  registerHandlers(actionHandlers: TeamsFxAdaptiveCardActionHandler[]) {
    if (actionHandlers) {
      this.middleware.actionHandlers.push(...actionHandlers);
    }
  }
}
