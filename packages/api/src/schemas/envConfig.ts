/* tslint:disable */
/**
 * This file was automatically generated by json-schema-to-typescript.
 * DO NOT MODIFY IT BY HAND. Instead, modify the source JSONSchema file,
 * and run json-schema-to-typescript to regenerate this file.
 */

/**
 * The schema of TeamsFx configuration.
 */
export interface EnvConfig {
  $schema?: string;
  description?: string;
  /**
   * Existing AAD app configuration.
   */
  auth?: {
    /**
     * The client id of existing AAD app for Teams app.
     */
    clientId?: string;
    /**
     * The client secret of existing AAD app for Teams app.
     */
    clientSecret?: string;
    /**
     * The object id of existing AAD app for Teams app.
     */
    objectId?: string;
    /**
     * The access_as_user scope id of existing AAD app for Teams app.
     */
    accessAsUserScopeId?: string;
    /**
     * The frontend endpoint for redirect URLs of existing AAD app for Teams app.
     */
    frontendEndpoint?: string;
    /**
     * The frontend domain for redirect URLs of existing AAD app for Teams app.
     */
    frontendDomain?: string;
    /**
     * The bot id for identifier URIs of existing AAD app for Teams app.
     */
    botId?: string;
    /**
     * The bot endpoint for redirect URLs of existing AAD app for Teams app.
     */
    botEndpoint?: string;
    [k: string]: unknown;
  };
  /**
   * The Azure resource related configuration.
   */
  azure?: {
    /**
     * The default subscription to provision Azure resources.
     */
    subscriptionId?: string;
    /**
     * The default resource group of Azure resources.
     */
    resourceGroupName?: string;
    [k: string]: unknown;
  };
  /**
   * Existing bot AAD app configuration.
   */
  bot?: {
    /**
     * The id of existing bot AAD app.
     */
    appId?: string;
    /**
     * The password of existing bot AAD app.
     */
    appPassword?: string;
    [k: string]: unknown;
  };
  /**
   * The Teams App manifest related configuration.
   */
  manifest: {
    /**
     * Teams app name.
     */
    appName: {
      /**
       * A short display name for teams app.
       */
      short: string;
      /**
       * The full name for teams app.
       */
      full?: string;
      [k: string]: unknown;
    };
    /**
     * Description for Teams app.
     */
    description?: {
      /**
       * A short description of the app used when space is limited. Maximum length is 80 characters.
       */
      short?: string;
      /**
       * The full description of the app. Maximum length is 4000 characters.
       */
      full?: string;
      [k: string]: unknown;
    };
    /**
     * Icons for Teams App.
     */
    icons?: {
      /**
       * A relative file path to a full color PNG icon. Size 192x192.
       */
      color?: string;
      /**
       * A relative file path to a transparent PNG outline icon. The border color needs to be white. Size 32x32.
       */
      outline?: string;
      [k: string]: unknown;
    };
    [k: string]: unknown;
  };
  /**
   * Skip to add user during SQL provision.
   */
  skipAddingSqlUser?: boolean;
  [k: string]: unknown;
}
