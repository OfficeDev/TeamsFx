// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
export interface IRenderingTemplate {
  /**
   * Url of the adaptive card template.
   */
  template_url: string;
  /**
   * Operation id.
   */
  operationId: string;
}

export interface IFunctionParameter {
  /**
   * Type. Must be one of the following: string, array, boolean, integer, number
   */
  type: string;
  /**
   * Description
   */
  description?: string;
  /**
   * Describes a single element in an array to be used as a parameter.
   * Items must only be present when type is array.
   */
  items: IFunctionParameter;
  /**
   * Represent valid values for this parameter. This property can only be present when type is string.
   */
  enum: string[];
}

export interface IFunctionParameters {
  type?: string;
  /**
   * A JSON object that contains the a map of the function parameter names and the definition of those function parameters.
   */
  properties?: { [k: string]: IFunctionParameter };
  /**
   * Names of properties that are required parameters.
   */
  required?: string[];
}

export interface IConfirmation {
  /**
   * Indicates the type of confirmation dialog presented.
   */
  type: string;
  /**
   *  Title to be displayed on the confirmation dialog.
   */
  title: string;
  /**
   * Text to ask the user if they want the plugin can be called.
   */
  body: string;
}

export interface ICapability {
  /**
   * Adaptive card describes how a function response can be rendered.
   */
  rendering_templates?: { [k: string]: IRenderingTemplate }; // TODO: verify whether it exists or not

  /**
   * Description of a confirmation prompt that should be presented to the user before invoking the function.
   */
  confirmation?: IConfirmation;
}

export interface IReturn {
  /**
   * Type. Must be one of the following: string, array, boolean, integer, number.
   */
  type: string;
  /**
   * Describes the information returned from the function
   */
  description: string;
}

export interface IState {
  [k: string]: {
    /**
     * The purpose of the function when used in a specific orchestrator state
     */
    description: string;
    /**
     * Instructions to the orchestrator on how to use this function while in a specific orchestrator state
     */
    instructions: string | string[];
    /**
     *  Examples to the orchestrator on this function can be invoked
     */
    examples: string | string[];
  };
}

export interface IFunction {
  /**
   * Name which uniquely identifies the function.
   */
  name: string;
  /**
   * Description.
   */
  description?: string;
  /**
   * Identify the set of parameters that can be passed to the function
   */
  parameters?: IFunctionParameters;
  returns?: IReturn;
  states?: IState;
  /**
   * A collection of data used to configure optional capabilities of the orchestrator while invoking the function
   */
  capabilities?: ICapability;
}

export interface ISpec {
  /**
   * Local runtime identifier that links to a specific function
   */
  local_endpoint?: string;
  /**
   * Url of the OpenAPI spec.
   */
  url?: string;
  /**
   * The progress style that will be used to display the progress of the function.
   * The value MUST be one of the following values: none, showUsage, showUsageWithInput, showUsageWithInputAndOutput.
   */
  progress_style?: string;
}

export interface IAuth {
  type: string;
}

export interface IRuntime {
  type: "openapi" | "localPlugin";
  auth: IAuth;
  run_for_functions?: string[];
  spec: ISpec;
}

export interface PluginManifest {
  /**
   * The schema version of the manifest
   */
  schema_version: string;
  /**
   * Human-readable name for the plugin
   */
  name_for_human: string;
  /**
   * Description for the plugin.
   */
  description_for_model?: string;
  /**
   * Human-readable description for the plugin.
   */
  description_for_human: string;
  /**
   * URL used to fetach the logo.
   */
  logo_url?: string;
  /**
   * An emaill address of a contact.
   */
  contact_email?: string;
  /**
   * An absolute URL that locates a document containing the legal information.
   */
  legal_info_url?: string;
  /**
   * A JSON array that contains an set of function objects
   */
  functions?: IFunction[];
  /**
   * A JSON array that contains an set of runtime objects
   */
  runtimes?: IRuntime[];
}
