// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * Help document
 * https://github.com/microsoft/vscode-azure-ext-sdk#azure-account-and-login
 */
import { Event, Terminal, Progress, CancellationToken } from "vscode";
import { TokenCredentialsBase } from "@azure/ms-rest-nodeauth";
import { Environment } from "@azure/ms-rest-azure-env";
import { SubscriptionModels } from "@azure/arm-subscriptions";

export type AzureLoginStatus = "Initializing" | "LoggingIn" | "LoggedIn" | "LoggedOut";

export interface AzureAccount {
	readonly status: AzureLoginStatus;
	readonly onStatusChanged: Event<AzureLoginStatus>;
	readonly waitForLogin: () => Promise<boolean>;
	readonly sessions: AzureSession[];
	readonly onSessionsChanged: Event<void>;
	readonly subscriptions: AzureSubscription[];
	readonly onSubscriptionsChanged: Event<void>;
	readonly waitForSubscriptions: () => Promise<boolean>;
	readonly filters: AzureResourceFilter[];
	readonly onFiltersChanged: Event<void>;
	readonly waitForFilters: () => Promise<boolean>;
}

export interface AzureSession {
	readonly environment: Environment;
	readonly userId: string;
	readonly tenantId: string;

	/**
	 * The credentials object for azure-sdk-for-js modules https://github.com/azure/azure-sdk-for-js
	 */
	readonly credentials2: TokenCredentialsBase;
}

export interface AzureSubscription {
	readonly session: AzureSession;
	readonly subscription: SubscriptionModels.Subscription;
}

export type AzureResourceFilter = AzureSubscription;