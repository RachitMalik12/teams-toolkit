// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
"use strict";

import { UserInteraction } from "../qm/ui";
import { CryptoProvider } from "./crypto";
import { ExpServiceProvider } from "./exp";
import { LogProvider } from "./log";
import { TokenProvider } from "./login";
import { TelemetryReporter } from "./telemetry";

export * from "./crypto";
export * from "./exp";
export * from "./log";
export * from "./login";
export * from "./telemetry";

export interface Tools {
  logProvider: LogProvider;
  tokenProvider: TokenProvider;
  telemetryReporter?: TelemetryReporter;
  ui: UserInteraction;
  cryptoProvider?: CryptoProvider;
  expServiceProvider?: ExpServiceProvider;
}
