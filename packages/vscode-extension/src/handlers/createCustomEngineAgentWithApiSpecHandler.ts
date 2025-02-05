// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import {
  CreateProjectResult,
  err,
  FxError,
  ok,
  Result,
  Stage,
  UserError,
} from "@microsoft/teamsfx-api";
import { getSystemInputs } from "../utils/systemEnvUtils";
import {
  ApiPluginStartOptions,
  CapabilityOptions,
  CustomCopilotRagOptions,
  KiotaLastCommands,
  ProjectTypeOptions,
  QuestionNames,
} from "@microsoft/teamsfx-core";
import { runCommand } from "./sharedOpts";
import * as vscode from "vscode";
import { openFolder } from "../utils/workspaceUtils";
import { ExtensionSource } from "../error/error";
import { ExtTelemetry } from "../telemetry/extTelemetry";
import { getTriggerFromProperty } from "../utils/telemetryUtils";
import {
  TelemetryEvent,
  TelemetryProperty,
  TelemetrySuccess,
} from "../telemetry/extTelemetryEvents";
import { localize } from "../utils/localizeUtils";

export async function createCustomEngineAgentWithApiSpec(
  args?: any[]
): Promise<Result<any, FxError>> {
  ExtTelemetry.sendTelemetryEvent(
    TelemetryEvent.CreatePluginWithManifestStart,
    getTriggerFromProperty(args)
  );
  if (!args || args.length > 99) {
    const error = new UserError(
      ExtensionSource,
      "invalidParameter",
      "Invalid parameter for creating Custom Engine Agent with API spec."
    );
    ExtTelemetry.sendTelemetryErrorEvent(TelemetryEvent.CreatePluginWithManifest, error);
    return err(error);
  }

  const specPath = args[0];

  const inputs = getSystemInputs();
  inputs[QuestionNames.ProjectType] = ProjectTypeOptions.customCopilot().id;
  inputs.capabilities = CapabilityOptions.customCopilotRag().id;
  inputs[QuestionNames.CustomCopilotRag] = CustomCopilotRagOptions.customApi().id;
  inputs[QuestionNames.ApiSpecLocation] = specPath;

  const result = await runCommand(Stage.create, inputs);

  if (result.isErr()) {
    ExtTelemetry.sendTelemetryErrorEvent(TelemetryEvent.CreatePluginWithManifest, result.error);
    return err(result.error);
  }

  const res = result.value as CreateProjectResult;
  const projectPathUri = vscode.Uri.file(res.projectPath);
  await openFolder(projectPathUri, true, res.warnings);
  ExtTelemetry.sendTelemetryEvent(TelemetryEvent.CreatePluginWithManifest, {
    [TelemetryProperty.Success]: TelemetrySuccess.Yes,
  });

  return ok({});
}
