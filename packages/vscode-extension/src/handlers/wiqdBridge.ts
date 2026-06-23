// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * wiqd integration bridge for ATK VS Code extension.
 *
 * This module enables ATK to invoke wiqd capabilities (enhanced validation,
 * provision, publish) directly from the editor without requiring the user
 * to open a terminal. It leverages VS Code's command API to call wiqd's
 * registered commands when the wiqd extension is active.
 */

import * as vscode from "vscode";

export interface WiqdResult {
  success: boolean;
  stdout: string;
  stderr: string;
  exitCode: number;
}

/**
 * Check whether the wiqd VS Code extension is installed and active.
 */
export function isWiqdAvailable(): boolean {
  const wiqdExtension = vscode.extensions.getExtension("Microsoft.wiqd");
  return wiqdExtension !== undefined;
}

/**
 * Check whether the wiqd extension is active (already activated).
 */
export function isWiqdActive(): boolean {
  const wiqdExtension = vscode.extensions.getExtension("Microsoft.wiqd");
  return wiqdExtension?.isActive ?? false;
}

/**
 * Invoke wiqd's enhanced validation on the current project.
 * Falls back to ATK's built-in validation if wiqd is not available.
 *
 * @param projectPath - Path to the agent project root
 * @returns WiqdResult if wiqd handled it, undefined if wiqd is unavailable
 */
export async function invokeWiqdValidation(
  projectPath?: string
): Promise<WiqdResult | undefined> {
  if (!isWiqdAvailable()) {
    return undefined;
  }

  try {
    const result = await vscode.commands.executeCommand<WiqdResult>(
      "wiqd.validate",
      projectPath
    );
    return result;
  } catch (err) {
    // wiqd command failed — log and return undefined so ATK falls back
    const message = err instanceof Error ? err.message : String(err);
    console.warn(`[ATK] wiqd.validate failed: ${message}`);
    return undefined;
  }
}

/**
 * Invoke wiqd's provision command on the current project.
 *
 * @param projectPath - Path to the agent project root
 * @returns WiqdResult if wiqd handled it, undefined if wiqd is unavailable
 */
export async function invokeWiqdProvision(
  projectPath?: string
): Promise<WiqdResult | undefined> {
  if (!isWiqdAvailable()) {
    return undefined;
  }

  try {
    const result = await vscode.commands.executeCommand<WiqdResult>(
      "wiqd.provision",
      projectPath
    );
    return result;
  } catch (err) {
    const message = err instanceof Error ? err.message : String(err);
    console.warn(`[ATK] wiqd.provision failed: ${message}`);
    return undefined;
  }
}

/**
 * Invoke wiqd's publish command on the current project.
 *
 * @param projectPath - Path to the agent project root
 * @returns WiqdResult if wiqd handled it, undefined if wiqd is unavailable
 */
export async function invokeWiqdPublish(
  projectPath?: string
): Promise<WiqdResult | undefined> {
  if (!isWiqdAvailable()) {
    return undefined;
  }

  try {
    const result = await vscode.commands.executeCommand<WiqdResult>(
      "wiqd.publish",
      projectPath
    );
    return result;
  } catch (err) {
    const message = err instanceof Error ? err.message : String(err);
    console.warn(`[ATK] wiqd.publish failed: ${message}`);
    return undefined;
  }
}

/**
 * Run the full wiqd lifecycle (validate → provision → publish).
 *
 * @returns WiqdResult from the last successful step, or undefined if wiqd is unavailable
 */
export async function invokeWiqdFullLifecycle(): Promise<void> {
  if (!isWiqdAvailable()) {
    vscode.window.showInformationMessage(
      "Install the wiqd extension for enhanced agent lifecycle management."
    );
    return;
  }

  await vscode.commands.executeCommand("wiqd.lifecycle");
}

/**
 * Offer wiqd validation as an enhancement after ATK's built-in validation.
 * Shows a notification asking if the user wants deeper wiqd validation rules.
 */
export async function offerWiqdEnhancedValidation(projectPath?: string): Promise<void> {
  if (!isWiqdAvailable()) return;

  const choice = await vscode.window.showInformationMessage(
    "Run enhanced validation with wiqd? (Catches additional semantic errors beyond schema validation.)",
    "Run wiqd Validate",
    "Skip"
  );

  if (choice === "Run wiqd Validate") {
    await invokeWiqdValidation(projectPath);
  }
}
