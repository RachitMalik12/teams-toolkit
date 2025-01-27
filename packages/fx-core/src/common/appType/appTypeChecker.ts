// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { ConfigFolderName } from "@microsoft/teamsfx-api";
import fs from "fs-extra";
import path from "path";
import { manifestUtils } from "../../component/driver/teamsApp/utils/ManifestUtils";
import { getAllFiles } from "../../folder";
import { validateProjectSettings } from "../projectSettingsHelper";
import { MetadataV3 } from "../versionMetadata";
import { M365AppMetadata } from "./appMetadata";

export const TeamsJsModule = "@microsoft/teams-js";
export const SPFxKey = "@microsoft/generator-sharepoint";

export function isValidProject(workspacePath?: string): boolean {
  if (!workspacePath) return false;
  try {
    return isValidProjectV3(workspacePath) || isValidProjectV2(workspacePath);
  } catch (e) {
    return false;
  }
}

export function isValidProjectV3(workspacePath: string): boolean {
  const ymlFilePath = path.join(workspacePath, MetadataV3.configFile);
  const localYmlPath = path.join(workspacePath, MetadataV3.localConfigFile);
  if (fs.pathExistsSync(ymlFilePath) || fs.pathExistsSync(localYmlPath)) {
    return true;
  }
  return false;
}

export function isValidProjectV2(workspacePath: string): boolean {
  const confFolderPath = path.resolve(workspacePath, `.${ConfigFolderName}`, "configs");
  const settingsFile = path.resolve(confFolderPath, "projectSettings.json");
  if (!fs.existsSync(settingsFile)) {
    return false;
  }
  const projectSettings: any = fs.readJsonSync(settingsFile);
  if (validateProjectSettings(projectSettings)) return false;
  return true;
}

export function isSPFxProject(files: string[]): boolean {
  for (const file of files) {
    if (path.basename(file) === ".yo-rc.json") {
      const content = fs.readJsonSync(file) as Record<string, unknown>;
      if (content["@microsoft/generator-sharepoint"]) {
        return true;
      }
    }
  }
  return false;
}

export function isValidOfficeAddInProject(files: string[]): boolean {
  try {
    const xmlManifestList = files.filter((fileName) => isOfficeXmlAddInManifest(fileName));
    const metaOsManifestList = files.filter((fileName) => isOfficeMetaOsAddInManifest(fileName));
    return (
      xmlManifestList &&
      xmlManifestList.length > 0 &&
      (!metaOsManifestList || metaOsManifestList.length == 0)
    );
  } catch (e) {
    return false;
  }
}

function isOfficeXmlAddInManifest(inputFileName: string): boolean {
  return (
    inputFileName.toLocaleLowerCase().indexOf("manifest") != -1 &&
    inputFileName.toLocaleLowerCase().endsWith(".xml")
  );
}

function isOfficeMetaOsAddInManifest(inputFileName: string): boolean {
  return (
    inputFileName.toLocaleLowerCase().indexOf("manifest") != -1 &&
    inputFileName.toLocaleLowerCase().endsWith(".json")
  );
}

class ProjectTypeChecker {
  public async checkProjectType(projectPath: string): Promise<M365AppMetadata> {
    if (!isValidProject(projectPath)) {
      return {
        // isM365Project: false,
        appType: "Invalid",
      };
    }
    const files = getAllFiles(projectPath, 2);
    // const languages = findUsedLanguages(files);
    if (isSPFxProject(files)) {
      return {
        appType: "SPFx",
      };
    } else if (isValidOfficeAddInProject(files)) {
      const srcPath = path.join(projectPath, "src");
      return {
        appType: "Office",
        metadata: {
          officeAddinProjectType: "XML",
          manifestOnly: !fs.existsSync(srcPath),
        },
      };
    } else {
      // Teams app or Declarative Agent
      const manifestRes = await manifestUtils.readAppManifest(projectPath);
      if (manifestRes.isErr()) {
        return {
          appType: "Invalid",
        };
      }
      if (IsDeclarativeAgentManifest(manifestRes.value)) {
        return {
          appType: "DA",
        };
      }
      return {
        appType: "Teams",
        metadata: {
          hasTeamsManifest: true,
        },
      };
    }
  }
}

export function getCapabilities(manifest: any): string[] {
  const capabilities: string[] = [];
  if (manifest.staticTabs && manifest.staticTabs.length > 0) {
    capabilities.push("staticTab");
  }
  if (manifest.configurableTabs && manifest.configurableTabs.length > 0) {
    capabilities.push("configurableTab");
  }
  if (manifest.bots && manifest.bots.length > 0) {
    capabilities.push("bot");
  }
  if (manifest.composeExtensions && manifest.composeExtensions.length > 0) {
    capabilities.push("composeExtension");
  }
  if (manifest.extensions && manifest.extensions.length > 0) {
    capabilities.push("extension");
  }
  if (manifest.copilotExtensions?.plugins && manifest.copilotExtensions.plugins.length > 0) {
    capabilities.push("plugin");
  }
  if (
    manifest.copilotExtensions?.declarativeCopilots &&
    manifest.copilotExtensions.declarativeCopilots.length > 0
  ) {
    capabilities.push("copilotGpt");
  }
  if (
    manifest.copilotAgents?.plugins &&
    manifest.copilotAgents.plugins.length > 0 &&
    !capabilities.includes("plugin")
  ) {
    capabilities.push("plugin");
  }
  if (
    manifest.copilotAgents?.declarativeAgents &&
    manifest.copilotAgents.declarativeAgents.length > 0 &&
    !capabilities.includes("copilotGpt")
  ) {
    capabilities.push("copilotGpt");
  }
  return capabilities;
}

export function IsDeclarativeAgentManifest(manifest: any): boolean {
  return !!(
    manifest.copilotAgents?.declarativeAgents && manifest.copilotAgents.declarativeAgents.length > 0
  );
}
export const projectTypeChecker = new ProjectTypeChecker();
