// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

interface M365AppCommonMetadata {
  // isM365Project: boolean;
  m365Metadata?: M365Metadata;
  lauguages?: ("ts" | "js" | "csharp" | "java" | "python" | "c")[];
}

export type M365AppMetadata = M365AppCommonMetadata &
  (
    | {
        appType: "Teams";
        metadata: TeamsAppMetadata;
      }
    | {
        appType: "Office";
        metadata: OfficeAddinMetadata;
      }
    | {
        appType: "SPFx";
      }
    | {
        appType: "DA"; // Declarative Agent
      }
    | {
        appType: "Invalid";
      }
  );

export type M365AppType = M365AppMetadata["appType"];

export interface M365Metadata {
  version: string;
  versionState: M365AppVersionState;
  projectId?: string;
}

export interface TeamsAppMetadata {
  hasTeamsManifest: boolean;
  manifestCapabilities?: string[];
  manifestAppId?: string;
  manifestVersion?: string;
  dependsOnTeamsJs?: boolean;
}

export enum M365AppVersionState {
  Compatible = "compatible",
  Upgradable = "upgradable",
  Unsupported = "unsupported",
  Invalid = "invalid",
}

export interface OfficeAddinMetadata {
  officeAddinProjectType?: string;
  manifestOnly: boolean;
}
