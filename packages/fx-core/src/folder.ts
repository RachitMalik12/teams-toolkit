// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
"use strict";

import * as fs from "fs";
import * as path from "path";
export function getTemplatesFolder(): string {
  return path.resolve(__dirname, "../templates");
}

export function getResourceFolder(): string {
  return path.resolve(__dirname, "../resource");
}

const ignoreFolderName = ["node_modules", "bin", "build", "dist", ".vscode", ".venv"];
export function getAllFiles(dirPath: string, depth: number, currentDepth = 0): string[] {
  if (!dirPath) return [];
  const baseName = path.parse(dirPath).base;
  if (ignoreFolderName.includes(baseName)) {
    return [];
  }
  let files: string[] = [];
  if (currentDepth > depth) return files;

  const items = fs.readdirSync(dirPath);
  for (const item of items) {
    const fullPath = path.join(dirPath, item);
    const stat = fs.statSync(fullPath);

    if (stat.isDirectory()) {
      files = files.concat(getAllFiles(fullPath, depth, currentDepth + 1));
    } else {
      files.push(fullPath);
    }
  }
  return files;
}
