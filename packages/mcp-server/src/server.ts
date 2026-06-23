// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import { fetchSchema, SchemaTypeEnum } from "./fetcher";
import { retrieveResource } from "./retriever";
import { execFile } from "child_process";
import { promisify } from "util";

const execFileAsync = promisify(execFile);

/**
 * Resolve the wiqd binary path. Checks WIQD_PATH env var first, then falls
 * back to "wiqd" on PATH.
 */
function resolveWiqdBin(): string {
  return process.env.WIQD_PATH || "wiqd";
}

/**
 * Run a wiqd CLI command and return structured output.
 */
async function runWiqd(
  args: string[],
  cwd?: string
): Promise<{ success: boolean; stdout: string; stderr: string; exitCode: number }> {
  const bin = resolveWiqdBin();
  try {
    const { stdout, stderr } = await execFileAsync(bin, args, {
      cwd: cwd || process.cwd(),
      timeout: 120_000,
      windowsHide: true,
    });
    return { success: true, stdout, stderr, exitCode: 0 };
  } catch (err: any) {
    return {
      success: false,
      stdout: err.stdout || "",
      stderr: err.stderr || err.message || "",
      exitCode: err.code ?? 1,
    };
  }
}

export function createServer(): McpServer {
  const server = new McpServer({
    name: "m365agentstoolkit-mcp",
    version: "0.1.0",
  });
  server.tool(
    "get_schema",
    'Get the schema for "App manifest", "Declarative agent manifest", "API plugin manifest", use it everytime before understanding, modifying or creating any of these manifest files.',
    {
      schema_name: SchemaTypeEnum.describe("name of schema"),
      schema_version: z
        .string()
        .describe(
          'version of schema in semantic versioning format vX.Y, where X is the major version and Y is the minor version (e.g. v1.0, v1.19, v2.1). Use "latest" if unsure.'
        ),
    },
    async ({ schema_name, schema_version }) => {
      const schema = await fetchSchema(schema_name, schema_version);

      return {
        content: [
          {
            type: "text",
            text: schema,
          },
        ],
      };
    }
  );

  server.tool(
    "get_knowledge",
    "Access comprehensive knowledge about Microsoft 365 and Microsoft 365 Copilot development. Use this tool everytime for questions related to Microsoft 365 and Microsoft 365 Copilot.",
    {
      question: z.string().describe("Question to use for knowledge retrieval"),
    },
    async ({ question }) => {
      const result = await retrieveResource("documents", question);

      return {
        content: [
          {
            type: "text",
            text: result,
          },
        ],
      };
    }
  );

  server.tool(
    "get_code_snippets",
    "Access templates and code snippets for Microsoft 365 and Microsoft 365 Copilot development, focusing on SDKs such as **@microsoft/teams-ai**, **@microsoft/teams-js**, and **botbuilder**. Use this tool when looking for implementation examples, starter templates, or SDK usage patterns.",
    {
      question: z
        .string()
        .describe(
          "Query to find relevant code snippets related to Microsoft 365 app or agent SDKs"
        ),
    },
    async ({ question }) => {
      const result = await retrieveResource("code", question);

      return {
        content: [
          {
            type: "text",
            text: result,
          },
        ],
      };
    }
  );

  server.tool(
    "troubleshoot",
    "Access troubleshooting solutions for common Microsoft 365 and Microsoft 365 Copilot development issues. Use this tool when encountering errors, unexpected behaviors, or implementation challenges.",
    {
      question: z.string().describe("Description of the issue or error you're experiencing"),
    },
    async ({ question }) => {
      const result = await retrieveResource("issues", question);

      return {
        content: [
          {
            type: "text",
            text: result,
          },
        ],
      };
    }
  );

  // ---------------------------------------------------------------------------
  // wiqd Lifecycle Tools — enhanced agent lifecycle management
  // ---------------------------------------------------------------------------

  server.tool(
    "validate_agent",
    "Validate a declarative agent project using wiqd's enhanced validation rules. " +
      "Goes beyond schema validation to catch semantic errors, missing capabilities, " +
      "and configuration issues. Use this before provisioning or publishing an agent.",
    {
      project_path: z
        .string()
        .describe(
          "Absolute path to the declarative agent project root (must contain appPackage/declarativeAgent.json or m365agents.yml)"
        ),
    },
    async ({ project_path }) => {
      const result = await runWiqd(
        ["agent", "validate", "--path", project_path, "--json"],
        project_path
      );

      return {
        content: [
          {
            type: "text",
            text: result.success
              ? `Validation passed.\n\n${result.stdout}`
              : `Validation failed (exit ${result.exitCode}).\n\n${result.stdout || result.stderr}`,
          },
        ],
      };
    }
  );

  server.tool(
    "provision_agent",
    "Provision a declarative agent to a target environment. Creates the required " +
      "cloud resources and registers the agent with the M365 platform. Requires prior " +
      "authentication via 'wiqd auth login'.",
    {
      project_path: z
        .string()
        .describe("Absolute path to the declarative agent project root"),
      environment: z
        .string()
        .optional()
        .describe(
          "Target environment name (e.g., 'dev', 'local'). Defaults to the project's default env."
        ),
    },
    async ({ project_path, environment }) => {
      const args = ["agent", "provision", "--path", project_path, "--json"];
      if (environment) {
        args.push("--env", environment);
      }

      const result = await runWiqd(args, project_path);

      return {
        content: [
          {
            type: "text",
            text: result.success
              ? `Agent provisioned successfully.\n\n${result.stdout}`
              : `Provisioning failed (exit ${result.exitCode}).\n\n${result.stdout || result.stderr}`,
          },
        ],
      };
    }
  );

  server.tool(
    "publish_agent",
    "Publish a provisioned declarative agent to the M365 app catalog so users can " +
      "install and interact with it. The agent must be provisioned first.",
    {
      project_path: z
        .string()
        .describe("Absolute path to the declarative agent project root"),
      environment: z
        .string()
        .optional()
        .describe("Target environment name. Defaults to the project's default env."),
    },
    async ({ project_path, environment }) => {
      const args = ["agent", "publish", "--path", project_path, "--json"];
      if (environment) {
        args.push("--env", environment);
      }

      const result = await runWiqd(args, project_path);

      return {
        content: [
          {
            type: "text",
            text: result.success
              ? `Agent published successfully.\n\n${result.stdout}`
              : `Publishing failed (exit ${result.exitCode}).\n\n${result.stdout || result.stderr}`,
          },
        ],
      };
    }
  );

  server.tool(
    "agent_lifecycle",
    "Run the full declarative agent lifecycle: validate, provision, publish. " +
      "Stops at the first failure. Use this for a complete end-to-end deployment.",
    {
      project_path: z
        .string()
        .describe("Absolute path to the declarative agent project root"),
      environment: z
        .string()
        .optional()
        .describe("Target environment name. Defaults to the project's default env."),
    },
    async ({ project_path, environment }) => {
      const steps = ["validate", "provision", "publish"];
      const results: string[] = [];

      for (const step of steps) {
        const args = ["agent", step, "--path", project_path, "--json"];
        if (environment && step !== "validate") {
          args.push("--env", environment);
        }

        const result = await runWiqd(args, project_path);
        results.push(`[${step}] ${result.success ? "PASS" : "FAIL"} (exit ${result.exitCode})`);

        if (!result.success) {
          results.push(`\nStopped at ${step}. Error:\n${result.stdout || result.stderr}`);
          return {
            content: [{ type: "text" as const, text: results.join("\n") }],
          };
        }
      }

      results.push("\nFull lifecycle completed successfully!");
      return {
        content: [{ type: "text" as const, text: results.join("\n") }],
      };
    }
  );

  return server;
}
