import type { PluginServerContext } from "@getpaseo/plugin/server";
import { createPreviewServer } from "./server/http";
import { openPreview } from "./shared/preview";

export default function contribute(server: PluginServerContext) {
  let preview: ReturnType<typeof createPreviewServer> | undefined;

  server.handle(openPreview, async ({ workspaceId }, { paseo }) => {
    // Resolve the directory on the daemon; clients submit only a workspace ID.
    const workspace = paseo.workspaces.ref(workspaceId);
    await workspace.refresh();
    const directory = workspace.directory;
    if (!directory) throw new Error("The workspace is unavailable. Reopen it and try again.");

    preview ??= createPreviewServer();
    const url = await (await preview).openWorkspace(directory);
    return { url };
  });

  return async () => {
    if (preview) await (await preview).close();
  };
}
