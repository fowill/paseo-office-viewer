import type { PluginHandlerContext, PluginServerContext } from "@getpaseo/plugin/server";
import { createPreviewServer } from "./server/http";
import { listPreviewDirectory, openPreview, selectPreviewFile } from "./shared/preview";

async function directoryOf(paseo: PluginHandlerContext['paseo'], workspaceId: string) {
  // Resolve the directory on the daemon; clients submit only a workspace ID.
  const workspace = paseo.workspaces.ref(workspaceId);
  await workspace.refresh();
  const directory = workspace.directory;
  if (!directory) throw new Error("The workspace is unavailable. Reopen it and try again.");
  return directory;
}

export default function contribute(server: PluginServerContext) {
  let preview: ReturnType<typeof createPreviewServer> | undefined;

  server.handle(openPreview, async ({ workspaceId, theme }, { paseo }) => {
    const directory = await directoryOf(paseo, workspaceId);
    preview ??= createPreviewServer();
    const url = await (await preview).openWorkspace(directory, { embed: true, theme });
    return { url };
  });

  server.handle(listPreviewDirectory, async ({ workspaceId, path }, { paseo }) => {
    const directory = await directoryOf(paseo, workspaceId);
    preview ??= createPreviewServer();
    return (await preview).list(directory, path);
  });

  server.handle(selectPreviewFile, async ({ workspaceId, path }, { paseo }) => {
    const directory = await directoryOf(paseo, workspaceId);
    preview ??= createPreviewServer();
    return (await preview).select(directory, path);
  });

  return async () => {
    if (preview) await (await preview).close();
  };
}
