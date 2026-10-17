import type { OwnedSubscription, PaseoWorkspaceListResult } from "@getpaseo/client";
import type { PluginButtonRegistration, PluginClientContext } from "@getpaseo/plugin/client";
import { PreviewPanel } from "./client/preview-panel";

export default function contribute(client: PluginClientContext) {
  const removePanel = client.addWorkspacePanel({
    id: "preview",
    title: "Office Viewer",
    icon: "FileSearch",
    context: "workspace",
    locations: ["explorer", "workspace"],
    Component: PreviewPanel,
  });
  const removeCommand = client.addCommandCenterItem({
    id: "open-preview",
    title: "Office Viewer · 文档预览",
    icon: "FileSearch",
    context: "workspace",
    keywords: ["preview", "word", "pdf", "excel", "powerpoint", "文档", "表格"],
    onSelect({ openPanel }) {
      openPanel("preview", { location: "explorer" });
    },
  });
  const buttons = new Map<string, PluginButtonRegistration>();
  let subscription: OwnedSubscription<PaseoWorkspaceListResult> | undefined;
  let unsubscribe: (() => void) | undefined;
  let disposed = false;

  function addButton(workspaceId: string) {
    if (disposed || buttons.has(workspaceId)) return;
    buttons.set(workspaceId, client.addHeaderButton({
      id: "open-" + workspaceId,
      workspaceId,
      button: {
        title: "Office Viewer",
        icon: "FileSearch",
        behavior: {
          kind: "action",
          onPress: () => client.openPanel("preview", { workspaceId, location: "explorer" }),
        },
      },
    }));
  }

  function syncButtons(snapshot: PaseoWorkspaceListResult) {
    const ids = new Set(snapshot.entries.map((workspace) => workspace.id));
    for (const [id, button] of buttons) {
      if (!ids.has(id)) {
        button.remove();
        buttons.delete(id);
      }
    }
    for (const id of ids) addButton(id);
  }

  void (async () => {
    try {
      const result = await client.paseo.workspaces.list({ subscribe: {} });
      subscription = result.subscription;
      if (disposed) {
        await subscription.release();
        return;
      }
      syncButtons(result);
      unsubscribe = subscription.subscribe({
        snapshot: syncButtons,
        update(message) {
          if (message.type !== "workspace_update" || disposed) return;
          const update = message.payload;
          if (update.kind === "upsert") addButton(update.workspace.id);
          else {
            buttons.get(update.id)?.remove();
            buttons.delete(update.id);
          }
        },
        error: (error) => console.warn("Office Viewer workspace subscription failed", error),
      });
    } catch (error) {
      console.warn("Office Viewer header buttons unavailable", error);
    }
  })();

  return async () => {
    disposed = true;
    unsubscribe?.();
    for (const button of buttons.values()) button.remove();
    buttons.clear();
    await Promise.all([removePanel(), removeCommand(), subscription?.release()]);
  };
}
