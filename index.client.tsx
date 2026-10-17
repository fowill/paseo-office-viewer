import type { PluginClientContext } from "@getpaseo/plugin/client";
import { PreviewPanel } from "./client/preview-panel";

export default function contribute(client: PluginClientContext) {
  client.addWorkspacePanel({
    id: "preview",
    title: "Office Viewer",
    icon: "FileSearch",
    context: "workspace",
    locations: ["workspace", "explorer"],
    Component: PreviewPanel,
  });
  client.addCommandCenterItem({
    id: "open-preview",
    title: "Office Viewer · 文档预览",
    icon: "FileSearch",
    context: "workspace",
    keywords: ["preview", "word", "pdf", "excel", "powerpoint", "文档", "表格"],
    onSelect({ openPanel }) {
      openPanel("preview");
    },
  });
  return () => {};
}
