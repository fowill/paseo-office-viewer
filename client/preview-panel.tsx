import { type PluginWorkspacePanelProps, useRpc, useWorkspace } from "@getpaseo/plugin/client";
import { Icon } from "@getpaseo/plugin/client/react-native";
import { useCallback, useEffect, useRef, useState } from "react";
import { Pressable, Text, View } from "react-native";
import { listPreviewDirectory, openPreview, parsePreviewTheme, selectPreviewFile } from "../shared/preview";
import { FileTree } from "./file-tree";

const openingUntil = new Map<string, number>();
const desktopMessage = "预览需要在运行 Paseo 主机的桌面版中打开";

export function PreviewPanel(props: PluginWorkspacePanelProps) {
  return <WorkspacePreviewPanel key={props.workspaceId} {...props} />;
}

function WorkspacePreviewPanel({ theme, layout, workspaceId, navigation }: PluginWorkspacePanelProps) {
  const workspace = useWorkspace(workspaceId, ({ name }) => ({ name }));
  const requestPreview = useRpc(openPreview);
  const requestDirectory = useRpc(listPreviewDirectory);
  const requestSelection = useRpc(selectPreviewFile);
  const [refresh, setRefresh] = useState(0);
  const [busyPath, setBusyPath] = useState("");
  const [message, setMessage] = useState("");
  const [error, setError] = useState("");
  const [treeStatus, setTreeStatus] = useState({ loading: false, error: "" });
  const [hoveredButton, setHoveredButton] = useState("");
  const selectionQueue = useRef(Promise.resolve());
  const selectionSequence = useRef(0);
  const mounted = useRef(true);
  const openBrowser = navigation?.openBrowser;
  const colors = theme.colors;
  const colorsKey = JSON.stringify(colors);

  useEffect(() => {
    mounted.current = true;
    return () => { mounted.current = false; };
  }, []);

  useEffect(() => {
    if (!workspace) return;
    let active = true;
    const previewTheme = parsePreviewTheme(colors);
    void requestPreview({ workspaceId, ...(previewTheme ? { theme: previewTheme } : {}) }).catch((cause) => {
      if (active) setError(cause instanceof Error ? cause.message : String(cause));
    });
    return () => { active = false; };
  }, [workspaceId, Boolean(workspace), colorsKey, requestPreview]);

  const readDirectory = useCallback(async (path: string) => {
    const { entries } = await requestDirectory({ workspaceId, path });
    return entries;
  }, [requestDirectory, workspaceId]);

  function selectFile(path: string) {
    const sequence = ++selectionSequence.current;
    setBusyPath(path);
    setMessage("");
    setError("");
    // Preserve click order at the server, even if an earlier file check is slow.
    selectionQueue.current = selectionQueue.current.then(async () => {
      if (sequence !== selectionSequence.current || !mounted.current) return;
      try {
        const result = await requestSelection({ workspaceId, path });
        if (sequence !== selectionSequence.current || !mounted.current) return;
        if (!openBrowser) setMessage(desktopMessage);
        else if (result.viewerConnected) openingUntil.delete(workspaceId);
        else if ((openingUntil.get(workspaceId) ?? 0) < Date.now()) {
          // The first browser needs time to establish its SSE connection.
          openingUntil.set(workspaceId, Date.now() + 10_000);
          openBrowser({ url: result.url, workspaceId });
        }
      } catch (cause) {
        if (sequence === selectionSequence.current && mounted.current) {
          openingUntil.delete(workspaceId);
          setError(cause instanceof Error ? cause.message : String(cause));
        }
      } finally {
        if (sequence === selectionSequence.current && mounted.current) setBusyPath("");
      }
    });
  }

  async function showPreview() {
    setError("");
    if (!openBrowser) {
      setMessage(desktopMessage);
      return;
    }
    try {
      const previewTheme = parsePreviewTheme(colors);
      const { url } = await requestPreview({ workspaceId, ...(previewTheme ? { theme: previewTheme } : {}) });
      if (!mounted.current) return;
      openingUntil.set(workspaceId, Date.now() + 10_000);
      openBrowser({ url, workspaceId });
    } catch (cause) {
      openingUntil.delete(workspaceId);
      if (mounted.current) setError(cause instanceof Error ? cause.message : String(cause));
    }
  }

  if (!workspace) {
    return <Text style={{ padding: 8, color: colors.foregroundMuted, fontSize: 13 }}>工作区不可用，请重新打开</Text>;
  }

  const statusError = error || treeStatus.error;
  const status = statusError || message || (treeStatus.loading ? "正在读取…" : "");
  return (
    <View style={{ flex: 1, backgroundColor: colors.surface0 }}>
      <View style={{ height: layout.compact ? 36 : 28, paddingHorizontal: 8, flexDirection: "row", alignItems: "center", borderBottomColor: colors.border, borderBottomWidth: 1 }}>
        <Text numberOfLines={1} style={{ flex: 1, color: colors.foreground, fontSize: 12, fontWeight: "600" }}>{workspace.name}</Text>
        <Pressable
          accessibilityRole="button"
          accessibilityLabel="刷新文件树"
          onHoverIn={() => setHoveredButton("refresh")}
          onHoverOut={() => setHoveredButton("")}
          onPress={() => { setError(""); setMessage(""); setRefresh((value) => value + 1); }}
          style={{ width: 22, height: 22, alignItems: "center", justifyContent: "center", backgroundColor: hoveredButton === "refresh" ? colors.surface1 : colors.surface0 }}
        ><Icon name="RefreshCw" size={14} color={colors.foregroundMuted} /></Pressable>
        <Pressable
          accessibilityRole="button"
          accessibilityLabel="打开预览"
          onHoverIn={() => setHoveredButton("open")}
          onHoverOut={() => setHoveredButton("")}
          onPress={() => void showPreview()}
          style={{ width: 22, height: 22, alignItems: "center", justifyContent: "center", backgroundColor: hoveredButton === "open" ? colors.surface1 : colors.surface0 }}
        ><Icon name="AppWindow" size={14} color={colors.foregroundMuted} /></Pressable>
      </View>
      <FileTree
        workspaceId={workspaceId}
        theme={theme}
        compact={layout.compact}
        readDirectory={readDirectory}
        onSelect={selectFile}
        onStatus={setTreeStatus}
        busyPath={busyPath}
        refresh={refresh}
      />
      {status ? <Text accessibilityRole={statusError ? "alert" : undefined} numberOfLines={1} style={{ height: layout.compact ? 24 : 20, paddingHorizontal: 8, lineHeight: layout.compact ? 24 : 20, fontSize: 11, color: statusError ? colors.statusDanger : colors.foregroundMuted }}>{status}</Text> : null}
    </View>
  );
}
