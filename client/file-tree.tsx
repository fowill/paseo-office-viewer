import type { PluginWorkspacePanelProps } from "@getpaseo/plugin/client";
import { FlatList, Icon } from "@getpaseo/plugin/client/react-native";
import { useCallback, useEffect, useRef, useState } from "react";
import { ActivityIndicator, type FlatList as NativeFlatList, Pressable, Text, View } from "react-native";
import type { PreviewEntry } from "../shared/preview";

type TreeMemory = {
  expanded: Set<string>;
  children: Map<string, PreviewEntry[]>;
  selected: string;
  scrollOffset: number;
};
type TreeRow = { key: string; depth: number; entry?: PreviewEntry; parent?: string; error?: boolean };
const trees = new Map<string, TreeMemory>();
const fileIcons: Record<string, string> = { pdf: "FileText", docx: "FileType", xlsx: "FileSpreadsheet", pptx: "Presentation" };

type FileTreeProps = {
  workspaceId: string;
  theme: PluginWorkspacePanelProps["theme"];
  compact: boolean;
  readDirectory(path: string): Promise<PreviewEntry[]>;
  onSelect(path: string): void;
  onStatus(status: { loading: boolean; error: string }): void;
  busyPath: string;
  refresh: number;
};

export function FileTree({ workspaceId, theme, compact, readDirectory, onSelect, onStatus, busyPath, refresh }: FileTreeProps) {
  const [memory] = useState(() => {
    let value = trees.get(workspaceId);
    if (!value) {
      value = { expanded: new Set(), children: new Map(), selected: "", scrollOffset: 0 };
      trees.set(workspaceId, value);
    }
    return value;
  });
  const [, render] = useState(0);
  const [hovered, setHovered] = useState("");
  const pending = useRef(new Set<string>());
  const errors = useRef(new Map<string, string>());
  const generation = useRef(0);
  const active = useRef(true);
  const list = useRef<NativeFlatList<TreeRow>>(null);
  const restoreScroll = useRef(true);
  const colors = theme.colors;
  const rowHeight = compact ? 32 : 24;

  const update = useCallback(() => {
    if (!active.current) return;
    render((value) => value + 1);
    onStatus({ loading: pending.current.size > 0, error: errors.current.get("") ?? "" });
  }, [onStatus]);

  const load = useCallback(async (path: string) => {
    if (pending.current.has(path)) return;
    const requestGeneration = generation.current;
    pending.current.add(path);
    errors.current.delete(path);
    update();
    try {
      const entries = await readDirectory(path);
      if (generation.current === requestGeneration) memory.children.set(path, entries);
    } catch (cause) {
      if (generation.current === requestGeneration) errors.current.set(path, cause instanceof Error ? cause.message : String(cause));
    } finally {
      if (generation.current === requestGeneration) {
        pending.current.delete(path);
        update();
      }
    }
  }, [memory, readDirectory, update]);

  useEffect(() => {
    active.current = true;
    if (!memory.children.has("")) void load("");
    return () => { active.current = false; generation.current++; pending.current.clear(); };
  }, [load, memory]);

  useEffect(() => {
    if (!refresh) return;
    generation.current++;
    pending.current.clear();
    errors.current.clear();
    restoreScroll.current = true;
    // Keep visible rows in place while their directory snapshots are replaced.
    const visibleDirectories = new Set(["", ...memory.expanded]);
    for (const path of memory.children.keys()) {
      if (!visibleDirectories.has(path)) memory.children.delete(path);
    }
    void Promise.all([...visibleDirectories].map(load)).then(() => {
      if (active.current) list.current?.scrollToOffset({ offset: memory.scrollOffset, animated: false });
    });
  }, [refresh, load, memory]);

  const rows: TreeRow[] = [];
  function appendDirectory(path: string, depth: number) {
    const entries = memory.children.get(path);
    if (errors.current.has(path)) {
      rows.push({ key: path + "/:error", parent: path, depth, error: true });
      return;
    }
    if (entries?.length === 0 && path) rows.push({ key: path + "/:empty", parent: path, depth });
    for (const entry of entries ?? []) {
      rows.push({ key: entry.path, entry, depth });
      if (entry.directory && memory.expanded.has(entry.path)) appendDirectory(entry.path, depth + 1);
    }
  }
  appendDirectory("", 0);

  function press(entry: PreviewEntry) {
    if (entry.directory) {
      if (memory.expanded.has(entry.path)) memory.expanded.delete(entry.path);
      else {
        memory.expanded.add(entry.path);
        if (!memory.children.has(entry.path) || errors.current.has(entry.path)) void load(entry.path);
      }
    } else {
      memory.selected = entry.path;
      onSelect(entry.path);
    }
    update();
  }

  return (
    <FlatList
      ref={list}
      accessibilityLabel="Office Viewer 文件树"
      style={{ flex: 1 }}
      data={rows}
      keyExtractor={(row) => row.key}
      getItemLayout={(_, index) => ({ length: rowHeight, offset: rowHeight * index, index })}
      contentOffset={{ x: 0, y: memory.scrollOffset }}
      onScroll={(event) => { memory.scrollOffset = event.nativeEvent.contentOffset.y; }}
      scrollEventThrottle={16}
      onContentSizeChange={() => {
        if (restoreScroll.current && rows.length) {
          restoreScroll.current = false;
          list.current?.scrollToOffset({ offset: memory.scrollOffset, animated: false });
        }
      }}
      ListEmptyComponent={!pending.current.has("") && memory.children.has("") ? <Text style={{ padding: 16, fontSize: 13, color: colors.foregroundMuted }}>没有可预览的 PDF、Word、Excel、PowerPoint 文件</Text> : null}
      renderItem={({ item }) => {
        const paddingLeft = (compact ? 12 : 8) + item.depth * (compact ? 16 : 14);
        if (!item.entry) {
          return <Pressable disabled={!item.error} onPress={() => void load(item.parent!)} accessibilityRole={item.error ? "button" : undefined} style={{ height: rowHeight, paddingLeft: paddingLeft + 22, justifyContent: "center" }}>
            <Text numberOfLines={1} style={{ color: item.error ? colors.statusDanger : colors.foregroundMuted, fontSize: compact ? 15 : 13 }}>{item.error ? "无法读取 · 重试" : "（空）"}</Text>
          </Pressable>;
        }
        const entry = item.entry;
        const expanded = memory.expanded.has(entry.path);
        const selected = memory.selected === entry.path;
        const icon = entry.directory ? (expanded ? "FolderOpen" : "Folder") : fileIcons[entry.name.split(".").pop()?.toLowerCase() ?? ""];
        return <Pressable
          accessibilityRole="button"
          accessibilityLabel={entry.name}
          accessibilityState={{ selected, expanded: entry.directory ? expanded : undefined, busy: pending.current.has(entry.path) || busyPath === entry.path }}
          onHoverIn={() => setHovered(entry.path)}
          onHoverOut={() => setHovered("")}
          onPress={() => press(entry)}
          style={{ height: rowHeight, paddingLeft, paddingRight: 8, flexDirection: "row", alignItems: "center", backgroundColor: selected ? colors.surface2 : hovered === entry.path ? colors.surface1 : colors.surface0 }}
        >
          <View style={{ width: 16, marginRight: 6 }}>{entry.directory ? <Icon name={expanded ? "ChevronDown" : "ChevronRight"} size={16} color={colors.foregroundMuted} /> : null}</View>
          <Icon name={icon} size={16} color={selected ? colors.accent : colors.foregroundMuted} />
          <Text numberOfLines={1} ellipsizeMode="middle" style={{ flex: 1, marginLeft: 6, color: colors.foreground, fontSize: compact ? 15 : 13, fontWeight: selected ? "500" : "400" }}>{entry.name}</Text>
          {pending.current.has(entry.path) || busyPath === entry.path ? <ActivityIndicator size={12} color={colors.foregroundMuted} style={{ marginLeft: 6 }} /> : null}
        </Pressable>;
      }}
    />
  );
}
