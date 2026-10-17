import { type PluginWorkspacePanelProps, useRpc, useWorkspace } from "@getpaseo/plugin/client";
import { useState } from "react";
import { Pressable, ScrollView, Text, View } from "react-native";
import { openPreview } from "../shared/preview";

export function PreviewPanel({ theme, layout, workspaceId, navigation }: PluginWorkspacePanelProps) {
  const workspace = useWorkspace(workspaceId, ({ name, directory }) => ({ name, directory }));
  const requestPreview = useRpc(openPreview);
  const [opening, setOpening] = useState(false);
  const [error, setError] = useState("");
  const openBrowser = navigation?.openBrowser;
  const colors = theme.colors;
  const disabled = opening || !workspace || !openBrowser;

  async function open() {
    if (disabled || !openBrowser) return;
    setOpening(true);
    setError("");
    try {
      const { url } = await requestPreview({ workspaceId });
      openBrowser({ url, workspaceId });
    } catch (cause) {
      setError(cause instanceof Error ? cause.message : String(cause));
    } finally {
      setOpening(false);
    }
  }

  return (
    <ScrollView
      style={{ flex: 1, backgroundColor: colors.surface0 }}
      contentContainerStyle={{ padding: layout.compact ? 20 : 32, gap: 24 }}
    >
      <View style={{ gap: 8 }}>
        <Text style={{ color: colors.foregroundMuted, fontSize: 12, letterSpacing: 2 }}>LOCAL DOCUMENTS · V1.0</Text>
        <Text style={{ color: colors.foreground, fontSize: 28, fontWeight: "600" }}>Office Viewer</Text>
        <Text style={{ color: colors.foregroundMuted, fontSize: 15, lineHeight: 23 }}>
          在 Paseo 中查看 Word、PDF、Excel 和 PowerPoint。
        </Text>
      </View>

      <View style={{ backgroundColor: colors.surface1, borderColor: colors.border, borderWidth: 1, borderRadius: 12, padding: 20, gap: 10 }}>
        <Text style={{ color: colors.foregroundMuted, fontSize: 12 }}>当前工作区</Text>
        <Text style={{ color: colors.foreground, fontSize: 17, fontWeight: "600" }}>{workspace?.name ?? "工作区不可用"}</Text>
        <Text selectable style={{ color: colors.foregroundMuted, fontSize: 13, lineHeight: 21 }}>{workspace?.directory ?? "请重新打开工作区。"}</Text>
      </View>

      <View style={{ gap: 12 }}>
        <Pressable
          accessibilityRole="button"
          accessibilityLabel="打开文档预览"
          accessibilityState={{ disabled, busy: opening }}
          disabled={disabled}
          onPress={() => void open()}
          style={{ alignSelf: "flex-start", backgroundColor: colors.accent, borderRadius: 8, paddingHorizontal: 22, paddingVertical: 13, opacity: disabled ? 0.5 : 1 }}
        >
          <Text style={{ color: colors.accentForeground, fontSize: 15, fontWeight: "600" }}>{opening ? "正在打开…" : "打开文档预览"}</Text>
        </Pressable>
        {error ? <Text accessibilityRole="alert" style={{ color: colors.statusDanger, lineHeight: 22 }}>{error}</Text> : null}
        <Text style={{ color: colors.foregroundMuted, fontSize: 13, lineHeight: 22 }}>
          {openBrowser
            ? "请在运行当前 Paseo 主机的这台电脑上使用。预览将在工作区浏览器中打开。"
            : "V1.0 预览需要 Paseo 桌面版。请在运行当前 Paseo 主机的电脑上打开此工作区。"}
        </Text>
      </View>

      <View style={{ gap: 8 }}>
        <Text style={{ color: colors.foreground, fontSize: 14 }}>支持 .pdf、.docx、.xlsx、.pptx</Text>
        <Text style={{ color: colors.foregroundMuted, fontSize: 13, lineHeight: 22 }}>
          浏览工作区文件夹，选择文件后预览。文件保持只读。复杂版式可能与 Microsoft Office 有差异。
        </Text>
      </View>
    </ScrollView>
  );
}
