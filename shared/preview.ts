import { defineRpc } from "@getpaseo/plugin";
import { z } from "zod";

const color = z.string().max(40).regex(/^#(?:[0-9a-fA-F]{3,4}|[0-9a-fA-F]{6}|[0-9a-fA-F]{8})$|^(?:rgb|rgba|hsl|hsla)\([\d.,\s%/]+\)$/);

export const themeSchema = z.object({
  surface0: color,
  surface1: color,
  surface2: color,
  border: color,
  foreground: color,
  foregroundMuted: color,
  accent: color,
  accentForeground: color,
  statusSuccess: color,
  statusWarning: color,
  statusDanger: color,
}).strict();
export type PreviewTheme = z.infer<typeof themeSchema>;

export function parsePreviewTheme(value: unknown): PreviewTheme | undefined {
  const result = themeSchema.safeParse(value);
  return result.success ? result.data : undefined;
}

const relativePath = z.string().max(1024).refine(value => !value.includes("\0"));
export const entrySchema = z.object({
  name: z.string(),
  path: z.string(),
  directory: z.boolean(),
  size: z.number(),
  modified: z.number(),
}).strict();
export type PreviewEntry = z.infer<typeof entrySchema>;

export const openPreview = defineRpc({
  name: "preview.open",
  input: z.object({ workspaceId: z.string().min(1), theme: themeSchema.optional() }).strict(),
  output: z.object({ url: z.string().url() }).strict(),
});

export const listPreviewDirectory = defineRpc({
  name: "preview.list",
  input: z.object({ workspaceId: z.string().min(1), path: relativePath }).strict(),
  output: z.object({ entries: z.array(entrySchema) }).strict(),
});

export const selectPreviewFile = defineRpc({
  name: "preview.select",
  input: z.object({ workspaceId: z.string().min(1), path: relativePath }).strict(),
  output: z.object({ url: z.string().url(), viewerConnected: z.boolean() }).strict(),
});
