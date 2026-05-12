export function normalizeTags(input: unknown) {
  if (!Array.isArray(input)) {
    return [];
  }
  const seen = new Set<string>();
  const tags: string[] = [];
  for (const value of input) {
    if (typeof value !== "string") {
      continue;
    }
    const tag = value.trim();
    if (!tag || seen.has(tag)) {
      continue;
    }
    seen.add(tag);
    tags.push(tag);
  }
  return tags;
}

/**
 * TeamEdit stores tags as top-level document metadata so Haven views and app
 * pickers can categorize documents without reading markdown-specific fields.
 */
export function readTags(data: Record<string, unknown> | null | undefined) {
  return normalizeTags(data?.tags);
}
