export function columnLabel(
  headers: readonly string[] | undefined,
  index: number,
): string {
  return headers?.[index]?.trim() || `Column ${index + 1}`;
}

export function toggleColumn(selected: number[], index: number): number[] {
  return (
    selected.includes(index)
      ? selected.filter((value) => value !== index)
      : [...selected, index]
  ).sort((a, b) => a - b);
}
