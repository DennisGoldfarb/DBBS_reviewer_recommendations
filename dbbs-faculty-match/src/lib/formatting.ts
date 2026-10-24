import type { EmbeddingProgressPayload } from "../types";

export const formatDuration = (valueInSeconds: number): string => {
  if (!Number.isFinite(valueInSeconds) || valueInSeconds < 0) {
    return "unknown";
  }

  const totalSeconds = Math.round(valueInSeconds);
  const hours = Math.floor(totalSeconds / 3600);
  const minutes = Math.floor((totalSeconds % 3600) / 60);
  const seconds = totalSeconds % 60;
  const parts: string[] = [];

  if (hours > 0) {
    parts.push(`${hours}h`);
  }

  if (minutes > 0 || hours > 0) {
    parts.push(`${minutes}m`);
  }

  parts.push(`${seconds}s`);
  return parts.join(" ");
};

export const describeEmbeddingProgress = (
  progress: EmbeddingProgressPayload,
): string => {
  const pieces: string[] = [];
  const baseMessage = progress.message?.trim();

  if (baseMessage) {
    pieces.push(baseMessage);
  } else if (progress.totalRows > 0) {
    const percent = Math.min(
      100,
      Math.max(
        0,
        Math.floor((progress.processedRows / progress.totalRows) * 100),
      ),
    );
    pieces.push(
      `Processed ${progress.processedRows} of ${progress.totalRows} faculty rows (${percent}%)`,
    );
  }

  if (progress.totalRows > 0) {
    const percent = Math.min(
      100,
      Math.max(
        0,
        Math.floor((progress.processedRows / progress.totalRows) * 100),
      ),
    );
    const summary = `${progress.processedRows} of ${progress.totalRows} rows (${percent}%)`;
    if (!baseMessage || !baseMessage.includes(`${progress.totalRows}`)) {
      pieces.push(summary);
    }
  }

  const remaining = progress.estimatedRemainingSeconds;
  if (
    typeof remaining === "number" &&
    Number.isFinite(remaining) &&
    remaining > 0
  ) {
    pieces.push(`ETA: ${formatDuration(remaining)}`);
  } else if (
    progress.phase === "complete" &&
    typeof progress.elapsedSeconds === "number" &&
    Number.isFinite(progress.elapsedSeconds) &&
    progress.elapsedSeconds > 0
  ) {
    pieces.push(`Completed in ${formatDuration(progress.elapsedSeconds)}`);
  }

  return pieces.filter(Boolean).join(" – ");
};

export const formatSimilarity = (value: number): string => {
  if (!Number.isFinite(value)) {
    return "n/a";
  }

  const percent = (value * 100).toFixed(1);
  return `${percent}%`;
};

export const formatStudentRank = (
  rank?: number,
  total?: number,
): string | null => {
  if (typeof rank !== "number") {
    return null;
  }

  if (typeof total === "number" && Number.isFinite(total) && total > 0) {
    return `#${rank} of ${total}`;
  }

  return `#${rank}`;
};

export const formatDatasetTimestamp = (value: string | null) => {
  if (!value) {
    return "Not available";
  }
  const parsed = new Date(value);
  if (Number.isNaN(parsed.getTime())) {
    return value;
  }
  return parsed.toLocaleString();
};
