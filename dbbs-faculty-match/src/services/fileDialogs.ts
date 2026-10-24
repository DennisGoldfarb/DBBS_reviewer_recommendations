import { open, type OpenDialogOptions } from "@tauri-apps/plugin-dialog";
import { useCallback } from "react";

export type FileSelectionHandler = (
  setter: (value: string) => void,
  options: OpenDialogOptions,
) => Promise<string | null>;

export function useFileSelection(
  onError: (message: string) => void,
): FileSelectionHandler {
  return useCallback(
    async (setter, options) => {
      try {
        const selection = await open(options);
        if (typeof selection !== "string") return null;
        setter(selection);
        return selection;
      } catch (error) {
        onError(
          `Unable to open a file dialog: ${error instanceof Error ? error.message : String(error)}`,
        );
        return null;
      }
    },
    [onError],
  );
}
