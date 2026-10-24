import { useEffect, useRef, useState } from "react";
import { isTauri } from "@tauri-apps/api/core";
import { getCurrentWindow } from "@tauri-apps/api/window";
import type { UnlistenFn } from "@tauri-apps/api/event";
import { backend } from "../services/backend";
import {
  isInsideDropZone,
  singleDroppedPath,
  validateUploadExtension,
} from "../lib/uploads";
import type { UploadKind } from "../types";

export function useNativeFileDrop(
  kind: UploadKind,
  onPath: (path: string) => void | Promise<void>,
  disabled: boolean,
) {
  const ref = useRef<HTMLDivElement>(null);
  const current = useRef({ onPath, disabled, kind });
  current.current = { onPath, disabled, kind };
  const [isOver, setIsOver] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [isValidating, setIsValidating] = useState(false);

  useEffect(() => {
    if (!isTauri()) return;
    let active = true;
    let pending = false;
    let unlisten: UnlistenFn | undefined;
    const bind = async () => {
      try {
        const stop = await getCurrentWindow().onDragDropEvent(
          async ({ payload }) => {
            if (!active) return;
            if (payload.type === "leave") {
              setIsOver(false);
              return;
            }
            const rect = ref.current?.getBoundingClientRect();
            const inside =
              !!rect &&
              isInsideDropZone(payload.position, rect, window.devicePixelRatio);
            const available = !current.current.disabled && !pending;
            setIsOver(inside && available && payload.type !== "drop");
            if (payload.type !== "drop" || !inside || !available) return;

            pending = true;
            setIsValidating(true);
            setError(null);
            try {
              const path = singleDroppedPath(payload.paths);
              const targetKind = current.current.kind;
              validateUploadExtension(path, targetKind);
              const validated = await backend.validateUploadPath({
                path,
                kind: targetKind,
              });
              if (active && !current.current.disabled)
                await current.current.onPath(validated);
            } catch (failure) {
              if (active)
                setError(
                  failure instanceof Error ? failure.message : String(failure),
                );
            } finally {
              pending = false;
              if (active) setIsValidating(false);
            }
          },
        );
        if (active) unlisten = stop;
        else stop();
      } catch (failure) {
        if (active)
          setError(
            `Unable to enable drag and drop: ${String(failure)}. Use Browse to select an item.`,
          );
      }
    };
    void bind();
    return () => {
      active = false;
      unlisten?.();
    };
  }, []);

  useEffect(() => {
    if (disabled) setIsOver(false);
  }, [disabled]);
  return { ref, isOver, error, isValidating };
}
