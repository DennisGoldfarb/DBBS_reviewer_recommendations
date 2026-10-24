import { useCallback, useEffect, useState } from "react";
import { backend } from "../services/backend";
import { listen, type UnlistenFn } from "@tauri-apps/api/event";
import type {
  FacultyDatasetStatus,
  StatusMessage,
  EmbeddingProgressPayload,
} from "../types";

export function useEmbeddingHelper() {
  const [isUpdatingEmbeddings, setIsUpdatingEmbeddings] = useState(false);

  const [embeddingStatus, setEmbeddingStatus] = useState<StatusMessage | null>(
    null,
  );

  const [embeddingProgress, setEmbeddingProgress] =
    useState<EmbeddingProgressPayload | null>(null);

  const [isEmbeddingHelperReady, setIsEmbeddingHelperReady] = useState(false);

  const [initializationAttempt, setInitializationAttempt] = useState(0);
  const retryInitialization = () => {
    setEmbeddingStatus(null);
    setEmbeddingProgress(null);
    setIsEmbeddingHelperReady(false);
    setInitializationAttempt((value) => value + 1);
  };

  useEffect(() => {
    let isMounted = true;

    const warmup = async () => {
      try {
        const ready = await backend.ensureEmbeddingHelperReady();
        if (isMounted) {
          setIsEmbeddingHelperReady(ready);
        }
      } catch (warmupError) {
        if (!isMounted) {
          return;
        }
        const message =
          warmupError instanceof Error
            ? warmupError.message
            : String(warmupError);
        setIsEmbeddingHelperReady(false);
        setEmbeddingStatus({
          variant: "error",
          message: `Unable to initialize the embedding helper: ${message}`,
        });
      }
    };

    warmup();

    return () => {
      isMounted = false;
    };
  }, [initializationAttempt]);

  useEffect(() => {
    let isMounted = true;
    let unlisten: UnlistenFn | null = null;

    listen<EmbeddingProgressPayload>("faculty-embedding-progress", (event) => {
      setEmbeddingProgress(event.payload);
      if (event.payload.phase === "standby") {
        setIsEmbeddingHelperReady(true);
        setEmbeddingStatus((current) => {
          if (
            current?.variant === "error" &&
            current.message.startsWith(
              "Unable to initialize the embedding helper",
            )
          ) {
            return null;
          }
          return current;
        });
      }
    })
      .then((fn) => {
        if (!isMounted) {
          fn();
          return;
        }
        unlisten = fn;
      })
      .catch(() => {
        /* ignore progress binding errors */
      });

    return () => {
      isMounted = false;
      if (unlisten) {
        unlisten();
      }
    };
  }, []);

  const runEmbeddingRefresh = useCallback(
    async (statusForValidation: FacultyDatasetStatus | null) => {
      const currentStatus = statusForValidation;
      setEmbeddingProgress(null);
      if (!currentStatus || !currentStatus.isValid) {
        const message =
          currentStatus?.message ??
          "Provide a valid faculty dataset before refreshing embeddings.";
        setEmbeddingStatus({ variant: "error", message });
        return;
      }

      setIsUpdatingEmbeddings(true);
      setEmbeddingProgress({
        phase: "initializing",
        message: "Loading embedding program…",
        processedRows: 0,
        totalRows: 0,
      });
      setEmbeddingStatus(null);

      try {
        const message = await backend.updateFacultyEmbeddings();
        setEmbeddingStatus({ variant: "success", message });
      } catch (updateError) {
        const message =
          updateError instanceof Error
            ? updateError.message
            : String(updateError);
        setEmbeddingStatus({ variant: "error", message });
      } finally {
        setIsUpdatingEmbeddings(false);
        setEmbeddingProgress(null);
      }
    },
    [],
  );

  return {
    isUpdatingEmbeddings,
    embeddingStatus,
    embeddingProgress,
    isEmbeddingHelperReady,
    runEmbeddingRefresh,
    retryInitialization,
  };
}
