import type { Event as TauriEvent, UnlistenFn } from "@tauri-apps/api/event";

const getTauriGlobal = (): Record<string, unknown> | null => {
  if (typeof globalThis === "undefined") {
    return null;
  }

  const globalObject = globalThis as Record<string, unknown>;
  const internals = globalObject["__TAURI_INTERNALS__"];
  if (internals && typeof internals === "object") {
    return internals as Record<string, unknown>;
  }

  const legacy = globalObject["__TAURI__"];
  if (legacy && typeof legacy === "object") {
    return legacy as Record<string, unknown>;
  }

  return null;
};

const resolveFunction = <T extends (...args: any[]) => unknown>(
  value: unknown,
): T | null => {
  if (typeof value === "function") {
    return value as T;
  }
  return null;
};

type InvokeHandler = <T>(command: string, args?: Record<string, unknown>) => Promise<T>;

let invokeLoader: Promise<InvokeHandler | null> | null = null;

const loadInvoke = async () => {
  if (!invokeLoader) {
    invokeLoader = (async () => {
      try {
        const module = await import("@tauri-apps/api/core");
        if (typeof module.invoke === "function") {
          return module.invoke as InvokeHandler;
        }
      } catch {
        /* fall back to globals */
      }

      const tauri = getTauriGlobal();
      if (tauri) {
        const direct = resolveFunction<InvokeHandler>((tauri as Record<string, unknown>).invoke);
        if (direct) {
          return direct;
        }
      }

      return null;
    })();
  }

  return invokeLoader;
};

export const safeInvoke = async <T>(command: string, args?: Record<string, unknown>): Promise<T> => {
  const loader = await loadInvoke();
  if (!loader) {
    throw new Error("The Tauri invoke API is unavailable.");
  }

  return loader<T>(command, args);
};

let eventModuleLoader: Promise<typeof import("@tauri-apps/api/event") | null> | null = null;

const loadEventModule = async (): Promise<typeof import("@tauri-apps/api/event") | null> => {
  if (!eventModuleLoader) {
    eventModuleLoader = import("@tauri-apps/api/event").catch(() => null);
  }

  return eventModuleLoader;
};

export const safeListen = async <TPayload>(
  event: string,
  handler: (event: TauriEvent<TPayload>) => void,
): Promise<UnlistenFn> => {
  const module = await loadEventModule();
  if (module && typeof module.listen === "function") {
    try {
      return await module.listen<TPayload>(event, handler);
    } catch {
      /* ignore and fall through */
    }
  }

  const tauri = getTauriGlobal();
  const eventApi = tauri?.event as Record<string, unknown> | undefined;
  const fallback = resolveFunction<
    (name: string, callback: (event: TauriEvent<TPayload>) => void) => Promise<UnlistenFn>
  >(eventApi?.listen);
  if (fallback) {
    try {
      return await fallback(event, handler);
    } catch {
      /* ignore */
    }
  }

  return () => {};
};

export interface FileDialogFilter {
  name: string;
  extensions: string[];
}

export interface FileDialogOptions {
  directory?: boolean;
  multiple?: boolean;
  defaultPath?: string;
  title?: string;
  filters?: FileDialogFilter[];
}

interface DialogModule {
  open?: (
    options: FileDialogOptions,
  ) => Promise<string | string[] | null>;
}

let dialogModuleLoader: Promise<DialogModule | null> | null = null;

const loadDialogModule = async (): Promise<DialogModule | null> => {
  if (!dialogModuleLoader) {
    dialogModuleLoader = import("@tauri-apps/plugin-dialog")
      .then((module) => module as unknown as DialogModule)
      .catch(() => null);
  }

  return dialogModuleLoader;
};

export const openFileDialog = async (
  options: FileDialogOptions,
): Promise<string | string[] | null> => {
  const module = await loadDialogModule();
  if (module?.open) {
    return module.open(options);
  }

  const tauri = getTauriGlobal();
  const dialogApi = tauri?.dialog as Record<string, unknown> | undefined;
  const fallback = resolveFunction<(
    opts: FileDialogOptions,
  ) => Promise<string | string[] | null>>(dialogApi?.open);
  if (fallback) {
    return fallback(options);
  }

  throw new Error("The Tauri file dialog API is unavailable.");
};
