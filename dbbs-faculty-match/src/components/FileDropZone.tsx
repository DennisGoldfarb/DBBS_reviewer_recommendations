import type { ReactNode } from "react";
import type { UploadKind } from "../types";
import { useNativeFileDrop } from "../hooks/useNativeFileDrop";

interface FileDropZoneProps {
  kind: UploadKind;
  label: string;
  onPath: (path: string) => void | Promise<void>;
  disabled?: boolean;
  children?: ReactNode;
  className?: string;
}

export function FileDropZone({
  kind,
  label,
  onPath,
  disabled = false,
  children,
  className = "",
}: FileDropZoneProps) {
  const drop = useNativeFileDrop(kind, onPath, disabled);
  return (
    <div
      ref={drop.ref}
      role="group"
      aria-label={label}
      aria-disabled={disabled}
      className={`file-drop-zone ${drop.isOver ? "file-drop-zone-active" : ""} ${className}`}
      onDragOver={(event) => event.preventDefault()}
      onDrop={(event) => event.preventDefault()}
    >
      <p className="file-drop-hint" role="status">
        {drop.isValidating
          ? "Checking dropped item…"
          : `Drop ${kind === "directory" ? "a folder" : "a file"} here, or use Browse.`}
      </p>
      {children}
      {drop.error && (
        <p className="preview-error" role="alert">
          {drop.error}
        </p>
      )}
    </div>
  );
}
