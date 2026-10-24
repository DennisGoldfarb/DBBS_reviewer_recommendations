import type { SpreadsheetPreview } from "../types";
import { columnLabel } from "../lib/columns";

export function PreviewTable({ preview }: { preview: SpreadsheetPreview }) {
  return (
    <table className="preview-table">
      <thead>
        <tr>
          {preview.headers.map((_, index) => (
            <th key={index}>{columnLabel(preview.headers, index)}</th>
          ))}
        </tr>
      </thead>
      <tbody>
        {preview.rows.map((row, rowIndex) => (
          <tr key={rowIndex}>
            {preview.headers.map((_, columnIndex) => (
              <td key={columnIndex} title={row[columnIndex] ?? ""}>
                {row[columnIndex] ?? ""}
              </td>
            ))}
          </tr>
        ))}
      </tbody>
    </table>
  );
}
