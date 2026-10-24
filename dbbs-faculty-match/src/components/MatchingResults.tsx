import type { SubmissionResponse, FacultyDatasetStatus } from "../types";
import { formatSimilarity, formatStudentRank } from "../lib/formatting";

interface MatchingResultsProps {
  result: SubmissionResponse | null;
  datasetStatus: FacultyDatasetStatus | null;
  handleDownloadDirectoryResults: () => Promise<void>;
  handleDownloadSpreadsheetResults: () => Promise<void>;
}
export function MatchingResults({
  result,
  handleDownloadDirectoryResults,
  handleDownloadSpreadsheetResults,
  datasetStatus,
}: MatchingResultsProps) {
  if (!result) return null;
  const directorySummaryText = result?.directoryResults
    ? (() => {
        const { processedDocuments, matchedDocuments, skippedDocuments } =
          result.directoryResults;
        const sentences: string[] = [];
        sentences.push(
          `Processed ${processedDocuments} document${
            processedDocuments === 1 ? "" : "s"
          }`,
        );
        sentences.push(
          matchedDocuments === 0
            ? "No matches were produced"
            : `${matchedDocuments} document${
                matchedDocuments === 1 ? "" : "s"
              } produced matches`,
        );
        if (skippedDocuments > 0) {
          sentences.push(
            `${skippedDocuments} document${
              skippedDocuments === 1 ? " was" : "s were"
            } skipped due to missing content or errors`,
          );
        }
        return `${sentences.join(". ")}.`;
      })()
    : null;

  const spreadsheetSummaryText = result?.spreadsheetResults
    ? (() => {
        const { processedRows, matchedRows, skippedRows } =
          result.spreadsheetResults;
        const sentences: string[] = [];
        sentences.push(
          `Processed ${processedRows} row${processedRows === 1 ? "" : "s"}`,
        );
        sentences.push(
          matchedRows === 0
            ? "No matches were produced"
            : `${matchedRows} row${matchedRows === 1 ? "" : "s"} produced matches`,
        );
        if (skippedRows > 0) {
          sentences.push(
            `${skippedRows} row${skippedRows === 1 ? " was" : "s were"} skipped due to missing content or errors`,
          );
        }
        return `${sentences.join(". ")}.`;
      })()
    : null;
  return (
    <section className="result-card">
      <h2>Submission ready</h2>
      <p>{result.summary}</p>

      {result.warnings.length > 0 && (
        <ul className="warning-list">
          {result.warnings.map((warning, index) => (
            <li key={index}>{warning}</li>
          ))}
        </ul>
      )}

      <div className="detail-grid">
        <div className="detail-card">
          <h3>Input configuration</h3>
          <dl>
            <dt>Type</dt>
            <dd>{result.details.taskType}</dd>
            {result.details.promptPreview && (
              <>
                <dt>Prompt preview</dt>
                <dd>
                  <pre className="prompt-preview">
                    {result.details.promptPreview}
                  </pre>
                </dd>
              </>
            )}
            {result.details.validatedPaths.length > 0 && (
              <>
                <dt>Validated paths</dt>
                <dd>
                  <ul className="path-list">
                    {result.details.validatedPaths.map((path) => (
                      <li key={`${path.label}-${path.path}`}>
                        <strong>{path.label}:</strong> {path.path}
                      </li>
                    ))}
                  </ul>
                </dd>
              </>
            )}
            {result.details.spreadsheetIdentifierColumns.length > 0 && (
              <>
                <dt>Identifier columns</dt>
                <dd>
                  <ul className="path-list">
                    {result.details.spreadsheetIdentifierColumns.map(
                      (column) => (
                        <li key={`identifier-${column}`}>{column}</li>
                      ),
                    )}
                  </ul>
                </dd>
              </>
            )}
            {result.details.spreadsheetPromptColumns.length > 0 && (
              <>
                <dt>Prompt columns</dt>
                <dd>
                  <ul className="path-list">
                    {result.details.spreadsheetPromptColumns.map((column) => (
                      <li key={`prompt-${column}`}>{column}</li>
                    ))}
                  </ul>
                </dd>
              </>
            )}
          </dl>
        </div>

        <div className="detail-card">
          <h3>Faculty scope</h3>
          <dl>
            <dt>Scope</dt>
            <dd>{result.details.facultyScope}</dd>
            {result.details.programFilters.length > 0 && (
              <>
                <dt>Programs</dt>
                <dd>
                  <ul className="path-list">
                    {result.details.programFilters.map((program) => (
                      <li key={program}>{program}</li>
                    ))}
                  </ul>
                </dd>
              </>
            )}
            {result.details.customFacultyPath && (
              <>
                <dt>Faculty roster</dt>
                <dd>
                  <div className="path-preview">
                    {result.details.customFacultyPath}
                  </div>
                </dd>
              </>
            )}
            {Object.keys(result.details.facultyRosterColumnMap).length > 0 && (
              <>
                <dt>Roster mapping</dt>
                <dd>
                  <ul className="path-list">
                    {Object.entries(result.details.facultyRosterColumnMap).map(
                      ([identifier, column]) => (
                        <li key={`roster-map-${identifier}-${column}`}>
                          <strong>{identifier}:</strong> {column}
                        </li>
                      ),
                    )}
                  </ul>
                </dd>
              </>
            )}
            {result.details.facultyRosterWarnings.length > 0 && (
              <>
                <dt>Roster warnings</dt>
                <dd>
                  <ul className="warning-list">
                    {result.details.facultyRosterWarnings.map(
                      (warning, index) => (
                        <li key={`roster-warning-detail-${index}`}>
                          {warning}
                        </li>
                      ),
                    )}
                  </ul>
                </dd>
              </>
            )}
          </dl>
        </div>

        <div className="detail-card">
          <h3>Recommendation limits</h3>
          <dl>
            <dt>Faculty per student</dt>
            <dd>{result.details.recommendationsPerStudent}</dd>
          </dl>
        </div>
      </div>

      {result.directoryResults && (
        <section className="directory-results">
          <h3>Directory match preview</h3>
          {directorySummaryText && (
            <p className="section-description">{directorySummaryText}</p>
          )}
          {result.directoryResults.preview.rows.length > 0 ? (
            <>
              <div className="spreadsheet-preview-card">
                <div className="spreadsheet-preview">
                  <table>
                    <thead>
                      <tr>
                        {result.directoryResults.preview.headers.map(
                          (header, headerIndex) => (
                            <th key={`directory-header-${headerIndex}`}>
                              {header}
                            </th>
                          ),
                        )}
                      </tr>
                    </thead>
                    <tbody>
                      {result.directoryResults.preview.rows.map(
                        (row, rowIndex) => (
                          <tr key={`directory-row-${rowIndex}`}>
                            {row.map((cell, cellIndex) => (
                              <td
                                key={`directory-cell-${rowIndex}-${cellIndex}`}
                              >
                                {cell}
                              </td>
                            ))}
                          </tr>
                        ),
                      )}
                    </tbody>
                  </table>
                </div>
              </div>
              <p className="small-note">
                Showing {result.directoryResults.preview.rows.length}
                {result.directoryResults.totalRows >
                result.directoryResults.preview.rows.length
                  ? ` of ${result.directoryResults.totalRows}`
                  : ""}{" "}
                row{result.directoryResults.totalRows === 1 ? "" : "s"}.
              </p>
            </>
          ) : (
            <p className="small-note">
              No match rows were generated for the available documents.
            </p>
          )}
          <div className="button-row">
            <button
              type="button"
              className="secondary"
              onClick={handleDownloadDirectoryResults}
            >
              Download match spreadsheet
            </button>
          </div>
        </section>
      )}

      {result.spreadsheetResults && (
        <section className="directory-results">
          <h3>Spreadsheet match preview</h3>
          {spreadsheetSummaryText && (
            <p className="section-description">{spreadsheetSummaryText}</p>
          )}
          {result.spreadsheetResults.preview.rows.length > 0 ? (
            <>
              <div className="spreadsheet-preview-card">
                <div className="spreadsheet-preview">
                  <table>
                    <thead>
                      <tr>
                        {result.spreadsheetResults.preview.headers.map(
                          (header, headerIndex) => (
                            <th key={`spreadsheet-header-${headerIndex}`}>
                              {header}
                            </th>
                          ),
                        )}
                      </tr>
                    </thead>
                    <tbody>
                      {result.spreadsheetResults.preview.rows.map(
                        (row, rowIndex) => (
                          <tr key={`spreadsheet-row-${rowIndex}`}>
                            {row.map((cell, cellIndex) => (
                              <td
                                key={`spreadsheet-cell-${rowIndex}-${cellIndex}`}
                              >
                                {cell}
                              </td>
                            ))}
                          </tr>
                        ),
                      )}
                    </tbody>
                  </table>
                </div>
              </div>
              <p className="small-note">
                Showing {result.spreadsheetResults.preview.rows.length}
                {result.spreadsheetResults.totalRows >
                result.spreadsheetResults.preview.rows.length
                  ? ` of ${result.spreadsheetResults.totalRows}`
                  : ""}{" "}
                row{result.spreadsheetResults.totalRows === 1 ? "" : "s"}.
              </p>
            </>
          ) : (
            <p className="small-note">
              No match rows were generated for the available prompts.
            </p>
          )}
          <div className="button-row">
            <button
              type="button"
              className="secondary"
              onClick={handleDownloadSpreadsheetResults}
            >
              Download match spreadsheet
            </button>
          </div>
        </section>
      )}

      {result.promptMatches.length > 0 &&
        !result.directoryResults &&
        !result.spreadsheetResults && (
          <section className="match-results">
            {result.promptMatches.map((match, matchIndex) => (
              <article className="match-card" key={`match-${matchIndex}`}>
                <header className="match-card-header">
                  <h3>Faculty matches for prompt</h3>
                  <pre className="match-prompt">{match.prompt}</pre>
                </header>
                {match.facultyMatches.length > 0 ? (
                  <ol className="match-list">
                    {match.facultyMatches.map((faculty, rank) => {
                      const identifierEntries = Object.entries(
                        faculty.identifiers,
                      );
                      const identifierOrder =
                        datasetStatus?.analysis?.identifierColumns ?? [];
                      const sortedIdentifierEntries = identifierEntries
                        .slice()
                        .sort((a, b) => {
                          const indexA = identifierOrder.indexOf(a[0]);
                          const indexB = identifierOrder.indexOf(b[0]);

                          if (
                            indexA !== -1 &&
                            indexB !== -1 &&
                            indexA !== indexB
                          ) {
                            return indexA - indexB;
                          }

                          if (indexA !== -1 && indexB === -1) {
                            return -1;
                          }

                          if (indexA === -1 && indexB !== -1) {
                            return 1;
                          }

                          return a[0].localeCompare(b[0]);
                        });
                      const hasIdentifiers = sortedIdentifierEntries.length > 0;
                      const studentRankLabel = formatStudentRank(
                        faculty.studentRankForFaculty,
                        faculty.studentRankTotal,
                      );

                      return (
                        <li
                          className="match-list-item"
                          key={`${faculty.rowIndex}-${rank}`}
                        >
                          <div className="match-list-header">
                            <div className="match-rankings">
                              <span className="match-rank">
                                Faculty rank #{rank + 1}
                              </span>
                              {studentRankLabel && (
                                <span className="match-student-rank">
                                  Student rank {studentRankLabel}
                                </span>
                              )}
                            </div>
                            <span className="match-score">
                              Similarity: {formatSimilarity(faculty.similarity)}
                            </span>
                          </div>
                          <div className="match-identifiers">
                            {hasIdentifiers ? (
                              sortedIdentifierEntries.map(([label, value]) => (
                                <span
                                  key={`${faculty.rowIndex}-${label}`}
                                  className="match-identifier"
                                >
                                  <span className="match-identifier-label">
                                    {label}:
                                  </span>{" "}
                                  {value}
                                </span>
                              ))
                            ) : (
                              <span className="match-identifier">
                                <span className="match-identifier-label">
                                  Row index:
                                </span>{" "}
                                {faculty.rowIndex}
                              </span>
                            )}
                          </div>
                          {faculty.facultyText && (
                            <pre className="match-faculty-text">
                              {faculty.facultyText}
                            </pre>
                          )}
                        </li>
                      );
                    })}
                  </ol>
                ) : (
                  <p className="match-empty">
                    No faculty matches were returned for this prompt.
                  </p>
                )}
              </article>
            ))}
          </section>
        )}
    </section>
  );
}
