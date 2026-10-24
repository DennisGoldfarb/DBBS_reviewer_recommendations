import { useCallback, useState } from "react";
import type { SubmissionResponse } from "../types";

export function useFeedback() {
  const [result, setResult] = useState<SubmissionResponse | null>(null);
  const [error, setError] = useState<string | null>(null);
  const clear = useCallback(() => {
    setError(null);
    setResult(null);
  }, []);
  return { result, setResult, error, setError, clear };
}
