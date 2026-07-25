/**
 * Safe, user-facing API error messages (shared by client components).
 */

export const PUBLIC_ERROR_MESSAGES = {
  accessDenied: "You do not have access to this portal.",
  notFound: "This record could not be found.",
  serverError: "Something went wrong loading this data.",
} as const;

export function messageForHttpStatus(status: number, fallback?: string): string {
  if (status === 401 || status === 403) {
    return PUBLIC_ERROR_MESSAGES.accessDenied;
  }
  if (status === 404) {
    return PUBLIC_ERROR_MESSAGES.notFound;
  }
  if (status >= 500) {
    return PUBLIC_ERROR_MESSAGES.serverError;
  }
  return fallback?.trim() || PUBLIC_ERROR_MESSAGES.serverError;
}

export async function readPublicApiError(response: Response): Promise<string> {
  let serverMessage: string | undefined;
  try {
    const data = (await response.json()) as { error?: string };
    serverMessage = data.error;
  } catch {
    serverMessage = undefined;
  }

  return messageForHttpStatus(response.status, serverMessage);
}
