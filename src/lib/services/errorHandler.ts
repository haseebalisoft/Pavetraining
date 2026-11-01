import "server-only";

/**
 * Shared API error types and safe public messages.
 * Never expose SharePoint/Graph internals to clients.
 */

export class UnauthorizedError extends Error {
  constructor(message = "Authentication required") {
    super(message);
    this.name = "UnauthorizedError";
  }
}

export class AccessDeniedError extends Error {
  constructor(message = "Access denied") {
    super(message);
    this.name = "AccessDeniedError";
  }
}

export class NotFoundError extends Error {
  constructor(message = "Not found") {
    super(message);
    this.name = "NotFoundError";
  }
}

export class ValidationError extends Error {
  constructor(message: string) {
    super(message);
    this.name = "ValidationError";
  }
}

export const PUBLIC_ERROR_MESSAGES = {
  accessDenied: "You do not have access to this portal.",
  notFound: "This record could not be found.",
  serverError: "Something went wrong loading this data.",
} as const;

/**
 * For admin long-running jobs (matrix refresh): turn an unexpected failure
 * into a ValidationError whose message says what happened, instead of the
 * generic 500 text. Raw Graph messages stay in the server log; only our own
 * curated timeout text is passed through.
 */
export function toAdminJobError(error: unknown, jobLabel: string): Error {
  if (
    error instanceof UnauthorizedError ||
    error instanceof AccessDeniedError ||
    error instanceof NotFoundError ||
    error instanceof ValidationError
  ) {
    return error;
  }
  console.error(`[${jobLabel}] failed`, error);
  const { name, code, message } = (error ?? {}) as {
    name?: unknown;
    code?: unknown;
    message?: unknown;
  };
  if (
    (name === "GraphTimeoutError" || code === "GraphTimeoutError") &&
    typeof message === "string"
  ) {
    return new ValidationError(`${jobLabel} failed: ${message}`);
  }
  return new ValidationError(
    `${jobLabel} failed on the server before it could finish (SharePoint error — details are in the server log). Any changes already saved are kept; try again.`,
  );
}

export function toPublicErrorPayload(error: unknown): {
  status: number;
  error: string;
} {
  if (error instanceof UnauthorizedError) {
    return { status: 401, error: PUBLIC_ERROR_MESSAGES.accessDenied };
  }

  if (error instanceof AccessDeniedError) {
    return { status: 403, error: PUBLIC_ERROR_MESSAGES.accessDenied };
  }

  if (error instanceof NotFoundError) {
    return { status: 404, error: PUBLIC_ERROR_MESSAGES.notFound };
  }

  if (error instanceof ValidationError) {
    return { status: 400, error: error.message };
  }

  return { status: 500, error: PUBLIC_ERROR_MESSAGES.serverError };
}
