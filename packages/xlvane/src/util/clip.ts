// Text from a file shown in messages and error details is shortened, so that a value of
// megabytes is not copied into every message, log line or report that shows it (D-H01).

/** The most characters of a value from the file that a message or detail shows. */
export const MAX_SHOWN = 200;

/** Shortens a value from the file for messages and details. */
export function clip(value: string): string {
  return value.length > MAX_SHOWN ? `${value.slice(0, MAX_SHOWN)}…` : value;
}
