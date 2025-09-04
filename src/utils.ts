/**
 * Returns the current date in YYYY-MM-DD format for the curren timezone.
 * @returns The current date as a string.
 */
export const dateToday = () => {
  let date = new Date();
  const offset = date.getTimezoneOffset();
  date = new Date(date.getTime() - (offset * 60 * 1000));
  return date.toISOString().split('T')[0];
};

/**
 * Formats a date string to a human-readable format.
 * @param dateString - The date string to format (ISO format or Word document format).
 * @returns A human-readable date string.
 */
export const formatCommentDate = (dateString: string): string => {
  try {
    const date = new Date(dateString);
    if (isNaN(date.getTime())) {
      return dateString; // Return original if invalid
    }
    return date.toLocaleDateString('en-US', {
      year: 'numeric',
      month: 'short',
      day: 'numeric'
    });
  } catch {
    return dateString; // Return original on error
  }
};
