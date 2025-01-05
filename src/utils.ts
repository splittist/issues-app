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
