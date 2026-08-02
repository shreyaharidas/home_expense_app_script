import { TWO_DIGIT } from "../constants";

export function currentTime(): string {
  const currentDate = new Date();

  const formattedDate = currentDate.toLocaleString("en-GB", {
    day: TWO_DIGIT,
    month: TWO_DIGIT,
    year: TWO_DIGIT,
    hour: TWO_DIGIT,
    minute: TWO_DIGIT,
    hour12: false,
  } as Intl.DateTimeFormatOptions);

  return formattedDate; // Example output: "23-12-24 15:45"
}
