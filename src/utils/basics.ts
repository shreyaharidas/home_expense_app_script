export function currentTime(): string {
  const currentDate = new Date();

  const formattedDate = currentDate.toLocaleString("en-GB", {
    day: "2-digit",
    month: "2-digit",
    year: "2-digit",
    hour: "2-digit",
    minute: "2-digit",
    hour12: false,
  });

  return formattedDate; // Example output: "23-12-24 15:45"
}
