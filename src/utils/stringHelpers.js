export const toTrimAndLowerCase = (str) => str?.replace(/\s/g, "")?.toLowerCase();

export function extractNameFromEmail(email) {
  // Split the email address by the "@" symbol
  const parts = email.split("@");

  // Check if there are two parts (username and domain)
  if (parts.length === 2) {
    // Extract the portion before "@" and remove extra characters
    const username = parts[0].replace(/[^a-zA-Z\s]/g, "").trim();

    // Return the cleaned up username
    return username;
  } else {
    // If the email format is not as expected, return null or handle the error accordingly
    return null;
  }
}
