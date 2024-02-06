import { getAccessToken } from "../../../utils/auth";
export default async function handler(req, res) {
  const auth = req.headers.authorization;
  try {
    const response = await getAccessToken(auth);
    res.status(response?.error ? 500 : 200).json({ ...response });
  } catch (error) {
    console.log(error);
    res.status(500).json({ error: error?.message });
  }
}

//"AADSTS65001: The user or administrator has not consented to use the application with ID '7f4c222f-891d-4883-8a18-18a3f0bbf3ed' named 'Karma-Outlook-Activities-UAT'. Send an interactive authorization request for this user and resource. Trace ID: 0bbfcebd-cf01-46b2-a16b-75bb7cf6b000 Correlation ID: 6aac78e4-04d9-42b5-9fa8-f16c356efefd Timestamp: 2024-02-06 20:01:03Z"
//"error": "Error Code: invalid_grant; Microsoft Graph error {\"error\":\"invalid_grant\",\"error_description\":\"AADSTS65001: The user or administrator has not consented to use the application with ID '7f4c222f-891d-4883-8a18-18a3f0bbf3ed' named 'Karma-Outlook-Activities-UAT'. Send an interactive authorization request for this user and resource. Trace ID: 6b19beb2-2306-4212-bcd2-6b40eb36a100 Correlation ID: f3deba8f-4516-41e8-aeba-c55a4edd44f7 Timestamp: 2024-02-06 20:10:17Z\",\"error_codes\":[65001],\"timestamp\":\"2024-02-06 20:10:17Z\",\"trace_id\":\"6b19beb2-2306-4212-bcd2-6b40eb36a100\",\"correlation_id\":\"f3deba8f-4516-41e8-aeba-c55a4edd44f7\",\"suberror\":\"consent_required\"}"
