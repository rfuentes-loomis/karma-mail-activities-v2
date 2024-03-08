import { getAdminToken, getGraphData } from "../../../utils/ms-graph";

export default async function handler(req, res) {
  if (req.method === "POST") {
    try {
      const body = req.body
      if (typeof body === "string" || !body || !body?.email) {
        return res.status(400).json({ message: "Missing email" });
      }
      const tokenResponse = await getAdminToken();
      const userResponse = await getGraphData(tokenResponse.access_token, `/users?$filter=mail eq '${body?.email}'`);
      const user = userResponse.value?.pop();
      res.status(200).json(user || null);
    } catch (error) {
      res.status(500).json({ error: error?.message || error });
    }
  } else {
    // Handle other HTTP methods
    res.status(405).json({ message: "Method Not Allowed" });
  }
}
