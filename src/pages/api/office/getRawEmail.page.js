import { getAdminToken, getGraphData } from "../../../utils/ms-graph";

export default async function handler(req, res) {
  try {
    const body = req.body;
    if (typeof body === "string" || !body) {
      return res.status(400).json({ message: "Missing body" });
    }
    const userId = body.userId;
    const messageId = body.messageId;
    if (!userId) return res.status(400).json({ message: "Missing userId" });
    if (!messageId) return res.status(400).json({ message: "Missing messageId" });

    const tokenResponse = await getAdminToken();
    const response = await getGraphData(tokenResponse.access_token, `/users/${userId}/messages/${messageId}/$value`);
    res.status(200).setHeader('content-type', 'text/plain').send(response);

  } catch (error) {
    res.status(500).json({ error: error?.message || error });
  }
}
