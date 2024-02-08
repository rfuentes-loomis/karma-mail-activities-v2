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