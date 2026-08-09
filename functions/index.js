const { onRequest } = require("firebase-functions/v2/https");
const cors = require('cors')({ origin: true });

/**
 * Thin proxy — forwards frontend requests to the GAS web app.
 * The GAS URL never appears in browser code. The GAS web app handles
 * all auth/sheet access as the owner account.
 * Requires one Firebase secret: GAS_EXEC_URL (the GAS web app deployment URL).
 */
exports.api = onRequest({
  secrets: ["GAS_EXEC_URL"],
  region: "australia-southeast1",
  maxInstances: 3,
  memory: "256MiB",
  timeoutSeconds: 30
}, async (req, res) => {
  return cors(req, res, async () => {
    try {
      const gasUrl = process.env.GAS_EXEC_URL;
      if (!gasUrl) {
        return res.status(500).json({ error: "GAS_EXEC_URL secret not configured." });
      }

      const gasRes = await fetch(gasUrl, {
        method: 'POST',
        headers: { 'Content-Type': 'text/plain;charset=utf-8' },
        body: JSON.stringify(req.body)
      });

      const data = await gasRes.json();

      if (data.error) {
        return res.status(400).json(data);
      }
      return res.json(data);
    } catch (error) {
      console.error('GAS proxy error:', error);
      res.status(500).json({ error: error.message });
    }
  });
});
