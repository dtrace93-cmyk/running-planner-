const express = require('express');

const app = express();
app.use(express.json());

// Simple in-memory token store. Replace with your persistent storage as needed.
const tokenStore = {
  accessToken: null,
  refreshToken: null,
  expiresAt: null,
};

async function ensureFreshAccessToken() {
  if (!tokenStore.accessToken) {
    throw new Error('No access token available');
  }

  const now = Date.now();
  if (tokenStore.expiresAt && now >= tokenStore.expiresAt) {
    if (!tokenStore.refreshToken) {
      throw new Error('No refresh token available');
    }
    // TODO: implement refresh logic against Strava and update tokenStore.accessToken and tokenStore.expiresAt
    // Placeholder to demonstrate failure path until wired to real refresh logic.
    throw new Error('Access token expired and refresh flow not implemented');
  }

  return tokenStore.accessToken;
}

// New status route to report connection health and attempt token refresh.
app.get('/api/strava/status', async (req, res) => {
  if (!tokenStore.accessToken) {
    return res.json({ connected: false });
  }

  try {
    await ensureFreshAccessToken();
    return res.json({ connected: true });
  } catch (err) {
    return res.json({ connected: false, error: err.message });
  }
});

// Export the app for use by a server entry point or tests.
module.exports = app;

if (require.main === module) {
  const port = process.env.PORT || 4000;
  app.listen(port, () => {
    console.log(`Server running on port ${port}`);
  });
}
