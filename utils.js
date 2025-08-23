const { google } = require("googleapis");

function getAuthClientFromToken(token) {
  const client = new google.auth.OAuth2(
    process.env.GOOGLE_CLIENT_ID,
    process.env.GOOGLE_CLIENT_SECRET,
    process.env.GOOGLE_REDIRECT_URI
  );
  client.setCredentials({ access_token: token });
  return client;
}

// Fetch values from a specific sheet
async function getAllSheetsData(spreadsheetId, authClient) {
  const sheetsApi = google.sheets({ version: "v4", auth: authClient });

  const { data } = await sheetsApi.spreadsheets.get({ spreadsheetId });
  const tabs = data.sheets.map((s) => s.properties.title);

  const allData = {};

  for (const tab of tabs) {
    const range = `${tab}`;
    const res = await sheetsApi.spreadsheets.values.get({
      spreadsheetId,
      range,
    });

    const rows = res.data.values || [];

    if (rows.length > 0) {
      const rawHeaders = rows[0];

      // clean headers: fill empty, handle duplicates
      const headers = [];
      const seen = {};

      rawHeaders.forEach((header, i) => {
        let h = header && header.trim() ? header.trim() : `Column${i + 1}`;

        if (seen[h]) {
          let count = seen[h] + 1;
          seen[h] = count;
          h = `${h}_${count}`;
        } else {
          seen[h] = 1;
        }

        headers.push(h);
      });

      const jsonData = rows.slice(1).map((row) => {
        const obj = {};
        headers.forEach((header, i) => {
          obj[header] = row[i] || "";
        });
        return obj;
      });

      allData[tab] = jsonData;
    } else {
      allData[tab] = [];
    }
  }

  return allData;
}

module.exports = { getAuthClientFromToken, getAllSheetsData };
