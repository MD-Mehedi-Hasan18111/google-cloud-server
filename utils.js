const { google } = require("googleapis");
const { v4: uuidv4 } = require("uuid");
const FormData = require("form-data");
const fs = require("fs");
const os = require("os");
const path = require("path");
const xlsx = require("xlsx");

function getAuthClientFromToken(token) {
  const client = new google.auth.OAuth2(
    process.env.GOOGLE_CLIENT_ID,
    process.env.GOOGLE_CLIENT_SECRET,
    process.env.GOOGLE_REDIRECT_URI
  );
  client.setCredentials({ access_token: token });
  return client;
}

async function uploadTabAsExcel(spreadsheetId, tabName, authClient) {
  // 1. Download tab values
  const sheetsApi = google.sheets({ version: "v4", auth: authClient });
  const res = await sheetsApi.spreadsheets.values.get({
    spreadsheetId,
    range: tabName,
  });

  const rows = res.data.values || [];

  // 2. Convert to XLSX buffer
  const ws = xlsx.utils.aoa_to_sheet(rows);
  const wb = xlsx.utils.book_new();
  xlsx.utils.book_append_sheet(wb, ws, tabName);

  const tmpFile = path.join(os.tmpdir(), `${tabName}-${Date.now()}.xlsx`);
  xlsx.writeFile(wb, tmpFile);

  try {
    // 3. Upload to S3 API
    const form = new FormData();
    form.append("file", fs.createReadStream(tmpFile));

    const uploadRes = await fetch(
      "https://dgo2elt3cskpo.cloudfront.net/upload_to_s3",
      {
        method: "POST",
        body: form,
      }
    );

    if (!uploadRes.ok) {
      throw new Error("Failed to upload tab to S3");
    }

    const { path: responsePath } = await uploadRes.json();
    return responsePath;
  } finally {
    // 4. Cleanup temp file
    fs.unlink(tmpFile, (err) => {
      if (err) console.error("Failed to delete temp file:", err);
    });
  }
}

async function getAllSheetsData(spreadsheetId, authClient) {
  const sheetsApi = google.sheets({ version: "v4", auth: authClient });
  const { data } = await sheetsApi.spreadsheets.get({ spreadsheetId });
  const tabs = data.sheets?.map((s) => s.properties?.title) || [];

  const allTables = [];

  for (const tab of tabs) {
    const res = await sheetsApi.spreadsheets.values.get({
      spreadsheetId,
      range: tab,
    });

    const rows = res.data.values || [];

    const isEmpty = rows.length === 0;
    const allBlank =
      !isEmpty &&
      rows.every((row) => row.every((cell) => !cell || cell.trim() === ""));

    if (isEmpty || allBlank) {
      // graphical tab → upload to S3
      const responsePath = await uploadTabAsExcel(
        spreadsheetId,
        tab,
        authClient
      );

      allTables.push({
        id: uuidv4(),
        tableName: tab,
        columns: [],
        rows: [],
        createdBy: "import",
        excelPreview: { path: responsePath },
      });

      continue;
    }

    // otherwise → normal table parsing
    const rawHeaders = rows[0];
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

    const columns = headers.map((h) => ({
      id: uuidv4(),
      dataType: "string",
      colName: h,
      width: 300,
    }));

    let rowData = [];
    if (rows.length > 1) {
      rowData = rows.slice(1).map((row) => {
        const obj = {};
        headers.forEach((header, i) => {
          obj[header] = row[i] || "";
        });
        return obj;
      });
    } else {
      // only headers → add 1000 blank rows
      rowData = Array.from({ length: 1000 }, () => {
        const obj = {};
        headers.forEach((h) => {
          obj[h] = "";
        });
        return obj;
      });
    }

    allTables.push({
      id: uuidv4(),
      tableName: tab,
      columns,
      rows: rowData,
      createdBy: "import",
    });
  }

  return allTables;
}

module.exports = { getAuthClientFromToken, getAllSheetsData };
