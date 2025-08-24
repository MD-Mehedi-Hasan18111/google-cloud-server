const { google } = require("googleapis");
const { v4: uuidv4 } = require("uuid");

function getAuthClientFromToken(token) {
  const client = new google.auth.OAuth2(
    process.env.GOOGLE_CLIENT_ID,
    process.env.GOOGLE_CLIENT_SECRET,
    process.env.GOOGLE_REDIRECT_URI
  );
  client.setCredentials({ access_token: token });
  return client;
}

async function isSheetGraphical(spreadsheetId, sheetId, tabName, authClient) {
  const sheetsApi = google.sheets({ version: "v4", auth: authClient });

  try {
    // Get sheet metadata to check for drawings and charts
    const sheetMetadata = await sheetsApi.spreadsheets.get({
      spreadsheetId,
      fields: "sheets(properties(sheetId,title),charts,drawings)",
    });

    const currentSheet = sheetMetadata.data.sheets.find(
      (s) => s.properties.sheetId === sheetId
    );

    // Check for drawings (images, shapes)
    const hasDrawings =
      currentSheet.drawings && currentSheet.drawings.length > 0;

    // Check for charts
    const hasCharts = currentSheet.charts && currentSheet.charts.length > 0;

    // Check for merged cells
    const mergedCellsRes = await sheetsApi.spreadsheets.get({
      spreadsheetId,
      fields: `sheets(properties(sheetId,title),merges)`,
    });

    const currentSheetMerges = mergedCellsRes.data.sheets.find(
      (s) => s.properties.sheetId === sheetId
    );
    const hasMergedCells =
      currentSheetMerges.merges && currentSheetMerges.merges.length > 0;

    // Check if sheet is mostly empty but has formatting
    const valuesRes = await sheetsApi.spreadsheets.values.get({
      spreadsheetId,
      range: tabName,
    });

    const rows = valuesRes.data.values || [];
    const isEmpty = rows.length === 0;
    const allBlank =
      !isEmpty &&
      rows.every((row) =>
        row.every((cell) => !cell || cell.toString().trim() === "")
      );

    // If sheet has drawings, charts, OR is completely blank → consider graphical
    if (hasDrawings || hasCharts || isEmpty || allBlank) {
      return true;
    }

    // If sheet has merged cells AND less than 50% of cells have data → consider graphical
    if (hasMergedCells) {
      const totalCells =
        rows.length * Math.max(...rows.map((row) => row.length));
      const filledCells = rows.reduce(
        (count, row) =>
          count +
          row.filter((cell) => cell && cell.toString().trim() !== "").length,
        0
      );

      const fillPercentage = (filledCells / totalCells) * 100;
      if (fillPercentage < 50) {
        return true;
      }
    }

    return false;
  } catch (error) {
    console.error(`Error detecting graphical sheet ${tabName}:`, error.message);
    // Fallback to original blank check
    const valuesRes = await sheetsApi.spreadsheets.values.get({
      spreadsheetId,
      range: tabName,
    });

    const rows = valuesRes.data.values || [];
    const isEmpty = rows.length === 0;
    const allBlank =
      !isEmpty &&
      rows.every((row) =>
        row.every((cell) => !cell || cell.toString().trim() === "")
      );

    return isEmpty || allBlank;
  }
}

async function uploadTabAsExcel(spreadsheetId, tabName, authClient, sheetId) {
  const drive = google.drive({ version: "v3", auth: authClient });
  const sheetsApi = google.sheets({ version: "v4", auth: authClient });

  let tempFileId = null;

  try {
    // 1. Create a temporary copy with only the specific sheet
    const copyResponse = await drive.files.copy({
      fileId: spreadsheetId,
      requestBody: {
        name: `Temp_${tabName}_${Date.now()}`,
      },
    });

    tempFileId = copyResponse.data.id;

    // 2. Delete all other sheets except the target one
    const spreadsheetData = await sheetsApi.spreadsheets.get({
      spreadsheetId: tempFileId,
    });

    const sheetsToDelete = spreadsheetData.data.sheets
      .filter((sheet) => sheet.properties.sheetId !== sheetId)
      .map((sheet) => sheet.properties.sheetId);

    if (sheetsToDelete.length > 0) {
      await sheetsApi.spreadsheets.batchUpdate({
        spreadsheetId: tempFileId,
        requestBody: {
          requests: sheetsToDelete.map((sheetId) => ({
            deleteSheet: {
              sheetId: sheetId,
            },
          })),
        },
      });
    }

    // 3. Export the temporary file as Excel
    const exportResponse = await drive.files.export(
      {
        fileId: tempFileId,
        mimeType:
          "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      },
      {
        responseType: "arraybuffer",
      }
    );

    const buffer = Buffer.from(exportResponse.data);

    // 4. Upload to S3 (same as before)
    const boundary =
      "----WebKitFormBoundary" + Math.random().toString(16).substring(2);
    const filename = `${tabName}.xlsx`;

    const formDataParts = [
      `--${boundary}\r\n`,
      `Content-Disposition: form-data; name="file"; filename="${filename}"\r\n`,
      `Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet\r\n\r\n`,
      buffer,
      `\r\n--${boundary}--\r\n`,
    ];

    const formDataBuffer = Buffer.concat(
      formDataParts.map((part) =>
        typeof part === "string" ? Buffer.from(part, "utf8") : part
      )
    );

    const uploadRes = await fetch(
      "https://dgo2elt3cskpo.cloudfront.net/upload_to_s3",
      {
        method: "POST",
        body: formDataBuffer,
        headers: {
          "Content-Type": `multipart/form-data; boundary=${boundary}`,
          "Content-Length": formDataBuffer.length.toString(),
        },
      }
    );

    if (!uploadRes.ok) {
      const errorText = await uploadRes.text();
      throw new Error(`S3 upload failed: ${uploadRes.status} - ${errorText}`);
    }

    const responseData = await uploadRes.json();
    return responseData;
  } catch (error) {
    console.error(`Failed to upload ${tabName} to S3:`, error.message);
    throw error;
  } finally {
    // 5. Clean up temporary file
    if (tempFileId) {
      try {
        await drive.files.delete({
          fileId: tempFileId,
        });
      } catch (deleteError) {
        console.warn("Could not delete temporary file:", deleteError.message);
      }
    }
  }
}

async function getAllSheetsData(spreadsheetId, authClient) {
  const sheetsApi = google.sheets({ version: "v4", auth: authClient });
  const { data } = await sheetsApi.spreadsheets.get({
    spreadsheetId,
    fields: "sheets(properties(sheetId,title))",
  });

  const tabs = data.sheets.map((s) => ({
    title: s.properties.title,
    sheetId: s.properties.sheetId,
  }));

  const allTables = [];

  for (const { title: tab, sheetId } of tabs) {
    try {
      console.log(`Processing tab: ${tab}`);

      // Check if this is a graphical sheet
      const isGraphical = await isSheetGraphical(
        spreadsheetId,
        sheetId,
        tab,
        authClient
      );

      if (isGraphical) {
        console.log(`Detected graphical sheet: ${tab}, uploading to S3`);

        try {
          const responsePath = await uploadTabAsExcel(
            spreadsheetId,
            tab,
            authClient,
            sheetId
          );

          allTables.push({
            id: uuidv4(),
            tableName: tab,
            columns: [],
            rows: [],
            createdBy: "import",
            excelPreview: { path: responsePath },
          });

          console.log(`Successfully uploaded ${tab} to S3: ${responsePath}`);
        } catch (uploadError) {
          console.error(
            `Failed to upload graphical sheet ${tab}:`,
            uploadError.message
          );
        }

        continue;
      }

      // Process as data table
      console.log(`Processing as data table: ${tab}`);

      const res = await sheetsApi.spreadsheets.values.get({
        spreadsheetId,
        range: tab,
      });

      const rows = res.data.values || [];
      const isEmpty = rows.length === 0;

      if (isEmpty) {
        continue;
      }

      const rawHeaders = rows[0] || [];
      const headers = [];
      const seen = {};

      rawHeaders.forEach((header, i) => {
        let h =
          header && header.toString().trim()
            ? header.toString().trim()
            : `Column${i + 1}`;
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

        // ✅ Pad rows until reaching 1000
        if (rowData.length < 1000) {
          const emptyRow = {};
          headers.forEach((header) => {
            emptyRow[header] = "";
          });

          // Fill missing rows in one go
          rowData = rowData.concat(
            Array.from({ length: 1000 - rowData.length }, () => ({
              ...emptyRow,
            }))
          );
        }
      } else {
        // only headers → add blank rows
        rowData = Array.from({ length: Math.min(100, 1000) }, () => {
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
    } catch (error) {
      console.error(`Error processing tab ${tab}:`, error.message);
    }
  }

  return allTables;
}

module.exports = { getAuthClientFromToken, getAllSheetsData };
