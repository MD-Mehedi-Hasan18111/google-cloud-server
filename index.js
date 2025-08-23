const express = require("express");
const cors = require("cors");
const { google } = require("googleapis");
const { getAllSheetsData, getAuthClientFromToken } = require("./utils");
const multer = require("multer");
require("dotenv").config();

const upload = multer({ storage: multer.memoryStorage() });

// origins
const origins = [
  "http://localhost:3000",
  "http://localhost:3000/home/",
  "https://dev-reef.netlify.app",
  "https://app.reef.lat",
];

const corsOptions = {
  origin: origins,
  credentials: true,
  optionsSuccessStatus: 200,
};

const app = express();
app.use(express.json());

app.use(cors(corsOptions)); // ✅ apply CORS
app.options("*", cors(corsOptions)); // ✅ handle preflight

const port = process.env.PORT || 3000;

// 4. List user’s Google Sheets files
app.get("/google/sheets", async (req, res) => {
  try {
    const token = req.headers.authorization?.split(" ")[1];
    const authClient = getAuthClientFromToken(token);

    const drive = google.drive({ version: "v3", auth: authClient });
    const { data } = await drive.files.list({
      q: "mimeType='application/vnd.google-apps.spreadsheet'",
      fields: "files(id, name)",
    });
    res.json(data.files);
  } catch (err) {
    console.error(err);
    res.status(500).send("Error listing sheets");
  }
});

// 5. List user’s Google Sheets files
app.get("/google/sheet/:id", async (req, res) => {
  try {
    const token = req.headers.authorization?.split(" ")[1];
    const authClient = getAuthClientFromToken(token);

    const allData = await getAllSheetsData(req.params.id, authClient);
    res.json(allData);
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: err.message });
  }
});

// 6. Upload Excel → Convert to Google Sheet
app.post("/google/sheet/create", upload.single("file"), async (req, res) => {
  try {
    const token = req.headers.authorization?.split(" ")[1];
    const authClient = getAuthClientFromToken(token);

    if (!req.file) {
      return res.status(400).json({ error: "No file uploaded" });
    }

    const drive = google.drive({ version: "v3", auth: authClient });

    // Upload Excel file and convert to Google Sheet
    const { data } = await drive.files.create({
      requestBody: {
        name: req.file.originalname.replace(/\.[^/.]+$/, ""), // remove extension
        mimeType: "application/vnd.google-apps.spreadsheet",
      },
      media: {
        mimeType: req.file.mimetype,
        body: Buffer.from(req.file.buffer),
      },
      fields: "id, webViewLink, webContentLink",
    });

    res.json({
      id: data.id,
      url: data.webViewLink, // Google Sheet open URL
    });
  } catch (err) {
    console.error("Error creating sheet", err);
    res.status(500).json({ error: err.message });
  }
});

// 7. Get List of Emails sent to hello@reef.lat
app.get("/emails", async (req, res) => {
  try {
    const token = req.headers.authorization?.split(" ")[1];
    const authClient = getAuthClientFromToken(token);
    const gmail = google.gmail({ version: "v1", auth: authClient });

    const response = await gmail.users.messages.list({
      userId: "me", // "me" means the authenticated user (hello@reef.lat)
      maxResults: 10, // change as needed
      q: "to:hello@reef.lat", // filter emails sent TO this address
    });

    const messages = response.data.messages || [];

    // Fetch full details of each email
    const emailDetails = await Promise.all(
      messages.map(async (msg) => {
        const fullMessage = await gmail.users.messages.get({
          userId: "me",
          id: msg.id,
        });

        return {
          id: msg.id,
          snippet: fullMessage.data.snippet,
          headers: fullMessage.data.payload.headers.filter((h) =>
            ["From", "To", "Subject", "Date"].includes(h.name)
          ),
        };
      })
    );

    res.json(emailDetails);
  } catch (err) {
    console.error("Error fetching emails", err);
    res.status(500).json({ error: "Failed to fetch emails" });
  }
});

// 8. Get Replies to an Email (Thread)
app.get("/emails/:id/replies", async (req, res) => {
  try {
    const { id } = req.params;
    const authClient = getAuthClientFromToken(token);
    const gmail = google.gmail({ version: "v1", auth: authClient });

    const message = await gmail.users.messages.get({
      userId: "me",
      id,
      format: "full",
    });

    const threadId = message.data.threadId;

    // Fetch the whole thread
    const thread = await gmail.users.threads.get({
      userId: "me",
      id: threadId,
    });

    res.json(thread.data.messages);
  } catch (err) {
    console.error("Error fetching replies", err);
    res.status(500).json({ error: "Failed to fetch replies" });
  }
});

app.listen(port, () => console.log(`Running on http://localhost:${port}`));

// Default route
app.get("/", (req, res) => {
  res.send("Google Cloud Server Running...");
});

module.exports = app;
