const express = require("express");
const cors = require("cors");
const { google } = require("googleapis");
const { getAllSheetsData } = require("./utils");
const multer = require("multer");
const { Readable } = require("stream");
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
  methods: ["GET", "POST"],
  credentials: true,
  optionsSuccessStatus: 200,
};

const app = express();
app.use(express.json());

app.use(cors(corsOptions)); // ✅ apply CORS

const port = process.env.PORT || 5000;

// 4. List user’s Google Sheets files
app.get("/google/sheets", async (req, res) => {
  try {
    const token = req.headers.authorization?.split(" ")[1];
    const authClient = new google.auth.OAuth2();
    authClient.setCredentials({ access_token: token });

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
    const authClient = new google.auth.OAuth2();
    authClient.setCredentials({ access_token: token });

    const allData = await getAllSheetsData(req.params.id, authClient);
    res.json(allData);
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: err.message });
  }
});

// Upload Excel → Convert to Google Sheet
app.post("/google/sheet/create", upload.single("file"), async (req, res) => {
  try {
    const token = req.headers.authorization?.split(" ")[1];
    if (!token) {
      return res.status(401).json({ error: "Missing access token" });
    }

    const authClient = new google.auth.OAuth2();
    authClient.setCredentials({ access_token: token });

    const drive = google.drive({ version: "v3", auth: authClient });

    let fileName;
    let mimeType;
    let buffer;

    if (req.file) {
      // 📂 Case 1: Uploaded file
      fileName = req.file.originalname.replace(/\.[^/.]+$/, "");
      mimeType = req.file.mimetype;
      buffer = req.file.buffer;
    } else if (req.body.fileUrl) {
      // 🌐 Case 2: Remote file URL
      const fileUrl = req.body.fileUrl;
      const response = await fetch(fileUrl);
      if (!response.ok) throw new Error("Failed to fetch file from URL");
      // Convert arrayBuffer -> Buffer
      const arrayBuffer = await response.arrayBuffer();
      buffer = Buffer.from(arrayBuffer);
      fileName =
        fileUrl
          .split("/")
          .pop()
          ?.replace(/_[a-zA-Z0-9]{3}(?=\.[^/.]+$)/, "")
          ?.replace(/\.[^/.]+$/, "")
          // Remove the file extension
          .replace(/\.[^/.]+$/, "")
          .replace(/-/g, " ")
          .replace(/_/g, " ") || "Untitled Sheet";
      mimeType =
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
    } else {
      return res.status(400).json({ error: "No file or file url provided" });
    }

    // Convert buffer → stream
    const bufferStream = Readable.from(buffer);

    // Upload & convert to Google Sheet
    const { data } = await drive.files.create({
      requestBody: {
        name: fileName,
        mimeType: "application/vnd.google-apps.spreadsheet",
      },
      media: {
        mimeType,
        body: bufferStream,
      },
      fields: "id, webViewLink",
    });

    res.json({
      id: data.id,
      url: data.webViewLink,
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
    const authClient = new google.auth.OAuth2();
    authClient.setCredentials({ access_token: token });
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
    const authClient = new google.auth.OAuth2();
    authClient.setCredentials({ access_token: token });
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
