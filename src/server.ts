import cors from "cors";
import express from "express";

import { config } from "./config";
import { emailRouter } from "./routes/email";
import { slackOAuthRouter } from "./routes/slackOAuth";
import { slackMessagingRouter } from "./routes/slackMessaging";
import { teamsOAuthRouter } from "./routes/teamsOAuth";
import { teamsMessagingRouter } from "./routes/teamsMessaging";

const app = express();

// CORS configuration with proper options
const corsOptions = {
  origin: (origin: string | undefined, callback: (err: Error | null, allow?: boolean) => void) => {
    // Allow requests with no origin (like mobile apps, Curl requests)
    if (!origin) {
      return callback(null, true);
    }

    // Allow requests from configured frontend origin
    if (config.frontendOrigin === "*") {
      return callback(null, true);
    }

    // Exact match with frontend origin
    if (origin === config.frontendOrigin) {
      return callback(null, true);
    }

    // Allow localhost in dev (handle different ports)
    if (config.frontendOrigin.includes("localhost") && origin.includes("localhost")) {
      return callback(null, true);
    }

    // Reject all other origins
    callback(new Error("Not allowed by CORS"));
  },
  credentials: true,
  methods: ["GET", "POST", "PUT", "DELETE", "PATCH", "OPTIONS"],
  allowedHeaders: ["Content-Type", "Authorization", "Accept", "x-organization-id"],
  exposedHeaders: ["Content-Type", "Content-Length"],
  maxAge: 86400, // 24 hours
};

app.use(cors(corsOptions));
app.use(express.json({ limit: "15mb" }));

// Explicit preflight handler
app.options("*", cors(corsOptions));

app.get("/health", (_req, res) => {
  res.status(200).json({ ok: true, service: "communication-backend" });
});

app.use("/api/v1/email", emailRouter);
app.use("/api/v1/slack/oauth", slackOAuthRouter);
app.use("/api/v1/slack", slackMessagingRouter);
app.use("/api/v1/teams/oauth", teamsOAuthRouter);
app.use("/api/v1/teams", teamsMessagingRouter);

app.use((_req, res) => {
  res.status(404).json({ status: "error", message: "Not found" });
});

app.listen(config.port, () => {
  console.log(`communication-backend is running on port ${config.port}`);
});
