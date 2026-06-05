import { Router } from "express";
import { listSlackChannels, listSlackDMs, sendSlackMessage } from "../services/slackMessagingService";
import { getSlackCredentials } from "../services/credentialStore";
import { getOrganizationIdFromRequest } from "../utils/requestContext";

export const slackMessagingRouter = Router();

slackMessagingRouter.get("/channels", async (req, res) => {
  try {
    const organizationId = getOrganizationIdFromRequest(req);
    const { botToken } = await getSlackCredentials(organizationId);
    const channels = await listSlackChannels(botToken, "public_channel,private_channel");
    return res.status(200).json({ channels });
  } catch (error) {
    const message = error instanceof Error ? error.message : "Failed to list channels";
    return res.status(400).json({ status: "error", message });
  }
});

slackMessagingRouter.get("/dms", async (req, res) => {
  try {
    const organizationId = getOrganizationIdFromRequest(req);
    const { botToken } = await getSlackCredentials(organizationId);
    const channels = await listSlackDMs(botToken);
    return res.status(200).json({ channels });
  } catch (error) {
    const message = error instanceof Error ? error.message : "Failed to list DMs";
    return res.status(400).json({ status: "error", message });
  }
});

slackMessagingRouter.post("/send", async (req, res) => {
  try {
    const organizationId = getOrganizationIdFromRequest(req);
    const { botToken } = await getSlackCredentials(organizationId);

    const result = await sendSlackMessage({
      botToken,
      channelId: req.body?.channelId,
      message: req.body?.message,
      threadTs: req.body?.threadTs,
      file: req.body?.file,
    });

    return res.status(200).json({ status: "success", ...result });
  } catch (error) {
    const message = error instanceof Error ? error.message : "Failed to send Slack message";
    return res.status(400).json({ status: "error", message });
  }
});
