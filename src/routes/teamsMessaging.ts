import { Router } from "express";
import { getTeamsChats, getTeamsChannels, getTeamsGroups, getTeamsPeople, sendTeamsMessage } from "../services/teamsMessagingService";
import { getTeamsCredentials } from "../services/credentialStore";
import { getOrganizationIdFromRequest } from "../utils/requestContext";

export const teamsMessagingRouter = Router();

teamsMessagingRouter.get("/chats", async (req, res) => {
  try {
    const organizationId = getOrganizationIdFromRequest(req);
    const { accessToken } = await getTeamsCredentials(organizationId);
    const chats = await getTeamsChats(accessToken);
    return res.status(200).json({ chats });
  } catch (error) {
    const message = error instanceof Error ? error.message : "Failed to list Teams chats";
    return res.status(400).json({ status: "error", message });
  }
});

teamsMessagingRouter.get("/groups", async (req, res) => {
  try {
    const organizationId = getOrganizationIdFromRequest(req);
    const { accessToken } = await getTeamsCredentials(organizationId);
    const groups = await getTeamsGroups(accessToken);
    return res.status(200).json({ groups });
  } catch (error) {
    const message = error instanceof Error ? error.message : "Failed to list Teams groups";
    return res.status(400).json({ status: "error", message });
  }
});

teamsMessagingRouter.get("/people", async (req, res) => {
  try {
    const organizationId = getOrganizationIdFromRequest(req);
    const { accessToken } = await getTeamsCredentials(organizationId);
    const people = await getTeamsPeople(accessToken);
    return res.status(200).json({ people });
  } catch (error) {
    const message = error instanceof Error ? error.message : "Failed to list Teams people";
    return res.status(400).json({ status: "error", message });
  }
});

teamsMessagingRouter.post("/send", async (req, res) => {
  try {
    const organizationId = getOrganizationIdFromRequest(req);
    const { accessToken } = await getTeamsCredentials(organizationId);

    const result = await sendTeamsMessage({
      accessToken,
      destinationType: req.body?.destinationType,
      destinationId: req.body?.destinationId,
      teamId: req.body?.teamId,
      message: req.body?.message,
      file: req.body?.file,
    });

    return res.status(200).json({ status: "success", ...result });
  } catch (error) {
    const message = error instanceof Error ? error.message : "Failed to send Teams message";
    return res.status(400).json({ status: "error", message });
  }
});
