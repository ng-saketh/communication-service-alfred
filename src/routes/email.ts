import { Router } from "express";
import { sendEmail } from "../services/emailService";

export const emailRouter = Router();

function getErrorStatus(message: string): number {
  if (/unauthorized/i.test(message)) {
    return 401;
  }

  if (/forbidden|rejected sender identity/i.test(message)) {
    return 403;
  }

  if (/required|missing|invalid|supported/i.test(message)) {
    return 422;
  }

  return 400;
}

emailRouter.post("/send", async (req, res) => {
  try {
    const payload = {
      to: req.body?.to,
      subject: req.body?.subject,
      text: req.body?.text,
      html: req.body?.html,
      attachments: req.body?.attachments,
    };

    const result = await sendEmail({
      to: payload.to,
      subject: payload.subject,
      text: payload.text,
      html: payload.html,
      attachments: payload.attachments,
    });

    return res.status(200).json({
      status: "success",
      message: "Email sent successfully",
      ...result,
    });
  } catch (error) {
    const message = error instanceof Error ? error.message : "Failed to send email";
    return res.status(getErrorStatus(message)).json({ status: "error", message });
  }
});
