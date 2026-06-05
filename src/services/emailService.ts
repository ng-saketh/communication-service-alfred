import sendgrid, { MailService as SendMail } from "@sendgrid/mail";
import { config } from "../config";
import { SendFilePayload } from "../types";

const ALLOWED_FILE_TYPES = new Set(["image/png", "application/pdf"]);

interface EmailServiceConfig {
  apiKey: string;
  fromEmail: string;
  fromName?: string;
}

interface SendEmailInput {
  to: string | string[];
  subject: string;
  text?: string;
  html?: string;
  attachments?: SendFilePayload[];
}

interface EmailServiceResponse {
  success: boolean;
  messageId?: string;
  accepted?: string[];
  error?: string;
}

class EmailService {
  private readonly sendgrid: SendMail;
  private readonly config: EmailServiceConfig;

  constructor(serviceConfig: EmailServiceConfig) {
    this.config = serviceConfig;
    this.sendgrid = sendgrid;
    this.sendgrid.setApiKey(this.config.apiKey);
  }

  private toSendgridAttachment(file: SendFilePayload) {
    const fileName = String(file.fileName || "").trim();
    const mimeType = String(file.mimeType || "").trim();
    const contentBase64 = String(file.contentBase64 || "").trim();

    if (!fileName) {
      throw new Error("Missing attachment fileName");
    }

    if (!ALLOWED_FILE_TYPES.has(mimeType)) {
      throw new Error("Only PNG and PDF attachments are supported");
    }

    if (!contentBase64) {
      throw new Error("Missing attachment contentBase64");
    }

    const normalizedBase64 = contentBase64.replace(/\s/g, "");
    const decoded = Buffer.from(normalizedBase64, "base64");
    if (!decoded.length) {
      throw new Error(`Attachment ${fileName} is empty after decoding`);
    }

    return {
      filename: fileName,
      type: mimeType,
      disposition: "attachment" as const,
      content: normalizedBase64,
    };
  }

  async sendEmail(input: SendEmailInput): Promise<EmailServiceResponse> {
    try {
      const recipients = Array.isArray(input.to) ? input.to : [input.to];
      const to = recipients.map((recipient) => String(recipient || "").trim()).filter(Boolean);

      if (!to.length) {
        throw new Error("At least one recipient is required");
      }

      const subject = String(input.subject || "").trim();
      if (!subject) {
        throw new Error("Email subject is required");
      }

      const text = typeof input.text === "string" ? input.text : "";
      const html = typeof input.html === "string" ? input.html : "";

      if (!text.trim() && !html.trim()) {
        throw new Error("Provide text and/or html content");
      }

      const attachments = (input.attachments || []).map((file) => this.toSendgridAttachment(file));

      const msg = {
        to,
        from: {
          email: this.config.fromEmail,
          name: this.config.fromName || "Alfred Platform",
        },
        subject,
        attachments,
        categories: ["transactional"],
        headers: {
          "X-Priority": "1",
          "X-Mailer": "Alfred Platform",
        },
        trackingSettings: {
          clickTracking: {
            enable: false,
            enableText: false,
          },
          openTracking: {
            enable: false,
          },
          subscriptionTracking: {
            enable: false,
          },
        },
        mailSettings: {
          bypassListManagement: {
            enable: false,
          },
        },
      } as unknown as sendgrid.MailDataRequired;

      if (text.trim()) {
        msg.text = text;
      }

      if (html.trim()) {
        msg.html = html;
      }

      const [response] = await this.sendgrid.send(msg);

      return {
        success: true,
        messageId: response.headers["x-message-id"] as string | undefined,
        accepted: to,
      };
    } catch (error: any) {
      const providerMessage =
        error?.response?.body?.errors?.[0]?.message || error?.message || "Failed to send email";

      console.error("EmailService.sendEmail failed:", {
        error: error?.message,
        code: error?.code,
        response: error?.response?.body,
      });

      return {
        success: false,
        error: providerMessage,
      };
    }
  }

}

const emailService = new EmailService({
  apiKey: config.sendgridApiKey,
  fromEmail: config.fromEmail,
  fromName: config.fromName,
});

export async function sendEmail(input: SendEmailInput) {
  const result = await emailService.sendEmail(input);

  if (!result.success) {
    throw new Error(result.error || "Failed to send email");
  }

  return {
    messageId: result.messageId,
    accepted: result.accepted || [],
  };
}
