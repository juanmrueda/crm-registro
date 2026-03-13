// =============================================
// Lambda: crm-send-pdf-email
// Runtime: Node.js 20.x
// Timeout: 60s | Memory: 256MB
// =============================================
// Variables de entorno requeridas:
//   FROM_EMAIL  - email verificado en SES
//   API_KEY     - clave para autenticar requests
// =============================================

import { SESClient, SendRawEmailCommand } from '@aws-sdk/client-ses';

const ses = new SESClient();
const FROM_EMAIL = process.env.FROM_EMAIL;
const API_KEY = process.env.API_KEY;
const BATCH_SIZE = 5;

export const handler = async (event) => {
    // CORS preflight
    if (event.requestContext?.http?.method === 'OPTIONS') {
        return corsResponse(200, '');
    }

    // Validate API Key
    const apiKey = event.headers?.['x-api-key'] || event.headers?.['X-Api-Key'];
    if (apiKey !== API_KEY) {
        return corsResponse(401, { error: 'No autorizado' });
    }

    let body;
    try {
        body = typeof event.body === 'string' ? JSON.parse(event.body) : event.body;
    } catch {
        return corsResponse(400, { error: 'JSON invalido' });
    }

    const { recipients, pdfBase64, pdfName, subject, htmlBody, senderName } = body;

    // Validate required fields
    if (!recipients?.length) return corsResponse(400, { error: 'recipients es requerido' });
    if (!pdfBase64) return corsResponse(400, { error: 'pdfBase64 es requerido' });
    if (!pdfName) return corsResponse(400, { error: 'pdfName es requerido' });
    if (!subject) return corsResponse(400, { error: 'subject es requerido' });
    if (!htmlBody) return corsResponse(400, { error: 'htmlBody es requerido' });

    const fromHeader = senderName ? `"${senderName}" <${FROM_EMAIL}>` : FROM_EMAIL;
    let sent = 0;
    let failed = 0;
    const errors = [];

    // Send in batches
    for (let i = 0; i < recipients.length; i += BATCH_SIZE) {
        const batch = recipients.slice(i, i + BATCH_SIZE);
        const promises = batch.map(async (recipient) => {
            try {
                const personalizedHtml = htmlBody.replace(/\{\{nombre\}\}/g, recipient.nombre || 'Estudiante');
                const rawEmail = buildMimeEmail({
                    from: fromHeader,
                    to: recipient.email,
                    subject,
                    html: personalizedHtml,
                    pdfBase64,
                    pdfName
                });

                await ses.send(new SendRawEmailCommand({
                    RawMessage: { Data: new TextEncoder().encode(rawEmail) }
                }));

                sent++;
            } catch (err) {
                failed++;
                errors.push({ email: recipient.email, error: err.message });
            }
        });

        await Promise.all(promises);
    }

    return corsResponse(200, { sent, failed, errors, total: recipients.length });
};

function buildMimeEmail({ from, to, subject, html, pdfBase64, pdfName }) {
    const boundary = `----=_Part_${Date.now()}_${Math.random().toString(36).slice(2)}`;

    return [
        `From: ${from}`,
        `To: ${to}`,
        `Subject: =?UTF-8?B?${btoa(unescape(encodeURIComponent(subject)))}?=`,
        `MIME-Version: 1.0`,
        `Content-Type: multipart/mixed; boundary="${boundary}"`,
        ``,
        `--${boundary}`,
        `Content-Type: text/html; charset=UTF-8`,
        `Content-Transfer-Encoding: 7bit`,
        ``,
        html,
        ``,
        `--${boundary}`,
        `Content-Type: application/pdf; name="${pdfName}"`,
        `Content-Transfer-Encoding: base64`,
        `Content-Disposition: attachment; filename="${pdfName}"`,
        ``,
        // Split base64 into 76-char lines (MIME standard)
        ...pdfBase64.match(/.{1,76}/g),
        ``,
        `--${boundary}--`
    ].join('\r\n');
}

function corsResponse(statusCode, body) {
    return {
        statusCode,
        headers: {
            'Content-Type': 'application/json',
            'Access-Control-Allow-Origin': '*',
            'Access-Control-Allow-Methods': 'POST, OPTIONS',
            'Access-Control-Allow-Headers': 'Content-Type, x-api-key'
        },
        body: typeof body === 'string' ? body : JSON.stringify(body)
    };
}
