// =============================================
// Lambda: crm-backend
// Runtime: Node.js 20.x
// Timeout: 60s | Memory: 256MB
// =============================================
// Variables de entorno requeridas:
//   FROM_EMAIL       - email verificado en SES
//   API_KEY          - clave para autenticar requests
//   APPS_SCRIPT_URL  - URL del Google Apps Script (para tracking)
// =============================================

import { SESClient, SendRawEmailCommand } from '@aws-sdk/client-ses';

const ses = new SESClient();
const FROM_EMAIL = process.env.FROM_EMAIL;
const API_KEY = process.env.API_KEY;
const APPS_SCRIPT_URL = process.env.APPS_SCRIPT_URL || '';
const BATCH_SIZE = 5;

// 1x1 transparent GIF in base64
const PIXEL_GIF = Buffer.from('R0lGODlhAQABAIAAAAAAAP///yH5BAEAAAAALAAAAAABAAEAAAIBRAA7', 'base64');

export const handler = async (event) => {
    const method = event.requestContext?.http?.method;
    const path = event.rawPath || event.requestContext?.http?.path || '';

    // CORS preflight
    if (method === 'OPTIONS') {
        return corsResponse(200, '');
    }

    // Route: GET /track-pixel (no auth needed - embedded in emails)
    if (method === 'GET' && path === '/track-pixel') {
        return handleTrackPixel(event);
    }

    // Route: POST /send-pdf (auth required)
    if (method === 'POST' && path === '/send-pdf') {
        // Validate API Key
        const apiKey = event.headers?.['x-api-key'] || event.headers?.['X-Api-Key'];
        if (apiKey !== API_KEY) {
            return corsResponse(401, { error: 'No autorizado' });
        }
        return handleSendPdf(event);
    }

    return corsResponse(404, { error: 'Not found' });
};

// ========================================
// TRACK PIXEL (email open tracking)
// ========================================
async function handleTrackPixel(event) {
    const params = event.queryStringParameters || {};
    const email = params.email;
    const claseId = params.claseId;

    // Log to Apps Script (must await before Lambda exits)
    if (email && claseId && APPS_SCRIPT_URL) {
        try {
            await fetch(APPS_SCRIPT_URL, {
                method: 'POST',
                headers: { 'Content-Type': 'text/plain' },
                body: JSON.stringify({
                    action: 'logTracking',
                    email: email,
                    claseId: claseId,
                    tipo: 'email-open'
                })
            });
        } catch (e) {
            console.error('Track pixel log error:', e);
        }
    }

    // Return 1x1 transparent GIF
    return {
        statusCode: 200,
        headers: {
            'Content-Type': 'image/gif',
            'Cache-Control': 'no-store, no-cache, must-revalidate, proxy-revalidate',
            'Pragma': 'no-cache',
            'Expires': '0',
            'Access-Control-Allow-Origin': '*'
        },
        body: PIXEL_GIF.toString('base64'),
        isBase64Encoded: true
    };
}

// ========================================
// SEND PDF EMAIL
// ========================================
async function handleSendPdf(event) {
    let body;
    try {
        body = typeof event.body === 'string' ? JSON.parse(event.body) : event.body;
    } catch {
        return corsResponse(400, { error: 'JSON invalido' });
    }

    const { recipients, pdfBase64, pdfName, attachments, subject, htmlBody, senderName, claseId } = body;

    // Build attachments array (support both old single-file and new multi-file format)
    const pdfAttachments = attachments || (pdfBase64 ? [{ base64: pdfBase64, name: pdfName }] : []);

    // Validate required fields
    if (!recipients?.length) return corsResponse(400, { error: 'recipients es requerido' });
    if (pdfAttachments.length === 0) return corsResponse(400, { error: 'Se requiere al menos un archivo adjunto' });
    if (!subject) return corsResponse(400, { error: 'subject es requerido' });
    if (!htmlBody) return corsResponse(400, { error: 'htmlBody es requerido' });

    const fromHeader = senderName ? `"${senderName}" <${FROM_EMAIL}>` : FROM_EMAIL;

    // Tracking pixel config
    const apiBaseUrl = process.env.API_BASE_URL || '';

    let sent = 0;
    let failed = 0;
    const errors = [];

    // Send in batches
    for (let i = 0; i < recipients.length; i += BATCH_SIZE) {
        const batch = recipients.slice(i, i + BATCH_SIZE);
        const promises = batch.map(async (recipient) => {
            try {
                let personalizedHtml = htmlBody.replace(/\{\{nombre\}\}/g, recipient.nombre || 'Estudiante');

                // Inject tracking pixel if claseId and API base URL are set
                if (claseId && apiBaseUrl) {
                    const pixelUrl = `${apiBaseUrl}/track-pixel?email=${encodeURIComponent(recipient.email)}&claseId=${encodeURIComponent(claseId)}`;
                    personalizedHtml += `<img src="${pixelUrl}" width="1" height="1" style="display:block;width:1px;height:1px;border:0;" alt="" />`;
                }

                const rawEmail = buildMimeEmail({
                    from: fromHeader,
                    to: recipient.email,
                    subject,
                    html: personalizedHtml,
                    attachments: pdfAttachments
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
}

// ========================================
// MIME EMAIL BUILDER
// ========================================
function buildMimeEmail({ from, to, subject, html, attachments }) {
    const boundary = `----=_Part_${Date.now()}_${Math.random().toString(36).slice(2)}`;

    const parts = [
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
    ];

    for (const att of attachments) {
        parts.push(
            ``,
            `--${boundary}`,
            `Content-Type: application/pdf; name="${att.name}"`,
            `Content-Transfer-Encoding: base64`,
            `Content-Disposition: attachment; filename="${att.name}"`,
            ``,
            ...att.base64.match(/.{1,76}/g),
        );
    }

    parts.push(``, `--${boundary}--`);
    return parts.join('\r\n');
}

// ========================================
// CORS RESPONSE
// ========================================
function corsResponse(statusCode, body) {
    return {
        statusCode,
        headers: {
            'Content-Type': 'application/json',
            'Access-Control-Allow-Origin': '*',
            'Access-Control-Allow-Methods': 'POST, GET, OPTIONS',
            'Access-Control-Allow-Headers': 'Content-Type, x-api-key'
        },
        body: typeof body === 'string' ? body : JSON.stringify(body)
    };
}
