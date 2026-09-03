// =============================================
// Lambda: backend de correo — Marketing Digital Avanzado (UAO)
// Runtime: Node.js 20.x
// Timeout: 60s | Memory: 256MB
// =============================================
// Variables de entorno requeridas:
//   FROM_EMAIL       - email verificado en SES
//   API_KEY          - token de administrador. Debe ser EXACTAMENTE el mismo
//                      valor que la Script Property ADMIN_TOKEN del Apps
//                      Script (se genera con tools/generar-token.html).
//   TRACKING_TOKEN   - secreto compartido con el Apps Script. Firma los
//                      pixeles de seguimiento y autentica logTracking.
//   APPS_SCRIPT_URL  - URL del Google Apps Script (para tracking)
//   API_BASE_URL     - URL base del API Gateway (para construir el pixel)
// =============================================

import { SESClient, SendRawEmailCommand } from '@aws-sdk/client-ses';
import { createHmac, timingSafeEqual } from 'node:crypto';

const ses = new SESClient();
const FROM_EMAIL = process.env.FROM_EMAIL;
const API_KEY = process.env.API_KEY;
const TRACKING_TOKEN = process.env.TRACKING_TOKEN || '';
const APPS_SCRIPT_URL = process.env.APPS_SCRIPT_URL || '';
const BATCH_SIZE = 5;

// Tipos de adjunto permitidos, por extension
const MIME_POR_EXT = {
    pdf: 'application/pdf',
    png: 'image/png',
    jpg: 'image/jpeg',
    jpeg: 'image/jpeg',
    doc: 'application/msword',
    docx: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    ppt: 'application/vnd.ms-powerpoint',
    pptx: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
    xls: 'application/vnd.ms-excel',
    xlsx: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
};

/**
 * Firma (email, claseId) con el secreto de tracking. Sin esto cualquiera
 * podia sumarse los puntos de "abrir el correo" con un simple GET.
 */
function firmarPixel(email, claseId) {
    if (!TRACKING_TOKEN) return '';
    return createHmac('sha256', TRACKING_TOKEN)
        .update(`${email}|${claseId}`)
        .digest('hex')
        .slice(0, 32);
}

function firmaValida(recibida, esperada) {
    if (!recibida || !esperada) return false;
    const a = Buffer.from(recibida);
    const b = Buffer.from(esperada);
    if (a.length !== b.length) return false;
    return timingSafeEqual(a, b);
}

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
        // Fail-closed: sin API_KEY configurada no se envia nada.
        if (!API_KEY || !firmaValida(apiKey, API_KEY)) {
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
    const sig = params.sig;

    // Solo se registra la apertura si la firma corresponde a este par
    // (email, claseId). Un GET fabricado a mano no otorga puntos.
    const firmaOk = email && claseId && firmaValida(sig, firmarPixel(email, claseId));

    if (firmaOk && APPS_SCRIPT_URL) {
        try {
            await fetch(APPS_SCRIPT_URL, {
                method: 'POST',
                headers: { 'Content-Type': 'text/plain' },
                body: JSON.stringify({
                    action: 'logTracking',
                    token: TRACKING_TOKEN,
                    email: email,
                    claseId: claseId,
                    tipo: 'email-open'
                })
            });
        } catch (e) {
            console.error('Track pixel log error:', e);
        }
    } else if (email || claseId) {
        console.warn('Track pixel con firma invalida', { claseId });
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
    if (pdfAttachments.some(a => !a || !a.base64)) return corsResponse(400, { error: 'Hay un adjunto vacio o invalido' });
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
                    const sig = firmarPixel(recipient.email, claseId);
                    const pixelUrl = `${apiBaseUrl}/track-pixel?email=${encodeURIComponent(recipient.email)}&claseId=${encodeURIComponent(claseId)}&sig=${sig}`;
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
        // Un base64 vacio devolvia null en .match() y reventaba el spread.
        const lineas = (att.base64 || '').match(/.{1,76}/g);
        if (!lineas) continue;

        const ext = (att.name || '').split('.').pop().toLowerCase();
        const mime = MIME_POR_EXT[ext] || 'application/octet-stream';
        // El nombre va dentro de comillas en la cabecera: fuera comillas y saltos.
        const nombre = (att.name || 'adjunto').replace(/["\r\n]/g, '');

        parts.push(
            ``,
            `--${boundary}`,
            `Content-Type: ${mime}; name="${nombre}"`,
            `Content-Transfer-Encoding: base64`,
            `Content-Disposition: attachment; filename="${nombre}"`,
            ``,
            ...lineas,
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
