import { put } from '@vercel/blob';
import { appendHistory, decodeBase64File, getBlobToken, sanitizeSegment, sendJson } from './_helpers.js';

export const config = { runtime: 'nodejs' };

export default async function handler(req, res) {
  if (req.method !== 'POST') {
    return sendJson(res, 405, { error: 'Method not allowed' });
  }

  try {
    const body = req.body || {};
    const customer = sanitizeSegment(body.customer, 'general');
    const batchCode = sanitizeSegment(body.batchCode, 'default-batch');
    const operator = String(body.operator || 'unknown').trim() || 'unknown';
    const fileType = sanitizeSegment(body.fileType, 'file');
    const kind = sanitizeSegment(body.kind, 'inputs');
    const fileName = String(body.fileName || `${fileType}.xlsx`).trim();
    const contentType =
      String(body.contentType || 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    const timestamp = new Date().toISOString().replace(/[:.]/g, '-');

    const buffer = decodeBase64File(body.base64Data);
    const path = `customers/${customer}/${batchCode}/${kind}/${timestamp}_${fileType}_${fileName}`;
    const uploaded = await put(path, buffer, {
      access: 'public',
      addRandomSuffix: false,
      allowOverwrite: false,
      contentType,
      token: getBlobToken()
    });

    await appendHistory({
      action: kind === 'outputs' ? 'generate' : 'upload',
      operator,
      customer,
      batchCode,
      fileType,
      fileName,
      path,
      url: uploaded.url,
      time: new Date().toISOString()
    });

    return sendJson(res, 200, { ok: true, file: uploaded, path });
  } catch (error) {
    return sendJson(res, 500, { error: error.message || '上传失败' });
  }
}
