import { list } from '@vercel/blob';
import { getBlobToken, sanitizeSegment, sendJson } from './_helpers.js';

export const config = { runtime: 'nodejs' };

export default async function handler(req, res) {
  if (req.method !== 'GET') {
    return sendJson(res, 405, { error: 'Method not allowed' });
  }

  try {
    const customer = sanitizeSegment(req.query.customer, 'general');
    const prefix = `customers/${customer}/`;
    const { blobs } = await list({ prefix, limit: 1000, token: getBlobToken() });
    const files = blobs.map((blob) => ({
      pathname: blob.pathname,
      size: blob.size,
      uploadedAt: blob.uploadedAt
    }));
    return sendJson(res, 200, { customer, files });
  } catch (error) {
    return sendJson(res, 500, { error: error.message || '读取文件列表失败' });
  }
}
