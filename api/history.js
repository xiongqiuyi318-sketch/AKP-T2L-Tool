import { list } from '@vercel/blob';
import { sendJson } from './_helpers.js';

export const config = { runtime: 'nodejs' };

export default async function handler(req, res) {
  if (req.method !== 'GET') {
    return sendJson(res, 405, { error: 'Method not allowed' });
  }

  try {
    const limit = Math.max(1, Math.min(100, Number(req.query.limit || 20)));
    const { blobs } = await list({ prefix: 'system/history.json', limit: 1 });
    if (!blobs.length) return sendJson(res, 200, { records: [] });

    const response = await fetch(blobs[0].url);
    if (!response.ok) return sendJson(res, 200, { records: [] });
    const all = await response.json();
    const records = Array.isArray(all) ? all.slice(0, limit) : [];
    return sendJson(res, 200, { records });
  } catch (error) {
    return sendJson(res, 500, { error: error.message || '读取历史失败' });
  }
}
