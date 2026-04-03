import { list } from '@vercel/blob';
import { sendJson } from './_helpers.js';

export const config = { runtime: 'nodejs' };

export default async function handler(req, res) {
  if (req.method !== 'GET') {
    return sendJson(res, 405, { error: 'Method not allowed' });
  }

  try {
    const { blobs } = await list({ prefix: 'customers/', limit: 1000 });
    const customers = Array.from(
      new Set(
        blobs
          .map((blob) => {
            const parts = String(blob.pathname || '').split('/');
            return parts.length >= 2 ? parts[1] : '';
          })
          .filter(Boolean)
      )
    ).sort();

    return sendJson(res, 200, { customers });
  } catch (error) {
    return sendJson(res, 500, { error: error.message || '读取客户失败' });
  }
}
