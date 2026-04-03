import { del } from '@vercel/blob';
import { appendHistory, getBlobToken, sanitizeSegment, sendJson } from './_helpers.js';

export const config = { runtime: 'nodejs' };

export default async function handler(req, res) {
  if (req.method !== 'POST') {
    return sendJson(res, 405, { error: 'Method not allowed' });
  }

  try {
    const { pathname, customer, operator } = req.body || {};
    if (!pathname || !String(pathname).startsWith('customers/')) {
      return sendJson(res, 400, { error: '非法文件路径' });
    }

    await del(pathname, { token: getBlobToken() });

    await appendHistory({
      action: 'delete',
      operator: String(operator || 'unknown'),
      customer: sanitizeSegment(customer || 'unknown'),
      fileName: pathname.split('/').pop(),
      path: pathname,
      time: new Date().toISOString()
    });

    return sendJson(res, 200, { ok: true });
  } catch (error) {
    return sendJson(res, 500, { error: error.message || '删除失败' });
  }
}
