import { head } from '@vercel/blob';
import { sendJson } from './_helpers.js';

export const config = { runtime: 'nodejs' };

export default async function handler(req, res) {
  if (req.method !== 'GET') {
    return sendJson(res, 405, { error: 'Method not allowed' });
  }

  try {
    const pathname = String(req.query.pathname || '').trim();
    if (!pathname || !pathname.startsWith('customers/')) {
      return sendJson(res, 400, { error: '非法文件路径' });
    }

    const meta = await head(pathname, { token: process.env.BLOB_READ_WRITE_TOKEN });
    return sendJson(res, 200, {
      pathname: meta.pathname,
      downloadUrl: meta.downloadUrl
    });
  } catch (error) {
    return sendJson(res, 500, { error: error.message || '获取下载链接失败' });
  }
}
