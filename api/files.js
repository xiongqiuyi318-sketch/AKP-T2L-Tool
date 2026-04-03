import { list } from '@vercel/blob';
import { getBlobToken, sanitizeSegment, sendJson } from './_helpers.js';

export const config = { runtime: 'nodejs' };

export default async function handler(req, res) {
  if (req.method !== 'GET') {
    return sendJson(res, 405, { error: 'Method not allowed' });
  }

  try {
    const customer = sanitizeSegment(req.query.customer, 'general');
    const batchCode = req.query.batchCode ? sanitizeSegment(req.query.batchCode, '') : '';
    const prefix = `customers/${customer}/`;
    const { blobs } = await list({ prefix, limit: 1000, token: getBlobToken() });

    const normalized = blobs.map((blob) => ({
      pathname: blob.pathname,
      size: blob.size,
      uploadedAt: blob.uploadedAt
    }));

    const batches = Array.from(
      new Set(
        normalized
          .map((blob) => {
            const parts = String(blob.pathname || '').split('/');
            return parts.length >= 4 && !['inputs', 'outputs'].includes(parts[2]) ? parts[2] : '';
          })
          .filter(Boolean)
      )
    )
      .sort()
      .reverse()
      .slice(0, 5);

    const isLegacyPath = (pathname) => {
      const parts = String(pathname || '').split('/');
      return parts.length >= 4 && ['inputs', 'outputs'].includes(parts[2]);
    };

    const legacyFiles = normalized.filter((blob) => isLegacyPath(blob.pathname));
    const files = batchCode
      ? normalized.filter((blob) => String(blob.pathname).startsWith(`${prefix}${batchCode}/`))
      : [];

    return sendJson(res, 200, { customer, batchCode, batches, files, legacyFiles });
  } catch (error) {
    return sendJson(res, 500, { error: error.message || '读取文件列表失败' });
  }
}
