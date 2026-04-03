import { list, put } from '@vercel/blob';

const HISTORY_BLOB_PATH = 'system/history.json';
const HISTORY_MAX_ITEMS = 200;

export function sendJson(res, status, payload) {
  res.status(status).json(payload);
}

export function sanitizeSegment(input, fallback = 'default') {
  const clean = String(input || '')
    .trim()
    .toLowerCase()
    .replace(/[^a-z0-9_-]+/g, '-')
    .replace(/^-+|-+$/g, '');
  return clean || fallback;
}

export function decodeBase64File(base64Data) {
  if (!base64Data || typeof base64Data !== 'string') {
    throw new Error('缺少文件数据');
  }
  return Buffer.from(base64Data, 'base64');
}

async function readHistory() {
  const { blobs } = await list({ prefix: HISTORY_BLOB_PATH, limit: 1 });
  if (!blobs.length) return [];
  const response = await fetch(blobs[0].url);
  if (!response.ok) return [];
  const parsed = await response.json();
  return Array.isArray(parsed) ? parsed : [];
}

export async function appendHistory(entry) {
  const current = await readHistory();
  const next = [entry, ...current].slice(0, HISTORY_MAX_ITEMS);
  await put(HISTORY_BLOB_PATH, JSON.stringify(next), {
    access: 'public',
    addRandomSuffix: false,
    allowOverwrite: true,
    contentType: 'application/json'
  });
}
