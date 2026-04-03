async function bufferToBase64(arrayBuffer) {
  let binary = '';
  const bytes = new Uint8Array(arrayBuffer);
  const chunkSize = 0x8000;
  for (let i = 0; i < bytes.length; i += chunkSize) {
    const chunk = bytes.subarray(i, i + chunkSize);
    binary += String.fromCharCode(...chunk);
  }
  return btoa(binary);
}

async function postJson(url, body) {
  const res = await fetch(url, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(body)
  });
  const text = await res.text();
  let data = {};
  if (text) {
    try {
      data = JSON.parse(text);
    } catch {
      throw new Error(`接口返回格式异常: ${url}`);
    }
  }
  if (!res.ok) throw new Error(data.error || '请求失败');
  return data;
}

export async function uploadFileToCloud({ file, customer, operator, fileType, kind = 'inputs' }) {
  const base64Data = await bufferToBase64(await file.arrayBuffer());
  return postJson('/api/upload', {
    customer,
    operator,
    fileType,
    kind,
    fileName: file.name,
    contentType: file.type || 'application/octet-stream',
    base64Data
  });
}

export async function uploadBufferToCloud({ buffer, filename, customer, operator, fileType }) {
  const base64Data = await bufferToBase64(buffer);
  return postJson('/api/upload', {
    customer,
    operator,
    fileType,
    kind: 'outputs',
    fileName: filename,
    contentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    base64Data
  });
}

export async function fetchCustomerFiles(customer) {
  const res = await fetch(`/api/files?customer=${encodeURIComponent(customer)}`);
  const text = await res.text();
  const data = text ? JSON.parse(text) : {};
  if (!res.ok) throw new Error(data.error || '读取客户文件失败');
  return data.files || [];
}

export async function fetchHistory(limit = 20) {
  const res = await fetch(`/api/history?limit=${limit}`);
  const text = await res.text();
  const data = text ? JSON.parse(text) : {};
  if (!res.ok) throw new Error(data.error || '读取历史失败');
  return data.records || [];
}

export async function getSignedDownloadUrl(pathname) {
  const res = await fetch(`/api/download?pathname=${encodeURIComponent(pathname)}`);
  const text = await res.text();
  const data = text ? JSON.parse(text) : {};
  if (!res.ok) throw new Error(data.error || '获取下载链接失败');
  return data.downloadUrl;
}

export async function deleteCloudFile({ pathname, customer, operator }) {
  const res = await fetch('/api/delete', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ pathname, customer, operator })
  });
  const text = await res.text();
  const data = text ? JSON.parse(text) : {};
  if (!res.ok) throw new Error(data.error || '删除失败');
  return data;
}
