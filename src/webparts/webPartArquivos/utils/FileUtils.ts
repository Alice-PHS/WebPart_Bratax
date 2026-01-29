import JSZip from 'jszip';

export const calculateHash = async (file: File | Blob): Promise<string> => {
  const arrayBuffer = await file.arrayBuffer();
  const hashBuffer = await crypto.subtle.digest('SHA-256', arrayBuffer);
  const hashArray = Array.from(new Uint8Array(hashBuffer));
  const hashHex = hashArray.map(b => b.toString(16).padStart(2, '0')).join('');
  return hashHex;
};

export const createZipPackage = async (files: File[]): Promise<Blob> => {
  const zip = new (JSZip as any)();
  files.forEach(f => zip.file(f.name, f));
  return await zip.generateAsync({ type: "blob" });
};