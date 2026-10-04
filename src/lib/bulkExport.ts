import type { OfficeFile } from '../types';
import { createOfficeFileBlob } from './fileIO';

/** Export every selected file in its native format, or reject the entire archive. */
export async function exportOfficeFilesZip(files: readonly OfficeFile[]): Promise<Blob> {
  if (!files.length) throw new Error('Select at least one file to download.');
  const { default: JSZip } = await import('jszip');
  const zip = new JSZip();
  const usedNames = new Set<string>();
  for (const file of files) {
    try {
      const { blob, filename } = await createOfficeFileBlob(file);
      const extensionAt = filename.lastIndexOf('.');
      const stem = filename.slice(0, extensionAt);
      const extension = filename.slice(extensionAt);
      let uniqueName = filename;
      for (let suffix = 2; usedNames.has(uniqueName.toLowerCase()); suffix++) uniqueName = `${stem} (${suffix})${extension}`;
      usedNames.add(uniqueName.toLowerCase());
      zip.file(uniqueName, await blob.arrayBuffer(), { createFolders: false });
    } catch (error) {
      const name = typeof file?.name === 'string' && file.name.trim() ? file.name : 'Untitled';
      const reason = error instanceof Error ? error.message : 'An unexpected export error occurred.';
      throw new Error(`Could not export “${name}”. ${reason} No ZIP archive was created.`, { cause: error });
    }
  }
  return zip.generateAsync({ type: 'blob', mimeType: 'application/zip', compression: 'DEFLATE', compressionOptions: { level: 6 } });
}
