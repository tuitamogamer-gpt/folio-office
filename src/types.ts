export type FileKind = 'document' | 'spreadsheet' | 'presentation';
export interface OfficeFile { id: string; name: string; kind: FileKind; updatedAt: number; createdAt: number; starred: boolean; trashed: boolean; content: any; color?: string; pageSetup?: { landscape: boolean; margin: 'normal' | 'narrow' | 'wide' }; }
export interface EditorProps { file: OfficeFile; onChange: (content: any) => void; onRename: (name: string) => void; onBack: () => void; onNotify: (message: string) => void; onPageSetupChange?: (setup: NonNullable<OfficeFile['pageSetup']>) => void; }
export interface SlideData { id: string; title: string; body: string; subtitle?: string; background: string; layout: 'title' | 'content' | 'split'; }
export interface SheetContent { cells: Record<string, string>; name: string; }
