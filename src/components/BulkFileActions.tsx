import { useEffect, useRef } from 'react';
import { Archive, RotateCcw, Star, Trash2, X } from 'lucide-react';
import './workspace-productivity.css';

export function SelectVisibleFiles({ count, total, onChange }: { count: number; total: number; onChange: () => void }) {
  const input = useRef<HTMLInputElement>(null);
  useEffect(() => { if (input.current) input.current.indeterminate = count > 0 && count < total; }, [count, total]);
  return <input ref={input} className="file-selection-checkbox" type="checkbox" aria-label="Select all visible files" checked={total > 0 && count === total} disabled={!total} onChange={onChange}/>;
}

interface Props { count: number; trashed: boolean; allStarred: boolean; busy: boolean; onClear: () => void; onDownload: () => void; onStar: () => void; onTrash: () => void; onRestore: () => void; onDelete: () => void; }
export default function BulkFileActions(props: Props) {
  if (!props.count) return null;
  return <div className="bulk-file-actions" role="toolbar" aria-label="Selected file actions">
    <div className="bulk-selection-count"><button className="icon-button" aria-label="Clear file selection" onClick={props.onClear}><X size={16}/></button><strong>{props.count} selected</strong></div>
    <div className="bulk-action-buttons">
      <button disabled={props.busy} onClick={props.onDownload}><Archive size={16}/><span>Download ZIP</span></button>
      {props.trashed ? <><button onClick={props.onRestore}><RotateCcw size={16}/><span>Restore</span></button><button className="danger" onClick={props.onDelete}><Trash2 size={16}/><span>Delete forever</span></button></> : <><button onClick={props.onStar}><Star size={16}/><span>{props.allStarred ? 'Unstar' : 'Star'}</span></button><button onClick={props.onTrash}><Trash2 size={16}/><span>Move to trash</span></button></>}
    </div>
  </div>;
}
