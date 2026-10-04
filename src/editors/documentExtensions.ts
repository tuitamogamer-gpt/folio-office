import { Extension, Mark, Node, mergeAttributes } from '@tiptap/core';
import Image from '@tiptap/extension-image';
import type { Node as ProseMirrorNode } from '@tiptap/pm/model';

const wordSegments = new Intl.Segmenter(undefined, { granularity: 'word' });
const characterSegments = new Intl.Segmenter(undefined, { granularity: 'grapheme' });

export interface WritingStatistics { words: number; characters: number; charactersWithoutSpaces: number; paragraphs: number; readingMinutes: number; }
export interface OutlineHeading { text: string; level: number; pos: number; }

function statisticsForParagraphs(paragraphs: string[]): WritingStatistics {
  const text = paragraphs.join('\n');
  let words = 0;
  let characters = 0;
  let charactersWithoutSpaces = 0;
  for (const segment of wordSegments.segment(text)) if (segment.isWordLike) words++;
  for (const segment of characterSegments.segment(text)) {
    if (!/[\r\n]/.test(segment.segment)) characters++;
    if (!/^\s+$/u.test(segment.segment)) charactersWithoutSpaces++;
  }
  return {
    words,
    characters,
    charactersWithoutSpaces,
    paragraphs: paragraphs.filter(paragraph => paragraph.trim()).length,
    readingMinutes: words ? words < 200 ? 0.5 : Math.ceil(words / 200) : 0,
  };
}

const writingCache = new WeakMap<ProseMirrorNode, { headings: OutlineHeading[]; statistics: WritingStatistics }>();

export function documentWritingState(doc: ProseMirrorNode, from: number, to: number) {
  let documentState = writingCache.get(doc);
  if (!documentState) {
    const headings: OutlineHeading[] = [];
    const paragraphs: string[] = [];
    doc.descendants((node, pos) => {
      if (node.type.name === 'heading') headings.push({ text: node.textContent, level: Number(node.attrs.level) || 1, pos });
      if (!node.isTextblock) return;
      paragraphs.push(node.textBetween(0, node.content.size, '', '\n'));
      return false;
    });
    documentState = { headings, statistics: statisticsForParagraphs(paragraphs) };
    writingCache.set(doc, documentState);
  }
  const selectedParagraphs: string[] = [];
  if (from < to) doc.nodesBetween(from, to, (node, pos) => {
    if (!node.isTextblock) return;
    const start = pos + 1;
    if (from < start + node.content.size && to > start) selectedParagraphs.push(node.textBetween(Math.max(0, from - start), Math.min(node.content.size, to - start), '', '\n'));
    return false;
  });
  return { ...documentState, selectionStatistics: from < to ? statisticsForParagraphs(selectedParagraphs) : null };
}

export const DocumentAttributes = Extension.create({
  name: 'documentAttributes',
  addGlobalAttributes() {
    return [
      {
        types: ['paragraph', 'heading'],
        attributes: {
          indent: {
            default: 0,
            parseHTML: element => Math.min(8, Math.max(0, Number(element.getAttribute('data-indent')) || Math.round(parseFloat(element.style.marginLeft || '0') / 36))),
            renderHTML: attributes => attributes.indent ? { 'data-indent': attributes.indent, style: `margin-left: ${attributes.indent * 36}pt` } : {},
          },
        },
      },
      {
        types: ['tableCell', 'tableHeader'],
        attributes: {
          backgroundColor: {
            default: null,
            parseHTML: element => element.style.backgroundColor || null,
            renderHTML: attributes => attributes.backgroundColor ? { style: `background-color: ${attributes.backgroundColor}` } : {},
          },
        },
      },
    ];
  },
});

export const PageBreak = Node.create({
  name: 'pageBreak',
  group: 'block',
  atom: true,
  selectable: true,
  parseHTML: () => [{ tag: 'div[data-page-break]' }],
  renderHTML: () => ['div', { 'data-page-break': 'true', class: 'doc-page-break', style: 'break-after: page; page-break-after: always', 'aria-label': 'Page break' }],
  addKeyboardShortcuts() {
    return { 'Mod-Enter': () => this.editor.chain().insertContent([{ type: 'pageBreak' }, { type: 'paragraph' }]).run() };
  },
});

export const CommentAnchor = Mark.create({
  name: 'commentAnchor',
  inclusive: false,
  addAttributes() {
    return { commentId: { default: null, parseHTML: element => element.getAttribute('data-comment-id'), renderHTML: attributes => ({ 'data-comment-id': attributes.commentId }) } };
  },
  parseHTML: () => [{ tag: 'span[data-comment-id]' }],
  renderHTML: ({ HTMLAttributes }) => ['span', mergeAttributes(HTMLAttributes, { class: 'doc-comment-anchor' }), 0],
});

export const DocumentImage = Image.extend({
  addAttributes() {
    return {
      ...this.parent?.(),
      width: {
        default: null,
        parseHTML: element => parseFloat(element.getAttribute('width') || element.style.width) || null,
        renderHTML: attributes => attributes.width ? { width: attributes.width, style: `width: ${attributes.width}px; max-width: 100%` } : {},
      },
      align: {
        default: 'left',
        parseHTML: element => element.getAttribute('data-align') || 'left',
        renderHTML: attributes => ({ 'data-align': attributes.align, style: `display: block; margin-left: ${attributes.align === 'left' ? '0' : 'auto'}; margin-right: ${attributes.align === 'right' ? '0' : 'auto'}` }),
      },
    };
  },
}).configure({ allowBase64: true });
