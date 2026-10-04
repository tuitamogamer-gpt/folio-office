import { Extension, Mark, Node, mergeAttributes } from '@tiptap/core';
import Image from '@tiptap/extension-image';

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
