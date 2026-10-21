export const NAMESPACES = {
  w: 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
  r: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
  a: 'http://schemas.openxmlformats.org/drawingml/2006/main',
  m: 'http://schemas.openxmlformats.org/officeDocument/2006/math',
  wp: 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
  v: 'urn:schemas-microsoft-com:vml',
  contentTypes: 'http://schemas.openxmlformats.org/package/2006/content-types',
  relationships: 'http://schemas.openxmlformats.org/package/2006/relationships',
} as const;

export const RELATIONSHIP_TYPES = {
  image: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image',
  header: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/header',
  footer: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer',
  numbering: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering',
  styles: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles',
  hyperlink: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink',
  footnotes: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes',
  endnotes: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes',
  comments: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments',
  fontTable: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable',
  theme: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme',
  settings: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings',
} as const;

export const CONTENT_TYPES = {
  document: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml',
  styles: 'application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml',
  numbering: 'application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml',
  header: 'application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml',
  footer: 'application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml',
} as const;

export const MIME_MAP: Record<string, string> = {
  png: 'image/png',
  jpeg: 'image/jpeg',
  jpg: 'image/jpeg',
  gif: 'image/gif',
  bmp: 'image/bmp',
  tiff: 'image/tiff',
  svg: 'image/svg+xml',
  emf: 'image/x-emf',
  wmf: 'image/x-wmf',
};

export const BUILT_IN_STYLES = new Set<string>([
  'Normal',
  'Heading1',
  'Heading2',
  'Heading3',
  'Heading4',
  'Heading5',
  'Heading6',
  'Heading7',
  'Heading8',
  'Heading9',
  'DefaultParagraphFont',
  'TableNormal',
  'NoList',
  'Title',
  'Subtitle',
  'BodyText',
  'ListParagraph',
  'Hyperlink',
  'FollowedHyperlink',
  'Header',
  'Footer',
]);
