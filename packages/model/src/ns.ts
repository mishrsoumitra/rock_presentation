// packages/model/src/ns.ts  -- namespaces and relationship types, Transitional and Strict

/** Logical namespace keys. Transitional and Strict URIs of the same vocabulary map to the same key, so a view
 *  written against 'a' reads a Strict file too. The first URI of each entry is the one we write. */
export const NAMESPACES = {
  a: ['http://schemas.openxmlformats.org/drawingml/2006/main', 'http://purl.oclc.org/ooxml/drawingml/main'],
  p: ['http://schemas.openxmlformats.org/presentationml/2006/main', 'http://purl.oclc.org/ooxml/presentationml/main'],
  r: ['http://schemas.openxmlformats.org/officeDocument/2006/relationships', 'http://purl.oclc.org/ooxml/officeDocument/relationships'],
  c: ['http://schemas.openxmlformats.org/drawingml/2006/chart', 'http://purl.oclc.org/ooxml/drawingml/chart'],
  cdr: ['http://schemas.openxmlformats.org/drawingml/2006/chartDrawing', 'http://purl.oclc.org/ooxml/drawingml/chartDrawing'],
  dgm: ['http://schemas.openxmlformats.org/drawingml/2006/diagram', 'http://purl.oclc.org/ooxml/drawingml/diagram'],
  pic: ['http://schemas.openxmlformats.org/drawingml/2006/picture', 'http://purl.oclc.org/ooxml/drawingml/picture'],
  m: ['http://schemas.openxmlformats.org/officeDocument/2006/math', 'http://purl.oclc.org/ooxml/officeDocument/math'],
  mc: ['http://schemas.openxmlformats.org/markup-compatibility/2006'],
  p14: ['http://schemas.microsoft.com/office/powerpoint/2010/main'],
  p15: ['http://schemas.microsoft.com/office/powerpoint/2012/main'],
  p159: ['http://schemas.microsoft.com/office/powerpoint/2015/09/main'],
  p16: ['http://schemas.microsoft.com/office/powerpoint/2015/main'],
  a14: ['http://schemas.microsoft.com/office/drawing/2010/main'],
  a16: ['http://schemas.microsoft.com/office/drawing/2014/main'],
  asvg: ['http://schemas.microsoft.com/office/drawing/2016/SVG/main'],
  dsp: ['http://schemas.microsoft.com/office/drawing/2008/diagram'],
  cx: ['http://schemas.microsoft.com/office/drawing/2014/chartex'],
  c14: ['http://schemas.microsoft.com/office/drawing/2007/8/2/chart'],
  v: ['urn:schemas-microsoft-com:vml'],
  o: ['urn:schemas-microsoft-com:office:office'],
} as const;

export type NsKey = keyof typeof NAMESPACES;

const BY_URI = new Map<string, NsKey>();
for (const [key, uris] of Object.entries(NAMESPACES) as [NsKey, readonly string[]][]) for (const u of uris) BY_URI.set(u, key);

/** The logical key of a namespace URI, or undefined for namespaces this model does not know. */
export const nsKey = (uri: string): NsKey | undefined => BY_URI.get(uri);

/** True when the URI is the Strict form of a vocabulary. */
export const isStrictUri = (uri: string) => uri.startsWith('http://purl.oclc.org/ooxml/');

/** The URI to write for a key, in the flavour (Transitional or Strict) of a document. */
export const uriFor = (key: NsKey, strict = false): string => {
  const uris: readonly string[] = NAMESPACES[key];
  return strict && uris[1] ? uris[1] : uris[0]!;
};

/**
 * Relationship types are compared by their last path segment, which is the same in both flavours
 * (".../officeDocument/2006/relationships/slide" and ".../ooxml/officeDocument/relationships/slide"),
 * and in Microsoft's own namespace (".../office/2007/relationships/diagramDrawing").
 */
export const relKind = (type: string) => type.slice(type.lastIndexOf('/') + 1);

const REL_BASE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/';
const REL_BASE_STRICT = 'http://purl.oclc.org/ooxml/officeDocument/relationships/';
/** The full relationship type to write for a kind, in the flavour of the document. */
export const relType = (kind: string, strict = false) => (strict ? REL_BASE_STRICT : REL_BASE) + kind;

export const REL = {
  officeDocument: 'officeDocument',
  slide: 'slide',
  slideLayout: 'slideLayout',
  slideMaster: 'slideMaster',
  notesSlide: 'notesSlide',
  notesMaster: 'notesMaster',
  handoutMaster: 'handoutMaster',
  theme: 'theme',
  themeOverride: 'themeOverride',
  image: 'image',
  hyperlink: 'hyperlink',
  chart: 'chart',
  chartUserShapes: 'chartUserShapes',
  diagramData: 'diagramData',
  diagramLayout: 'diagramLayout',
  diagramQuickStyle: 'diagramQuickStyle',
  diagramColors: 'diagramColors',
  diagramDrawing: 'diagramDrawing',
  oleObject: 'oleObject',
  package: 'package',
  video: 'video',
  audio: 'audio',
  media: 'media',
  tableStyles: 'tableStyles',
  presProps: 'presProps',
  viewProps: 'viewProps',
  tags: 'tags',
  font: 'font',
  vmlDrawing: 'vmlDrawing',
  comments: 'comments',
} as const;

export const CONTENT_TYPE = {
  presentation: 'application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml',
  template: 'application/vnd.openxmlformats-officedocument.presentationml.template.main+xml',
  slideshow: 'application/vnd.openxmlformats-officedocument.presentationml.slideshow.main+xml',
  macroPresentation: 'application/vnd.ms-powerpoint.presentation.macroEnabled.main+xml',
  macroTemplate: 'application/vnd.ms-powerpoint.template.macroEnabled.main+xml',
  macroSlideshow: 'application/vnd.ms-powerpoint.slideshow.macroEnabled.main+xml',
  slide: 'application/vnd.openxmlformats-officedocument.presentationml.slide+xml',
  slideLayout: 'application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml',
  slideMaster: 'application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml',
  notesSlide: 'application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml',
  theme: 'application/vnd.openxmlformats-officedocument.theme+xml',
} as const;
