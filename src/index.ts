import { DocxMerger } from './core/docx-merger';

export { DocxMerger };
export default DocxMerger;

export * from './types';
export * from './utils/constants';
export { parseXml, serializeXml, getDirectChildren, walkElements } from './utils/xml';
export { ContentTypesManager } from './core/content-types';
export { NumberingManager } from './core/numbering';
export { StylesManager } from './core/styles';
export { RelationsManager } from './core/relations';
