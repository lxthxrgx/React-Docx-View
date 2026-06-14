import path from 'path';
import { parseZip, ConvertPPI } from '../../packages/docx-view-engine/file';
import DocViewer from './DocViewer';
import { XMLParser } from 'fast-xml-parser';

export interface DocumentPageSettings {
  width: number,
  height: number,
  pageIdentation: DocumentPageIndentation
}

export interface DocumentPageIndentation {
  top: number,
  bottom: number,
  right: number,
  left: number,
  header: number,
  footer: number
}

export default function Home() {
  const filePath = path.join(process.cwd(), 'document.docx');
  const files = parseZip(filePath);

  const documentXml = files.get('word/document.xml')?.toString('utf8') ?? '';
  const stylesXml = files.get('word/styles.xml')?.toString('utf8') ?? '';

  const parser = new XMLParser({ ignoreAttributes: false, attributeNamePrefix: '@_' });
  const main = parser.parse(documentXml);
  const result = main['w:document']['w:body']

  const sectPr = result['w:sectPr'];

  const widthPage = ConvertPPI(Number(sectPr['w:pgSz']['@_w:w']));
  const heightPage = ConvertPPI(Number(sectPr['w:pgSz']['@_w:h']));

  const top = ConvertPPI(Number(sectPr['w:pgMar']['@_w:top']));
  const right = ConvertPPI(Number(sectPr['w:pgMar']['@_w:right']));
  const bottom = ConvertPPI(Number(sectPr['w:pgMar']['@_w:bottom']));
  const left = ConvertPPI(Number(sectPr['w:pgMar']['@_w:left']));
  const header = ConvertPPI(Number(sectPr['w:pgMar']['@_w:header']));
  const footer = ConvertPPI(Number(sectPr['w:pgMar']['@_w:footer']));

  const paragraphs = Array.isArray(result['w:p']) ? result['w:p'] : [result['w:p']];

  const text = paragraphs.map((p: any) => {
    const runs = Array.isArray(p['w:r']) ? p['w:r'] : [p['w:r']];
    return runs
      .filter((r: any) => r && r['w:t'])
      .map((r: any) => {
        const t = r['w:t'];
        return typeof t === 'object' ? t['#text'] ?? '' : t ?? '';
      })
      .join('');
  }).join('\n');

  const docxSettings: DocumentPageSettings = {
    width: widthPage,
    height: heightPage,
    pageIdentation: {
      top,
      bottom,
      right,
      left,
      header,
      footer
    }
  };

  return <DocViewer
    documentXml={documentXml}
    stylesXml={stylesXml}
    docxSettings={docxSettings}
    text={text}
  />;
}