import path from 'path';
import { parseZip, ConvertPPI } from '../../packages/docx-view-engine/file';
import DocViewer from './DocViewer';
import { XMLParser } from 'fast-xml-parser';

export interface DocumentPageSettings {
  width: number;
  height: number;
  pageIdentation: DocumentPageIndentation;
}

export interface DocumentPageIndentation {
  top: number;
  bottom: number;
  right: number;
  left: number;
  header: number;
  footer: number;
}

const toArray = (x: any): any[] => (x == null ? [] : Array.isArray(x) ? x : [x]);

const isOn = (tag: any) => {
  if (tag === undefined) return false;
  const v = tag?.['@_w:val'];
  return !(v === '0' || v === 'false' || v === 'off');
};

// numbering.xml -> карта "numId:ilvl" => формат ("bullet", "decimal", ...)
function parseNumbering(numberingXml: string, parser: XMLParser) {
  const formats = new Map<string, string>();
  if (!numberingXml) return formats;

  const numbering = parser.parse(numberingXml)['w:numbering'] ?? {};

  const abstracts = new Map<string, any[]>();
  for (const a of toArray(numbering['w:abstractNum'])) {
    abstracts.set(a['@_w:abstractNumId'], toArray(a['w:lvl']));
  }

  for (const n of toArray(numbering['w:num'])) {
    const numId = n['@_w:numId'];
    const abstractId = n['w:abstractNumId']?.['@_w:val'];
    for (const lvl of abstracts.get(String(abstractId)) ?? []) {
      formats.set(`${numId}:${lvl['@_w:ilvl']}`, lvl['w:numFmt']?.['@_w:val'] ?? 'bullet');
    }
  }
  return formats;
}

export default function Home() {
  const filePath = path.join(process.cwd(), 'document2.docx');
  const files = parseZip(filePath);

  const documentXml = files.get('word/document.xml')?.toString('utf8') ?? '';
  const stylesXml = files.get('word/styles.xml')?.toString('utf8') ?? '';
  const numberingXml = files.get('word/numbering.xml')?.toString('utf8') ?? '';

  const parser = new XMLParser({ ignoreAttributes: false, attributeNamePrefix: '@_' });
  const main = parser.parse(documentXml);
  const result = main['w:document']['w:body'];

  const numFormats = parseNumbering(numberingXml, parser);

  const sectPr = result['w:sectPr'] ?? {};
  const pgSz = sectPr['w:pgSz'] ?? {};
  const pgMar = sectPr['w:pgMar'] ?? {};

  const widthPage = ConvertPPI(Number(pgSz['@_w:w'] ?? 11906));
  const heightPage = ConvertPPI(Number(pgSz['@_w:h'] ?? 16838));
  const top = ConvertPPI(Number(pgMar['@_w:top'] ?? 1440));
  const right = ConvertPPI(Number(pgMar['@_w:right'] ?? 1440));
  const bottom = ConvertPPI(Number(pgMar['@_w:bottom'] ?? 1440));
  const left = ConvertPPI(Number(pgMar['@_w:left'] ?? 1440));
  const header = ConvertPPI(Number(pgMar['@_w:header'] ?? 720));
  const footer = ConvertPPI(Number(pgMar['@_w:footer'] ?? 720));

  const counters = new Map<string, number>();

  const paragraphs = toArray(result['w:p']);

  const paragraphsData = paragraphs.map((p: any) => {
    const pPr = p?.['w:pPr'] ?? {};

    let list: { marker: string; level: number } | undefined;
    const numPr = pPr['w:numPr'];
    if (numPr) {
      const numId = String(numPr['w:numId']?.['@_w:val'] ?? '');
      const level = Number(numPr['w:ilvl']?.['@_w:val'] ?? 0);

      if (numId !== '0') {
        const fmt = numFormats.get(`${numId}:${level}`) ?? 'bullet';
        let marker = '•';

        if (fmt !== 'bullet' && fmt !== 'none') {
          const key = `${numId}:${level}`;
          const next = (counters.get(key) ?? 0) + 1;
          counters.set(key, next);
          for (let l = level + 1; l < 9; l++) counters.delete(`${numId}:${l}`);

          if (fmt === 'lowerLetter') marker = `${String.fromCharCode(96 + next)}.`;
          else if (fmt === 'upperLetter') marker = `${String.fromCharCode(64 + next)}.`;
          else marker = `${next}.`;
        }
        list = { marker, level };
      }
    }

    const spacing = pPr['w:spacing'] ?? {};
    const spaceBefore =
      spacing['@_w:beforeAutospacing'] === '1' || spacing['@_w:beforeAutospacing'] === 1
        ? '1em'
        : spacing['@_w:before'] !== undefined
          ? `${Number(spacing['@_w:before']) / 20}pt`
          : undefined;
    const spaceAfter =
      spacing['@_w:afterAutospacing'] === '1' || spacing['@_w:afterAutospacing'] === 1
        ? '1em'
        : spacing['@_w:after'] !== undefined
          ? `${Number(spacing['@_w:after']) / 20}pt`
          : undefined;

    const runs = toArray(p?.['w:r'])
      .filter((r: any) => r && r['w:t'] !== undefined)
      .map((r: any) => {
        const t = r['w:t'];
        const rPr = r['w:rPr'] ?? {};
        const sz = rPr['w:sz']?.['@_w:val'];
        const color = rPr['w:color']?.['@_w:val'];
        const underline = rPr['w:u']?.['@_w:val'];

        return {
          text: String(typeof t === 'object' ? t['#text'] ?? '' : t ?? ''),
          font: rPr['w:rFonts']?.['@_w:ascii'],
          bold: isOn(rPr['w:b']),
          italic: isOn(rPr['w:i']),
          strike: isOn(rPr['w:strike']),
          underline: underline !== undefined && underline !== 'none',
          size: sz ? Number(sz) / 2 : undefined,
          color: color && color !== 'auto' ? `#${color}` : undefined,
        };
      });

    return {
      text: runs.map((r) => r.text).join(''),
      runs,
      list,
      spaceBefore,
      spaceAfter,
    };
  });

  const text = paragraphsData.map((p) => p.text).join('\n');

  const docxSettings: DocumentPageSettings = {
    width: widthPage,
    height: heightPage,
    pageIdentation: { top, bottom, right, left, header, footer },
  };

  return (
    <DocViewer
      documentXml={documentXml}
      stylesXml={stylesXml}
      docxSettings={docxSettings}
      text={text}
      textStyle={paragraphsData}
    />
  );
}