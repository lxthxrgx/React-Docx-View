'use client'
import { DocumentPageSettings } from "./page";

export interface RunData {
  text: string;
  font?: string;
  bold?: boolean;
  italic?: boolean;
  strike?: boolean;
  underline?: boolean;
  size?: number;
  color?: string;
}

export interface ParagraphData {
  text: string;
  runs: RunData[];
  list?: { marker: string; level: number };
  spaceBefore?: string;
  spaceAfter?: string;
}

interface Props {
  documentXml: string;
  stylesXml: string;
  docxSettings: DocumentPageSettings;
  text: string;
  textStyle: ParagraphData[];
}

export default function DocViewer({ docxSettings, textStyle }: Props) {
  return (
    <div style={{ padding: '10px', display: 'flex', justifyContent: 'center' }}>
      <div
        className="page"
        style={{
          background: 'white',
          width: docxSettings.width + 'px',
          height: docxSettings.height + 'px',
          paddingTop: docxSettings.pageIdentation.top + 'px',
          paddingBottom: docxSettings.pageIdentation.bottom + 'px',
          paddingLeft: docxSettings.pageIdentation.left + 'px',
          paddingRight: docxSettings.pageIdentation.right + 'px',
          boxSizing: 'border-box',
        }}
      >
        <div className="interation" style={{ color: 'black' }}>
          {textStyle.map((p, i) => (
            <p
              key={i}
              style={{
                margin: 0,
                marginTop: p.spaceBefore,
                marginBottom: p.spaceAfter,
                minHeight: '1em',
                display: p.list ? 'flex' : undefined,
                paddingLeft: p.list ? `${(p.list.level + 1) * 36}px` : undefined,
              }}
            >
              {p.list && (
                <span style={{ marginLeft: '-20px', width: '20px', flexShrink: 0 }}>
                  {p.list.marker}
                </span>
              )}
              <span>
                {p.runs.map((r, j) => (
                  <span
                    key={j}
                    style={{
                      fontFamily: r.font,
                      fontWeight: r.bold ? 700 : undefined,
                      fontStyle: r.italic ? 'italic' : undefined,
                      fontSize: r.size ? `${r.size}pt` : undefined,
                      color: r.color,
                      textDecoration:
                        [r.underline && 'underline', r.strike && 'line-through']
                          .filter(Boolean)
                          .join(' ') || undefined,
                    }}
                  >
                    {r.text}
                  </span>
                ))}
              </span>
            </p>
          ))}
        </div>
      </div>
    </div>
  );
}