'use client'
import { DocumentPageSettings } from "./page";

interface Props {
  documentXml: string;
  stylesXml: string;
  docxSettings: DocumentPageSettings
  text: string
}

export default function DocViewer({ documentXml, stylesXml, docxSettings, text }: Props) {
  console.log(text)
  return (
    <>
      <div style={{ padding: '10px', display: 'flex', justifyContent: 'center' }}>
        <div style={{
          background: 'white',
          width: docxSettings.width + 'px',
          height: docxSettings.height + 'px',
          paddingTop: docxSettings.pageIdentation.top + 'px',
          paddingBottom: docxSettings.pageIdentation.bottom + 'px',
          paddingLeft: docxSettings.pageIdentation.left + 'px',
          paddingRight: docxSettings.pageIdentation.right + 'px',
        }} className='page'>
          <div className='interation' style={{color: 'black'}}>
            {text}
          </div>
        </div>
      </div>
    </>
  );
}