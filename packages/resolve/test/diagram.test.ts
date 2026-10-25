// SmartArt from other producers has no cached drawing: the data model's text is shown as an outline instead.
import { describe, expect, test } from 'vitest';
import { OpcPackage } from '@pptx/opc';
import { PresentationDocument } from '@pptx/model';
import { deckBytes, deckFiles } from './deck.ts';
import { zipFiles } from '../../opc/test/fixtures.ts';
import { resolveSlide } from '../src/index.ts';

const DGM = 'http://schemas.openxmlformats.org/drawingml/2006/diagram';
const pt = (id: string, text: string, type = '') => `<dgm:pt modelId="${id}"${type ? ` type="${type}"` : ''}><dgm:prSet/><dgm:spPr/>`
  + (text ? `<dgm:t><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US"/><a:t>${text}</a:t></a:r></a:p></dgm:t>` : '') + '</dgm:pt>';
const cxn = (id: string, src: string, dest: string, ord: number, type = '') => `<dgm:cxn modelId="${id}"${type ? ` type="${type}"` : ''} srcId="${src}" destId="${dest}" srcOrd="${ord}" destOrd="0"/>`;

function deckWithDiagram(): Uint8Array {
  const files = deckFiles();
  const data = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
    + `<dgm:dataModel xmlns:dgm="${DGM}" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><dgm:ptLst>`
    + pt('0', '', 'doc') + pt('1', 'Plan') + pt('2', 'Build') + pt('3', 'Parse') + pt('4', 'Render') + pt('5', 'Help', 'asst') + pt('9', 'ignored', 'parTrans')
    + '</dgm:ptLst><dgm:cxnLst>'
    + cxn('c1', '0', '2', 1) + cxn('c2', '0', '1', 0) + cxn('c3', '2', '4', 1) + cxn('c4', '2', '3', 0) + cxn('c5', '1', '5', 0) + cxn('c6', '1', '9', 1, 'presOf')
    + '</dgm:cxnLst></dgm:dataModel>';
  files['ppt/diagrams/data1.xml'] = data;
  files['[Content_Types].xml'] = files['[Content_Types].xml']!.replace('</Types>', '<Override PartName="/ppt/diagrams/data1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.diagramData+xml"/></Types>');
  files['ppt/slides/_rels/slide1.xml.rels'] = files['ppt/slides/_rels/slide1.xml.rels']!.replace('</Relationships>',
    '<Relationship Id="rId9" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramData" Target="../diagrams/data1.xml"/></Relationships>');
  files['ppt/slides/slide1.xml'] = files['ppt/slides/slide1.xml']!.replace('</p:spTree>',
    '<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="20" name="Diagram"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>'
    + '<p:xfrm><a:off x="0" y="0"/><a:ext cx="6000000" cy="4000000"/></p:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/diagram">'
    + `<dgm:relIds xmlns:dgm="${DGM}" r:dm="rId9" r:lo="rId9" r:qs="rId9" r:cs="rId9"/></a:graphicData></a:graphic></p:graphicFrame></p:spTree>`);
  return zipFiles(files);
}

describe('SmartArt without its cached drawing', () => {
  test('shows the data model as an outline: tree order, depth as level, transitions left out', () => {
    const doc = new PresentationDocument(OpcPackage.open(deckWithDiagram()));
    const frame = resolveSlide(doc, '/ppt/slides/slide1.xml').shapes.find(s => s.name === 'Diagram')!;
    expect(frame.graphic?.kind).toBe('diagram');
    const fallback = frame.graphic?.kind === 'diagram' ? frame.graphic.fallback : undefined;
    expect(fallback?.paragraphs.map(p => [p.props.level, p.runs.map(r => r.text).join('')])).toEqual([
      [0, 'Plan'], [1, 'Help'], [0, 'Build'], [1, 'Parse'], [1, 'Render'],
    ]);
    // bullets come from the master's body style, as a body placeholder's would
    expect(fallback?.paragraphs[0]!.props.bullet?.char).toBe('•');
  });

  test('a deck without diagrams is unaffected', () => {
    const doc = new PresentationDocument(OpcPackage.open(deckBytes()));
    expect(resolveSlide(doc, '/ppt/slides/slide1.xml').shapes.some(s => s.graphic?.kind === 'diagram')).toBe(false);
  });
});
