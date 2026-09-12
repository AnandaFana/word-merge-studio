import JSZip from 'jszip';
import { createDemo } from './demo';
const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const REL = 'http://schemas.openxmlformats.org/package/2006/relationships';
// Public, fictional fixtures. Never derived from user documents.
export async function createSectionDemo(key: string): Promise<Uint8Array> {
  const zip = await JSZip.loadAsync(await createDemo('left'));
  const projects: Record<string, [string, string, string]> = {
    garden: [
      'Community Garden / 社区花园',
      'Plant seasonal herbs and share tools. 种植香草，共享园艺工具。',
      'Georgia',
    ],
    library: [
      'Reading Room / 社区书屋',
      'Create a quiet reading space and exchange books. 提供阅读与图书交换空间。',
      'Arial',
    ],
    cafe: [
      'Neighbourhood Cafe / 街角咖啡',
      'Host monthly neighbourhood gatherings. 每月举办邻里交流活动。',
      'Calibri',
    ],
  };
  const [name, goal, font] = projects[key] ?? projects.garden;
  const p = (s: string, level = 0) =>
    `<w:p><w:pPr><w:pStyle w:val="${level ? 'Heading' + level : 'Normal'}"/></w:pPr><w:r><w:t>${s}</w:t></w:r></w:p>`;
  const table = `<w:tbl><w:tblPr><w:tblW w:w="8000" w:type="dxa"/><w:tblBorders>${['top', 'left', 'bottom', 'right', 'insideH', 'insideV'].map((s) => `<w:${s} w:val="single" w:sz="4" w:color="AABBAA"/>`).join('')}</w:tblBorders></w:tblPr><w:tblGrid><w:gridCol w:w="4000"/><w:gridCol w:w="4000"/></w:tblGrid>${[
    ['Stage / 阶段', 'Output / 产出'],
    ['Prototype / 原型', key + ' pilot'],
    ['Review / 评审', 'Public feedback / 公开反馈'],
  ]
    .map(
      (row) =>
        `<w:tr>${row.map((s) => `<w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr>${p(s)}</w:tc>`).join('')}</w:tr>`,
    )
    .join('')}</w:tbl>`;
  const drawing = `<w:p><w:r><w:drawing><wp:inline><wp:extent cx="2286000" cy="952500"/><wp:docPr id="1" name="${key} illustration" descr="Three green bars representing project stages"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr><pic:cNvPr id="1" name="stages.png"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="image1"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="2286000" cy="952500"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>`;
  zip.file(
    'word/document.xml',
    `<w:document xmlns:w="${W}" xmlns:r="${R}" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><w:body>${p('1. Overview / 概述', 1)}${p(name)}${p('1.1 Purpose / 定位', 2)}${p(goal)}${drawing}${p('图 1 项目阶段示意 / Project stages')}${p('2. Goals / 目标', 1)}${p('2.1 Outcomes / 预期成果', 2)}${p('Publish a practical guide for local volunteers. 为志愿者提供实用指南。')}${p('3. Schedule / 计划', 1)}${p('3.1 Milestones / 里程碑', 2)}${table}<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1200" w:right="1200" w:bottom="1200" w:left="1200"/></w:sectPr></w:body></w:document>`,
  );
  zip.file(
    'word/styles.xml',
    `<w:styles xmlns:w="${W}"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="${font}" w:hAnsi="${font}" w:eastAsia="${key === 'library' ? '微软雅黑' : '宋体'}"/><w:sz w:val="24"/></w:rPr></w:rPrDefault></w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:pPr><w:spacing w:line="300" w:after="100"/></w:pPr></w:style>${[1, 2].map((l) => `<w:style w:type="paragraph" w:styleId="Heading${l}"><w:name w:val="Heading ${l}"/><w:basedOn w:val="Normal"/><w:pPr><w:keepNext/><w:outlineLvl w:val="${l - 1}"/></w:pPr><w:rPr><w:b/><w:color w:val="${key === 'cafe' ? '855630' : '315B45'}"/><w:sz w:val="${l === 1 ? 32 : 28}"/></w:rPr></w:style>`).join('')}</w:styles>`,
  );
  zip.file(
    'word/_rels/document.xml.rels',
    `<Relationships xmlns="${REL}"><Relationship Id="styles" Type="${R}/styles" Target="styles.xml"/><Relationship Id="image1" Type="${R}/image" Target="media/stages.png"/></Relationships>`,
  );
  const types = await zip.file('[Content_Types].xml')!.async('string');
  zip.file(
    '[Content_Types].xml',
    types.replace(
      '</Types>',
      '<Default Extension="png" ContentType="image/png"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>',
    ),
  );
  zip.file(
    'word/media/stages.png',
    'iVBORw0KGgoAAAANSUhEUgAAAPAAAABkCAIAAAA3wCqQAAABw0lEQVR4nO3csU3DQBiAUQcxEFvQRaKkYw6aKA1z0FEi0bEFi0Sp6GmpaHLI5y/vDWD9sj7dXWHf7vx9WqDiZu0BYCRBkyJoUgRNiqBJETQpgiZF0KQImhRBkyJoUgRNiqBJETQpgiZF0KQImhRBkyJoUgRNiqBJETQpgiZF0KQImhRBkyJoUgRNiqBJuV17AFawPzzN894/jq8Dn7Zzne5V2c+U8n9k7chBiqCvyH7W5XngbIImRdCkCJoUQZMiaFIETYqgSRE0KYImRdCkCJoUQZMiaFIETYqgSRE0Kf4p/Mvh7XmZxvHxZe0RNkDQG0j590iy/psjBymC3sbyvInZZiBoUgRNiqBJETQpgiZF0KQImhRBkyJoUgRNiqBJETQpgiZF0KQImhRBkzLmF6zPr/dlGvd3D2uPwGaDnirl3yPJ+jo5cpByE1ueNzEb/8cKTYqgSRE0KYImRdCkCJoUQZMiaFIETYqgSbko6Jk/ALpktpkvFb9kto/j6zKrUbNZoUnZnb9PsS+BBu4bU90uPnDf2B+elmmM3TfGBA2TcOQgRdCkCJoUQZMiaFIETYqgSRE0KYImRdCkCJoUQbOU/ACj+DzrZGxgLAAAAABJRU5ErkJggg==',
    { base64: true },
  );
  return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
}
