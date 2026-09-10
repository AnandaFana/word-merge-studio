import JSZip from 'jszip';
import type { Side } from './engine';

const escape = (s: string) =>
  s.replaceAll('&', '&amp;').replaceAll('<', '&lt;').replaceAll('>', '&gt;');
export async function createDemo(side: Side): Promise<Uint8Array> {
  const zip = new JSZip();
  const lines =
    side === 'left'
      ? [
          '智能文档协作平台',
          '项目方案 · 讨论稿',
          '一、项目概述',
          '本项目旨在建立统一的文档协作平台，帮助团队完成文档整理、版本对比与内容归档。',
          '系统采用集中式部署，支持团队在办公网络内使用。',
          '二、建设目标',
          '第一阶段实现文档上传、全文检索和版本管理，预计于 6 月底完成。',
          '项目预算为 30 万元，实施周期为三个月。',
          '三、交付安排',
          '交付源代码、部署手册及用户操作指南。',
          '四、验收要求',
          '系统通过功能测试后组织验收。',
        ]
      : [
          '智能文档协作平台',
          '项目方案 · 修订稿',
          '一、项目概述',
          '本项目旨在建立统一的文档协作平台，帮助团队完成文档整理、版本对比、内容合并与知识归档。',
          '系统采用本地优先部署，支持团队在离线环境中使用。',
          '二、建设目标',
          '第一阶段实现文档上传、全文检索和版本管理，预计于 7 月底完成。',
          '项目预算为 35 万元，实施周期为四个月。',
          '三、交付安排',
          '交付源代码、部署手册及用户操作指南。',
          '增加两次用户培训，并提供六个月的技术支持。',
          '四、验收要求',
          '系统通过功能测试和格式保真测试后组织验收。',
        ];
  const paragraphs = lines
    .map(
      (s, i) =>
        `<w:p><w:pPr><w:spacing w:after="${i < 2 ? '180' : '120'}" w:line="360"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Calibri" w:eastAsia="${side === 'left' ? '宋体' : '微软雅黑'}"/>${i === 0 || /^[一二三四]、/.test(s) ? '<w:b/>' : ''}<w:sz w:val="${i === 0 ? '36' : '24'}"/></w:rPr><w:t>${escape(s)}</w:t></w:r></w:p>`,
    )
    .join('');
  zip.file(
    '[Content_Types].xml',
    '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>',
  );
  zip.file(
    '_rels/.rels',
    '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>',
  );
  zip.file(
    'word/document.xml',
    `<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${paragraphs}<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr></w:body></w:document>`,
  );
  return zip.generateAsync({ type: 'uint8array' });
}
