// PATCHED VERSION
// Critical fix:
// const wordDocument = new Document(...)
// window.document.createElement(...)
// instead of document.createElement(...)

export async function exportWord(model) {
  const docx = await import('https://cdn.jsdelivr.net/npm/docx@9.7.1/+esm');
  const {
    Document, Packer, Paragraph, TextRun, HeadingLevel,
    Table, TableCell, TableRow, WidthType, ShadingType
  } = docx;

  const children = [
    new Paragraph({
      heading: HeadingLevel.TITLE,
      children: [new TextRun({ text: model.title || 'Untitled Procedure', bold: true })]
    })
  ];

  const wordDocument = new Document({
    sections: [{ children }]
  });

  const blob = await Packer.toBlob(wordDocument);

  const link = window.document.createElement('a');
  const downloadUrl = URL.createObjectURL(blob);

  link.href = downloadUrl;
  link.download = `${(model.title || 'SOP').replace(/[^a-z0-9_-]+/gi,'_')}.docx`;

  window.document.body.appendChild(link);
  link.click();
  link.remove();

  setTimeout(() => URL.revokeObjectURL(downloadUrl), 1000);
}
