export async function exportWord(model) {
  const docx = await import('https://cdn.jsdelivr.net/npm/docx@9.7.1/+esm');
  const {
    AlignmentType,
    PageBreak,
    Document,
    HeadingLevel,
    ImageRun,
    Packer,
    Paragraph,
    ShadingType,
    Table,
    TableCell,
    TableRow,
    TableOfContents,
    ExternalHyperlink,
    TextRun,
    WidthType
  } = docx;

  const lines = (text) => String(text || '')
    .split(/\n+/)
    .map((item) => item.trim())
    .filter(Boolean);

  const figureMap = new Map((model.figureCatalog || []).map((item) => [item.shotId, item.figureId]));
  const resolveFigureReferences = (text) => String(text || '').replace(/\{\{FIG:([^}]+)\}\}/g, (match, id) => figureMap.get(id) || 'Figure reference unavailable');
  const safeName = (text) => (text || 'SOP')
    .replace(/[^a-z0-9_-]+/gi, '_');

  async function imageBytes(dataUrl) {
    return new Uint8Array(await (await fetch(dataUrl)).arrayBuffer());
  }

  function getImageDimensions(dataUrl, maxWidth = 480, maxHeight = 340) {
    return new Promise((resolve, reject) => {
      const image = new Image();

      image.onload = () => {
        const width = image.naturalWidth || maxWidth;
        const height = image.naturalHeight || maxHeight;
        const scale = Math.min(maxWidth / width, maxHeight / height, 1);

        resolve({
          width: Math.max(1, Math.round(width * scale)),
          height: Math.max(1, Math.round(height * scale))
        });
      };

      image.onerror = () => reject(new Error('The screenshot could not be read.'));
      image.src = dataUrl;
    });
  }

  const children = [];
  const createdDate = model.createdAt ? new Date(model.createdAt).toLocaleDateString() : 'Not provided';
  const updatedDate = model.updatedAt ? new Date(model.updatedAt).toLocaleDateString() : 'Not provided';

  children.push(
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 1100, after: 260 }, children: [new TextRun({ text: model.title || 'Untitled Procedure', bold: true, color: '17365D', size: 44 })] }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 180 }, children: [new TextRun({ text: model.category || 'General SOP', bold: true, color: '666666', size: 24 })] }),
    new Table({
      width: { size: 80, type: WidthType.PERCENTAGE },
      alignment: AlignmentType.CENTER,
      rows: [
        ['Owner', model.owner || 'Not provided'],
        ['Prepared by', model.prepared || 'Not provided'],
        ['Version', model.version || '1.0'],
        ['Status', model.status || 'Current'],
        ['Created', createdDate],
        ['Last updated', updatedDate],
        ['Revision notes', model.revisionNote || 'Initial procedure'],
        ['Version family', model.lifecycleSummary ? String(model.lifecycleSummary.totalVersions) + ' total version(s)' : '1 version'],
        ['Archived versions', model.lifecycleSummary ? String(model.lifecycleSummary.archivedVersions) : '0']
      ].map((row) => new TableRow({ children: row.map((value, index) => new TableCell({ shading: index === 0 ? { fill: 'D9EAF7', type: ShadingType.CLEAR } : undefined, children: [new Paragraph({ children: [new TextRun({ text: value, bold: index === 0, size: 22 })] })] })) }))
    }),
    new Paragraph({ children: [new PageBreak()] }),
    new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun({ text: 'Table of Contents', bold: true, color: '17365D', size: 30 })] }),
    new TableOfContents('Contents', { hyperlink: true, headingStyleRange: '1-2' }),
    new Paragraph({ children: [new PageBreak()] })
  );

  children.push(
    new Paragraph({
      heading: HeadingLevel.TITLE,
      spacing: { after: 100 },
      children: [
        new TextRun({
          text: model.title || 'Untitled Procedure',
          bold: true,
          color: '17365D',
          size: 40
        })
      ]
    }),
    new Paragraph({
      spacing: { after: 220 },
      children: [
        new TextRun({
          text: 'Internal Use | Draft for review',
          italics: true,
          color: '666666',
          size: 20
        })
      ]
    })
  );

  const metadataRows = [
    ['Owner', model.owner || 'Not provided', 'Version', model.version || 'Not provided'],
    ['Systems', model.systems || 'Not provided', 'Frequency', model.frequency || 'Not provided']
  ];

  if ((model.prepared || '').trim()) {
    metadataRows.push(['Prepared by', model.prepared, 'Status', 'Draft for review']);
  }

  children.push(
    new Table({
      width: { size: 100, type: WidthType.PERCENTAGE },
      rows: metadataRows.map((row) =>
        new TableRow({
          children: row.map((value, index) =>
            new TableCell({
              shading: index % 2 === 0
                ? { fill: 'D9EAF7', type: ShadingType.CLEAR }
                : undefined,
              children: [
                new Paragraph({
                  spacing: { before: 40, after: 40 },
                  children: [
                    new TextRun({
                      text: value,
                      bold: index % 2 === 0,
                      color: '000000',
                      size: 20
                    })
                  ]
                })
              ]
            })
          )
        })
      )
    })
  );

  const stepSectionForReferences = (model.sections.find((section) => section.kind === 'steps') || {}).value || [];
  for (let sectionIndex = 0; sectionIndex < model.sections.length; sectionIndex += 1) {
    const section = model.sections[sectionIndex];

    children.push(
      new Paragraph({
        heading: HeadingLevel.HEADING_1,
        spacing: { before: 300, after: 120 },
        children: [
          new TextRun({
            text: `${sectionIndex + 1}. ${section.title}`,
            bold: true,
            color: '17365D',
            size: 30
          })
        ]
      })
    );

    if (section.kind === 'text') {
      children.push(
        new Paragraph({
          spacing: { after: 140, line: 300 },
          children: [
            new TextRun({
              text: resolveFigureReferences(section.value),
              color: '000000',
              size: 24
            })
          ]
        })
      );
    }

    if (section.kind === 'list') {
      children.push(
        ...section.value.map((item) =>
          new Paragraph({
            bullet: { level: 0 },
            spacing: { after: 100, line: 300 },
            children: [
              new TextRun({
                text: resolveFigureReferences(item),
                color: '000000',
                size: 24
              })
            ]
          })
        )
      );
    }

    if (section.kind === 'revision') {
      const revisions = Array.isArray(section.value) ? section.value : (model.revisionHistory || []);
      if (revisions.length) {
        children.push(
          new Table({
            width: { size: 100, type: WidthType.PERCENTAGE },
            rows: [
              new TableRow({
                children: ['Version', 'Date', 'Change', 'Status'].map((value) => new TableCell({
                  shading: { fill: 'D9EAF7', type: ShadingType.CLEAR },
                  children: [new Paragraph({ children: [new TextRun({ text: value, bold: true, size: 20 })] })]
                }))
              }),
              ...revisions.map((item) => new TableRow({
                children: [
                  item.version || '',
                  item.date ? new Date(item.date).toLocaleDateString() : '',
                  item.note || '',
                  item.status || ''
                ].map((value) => new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: String(value), size: 20 })] })] }))
              }))
            ]
          })
        );
      }
    }
    if (section.kind === 'steps') {
      for (let stepIndex = 0; stepIndex < section.value.length; stepIndex += 1) {
        const step = section.value[stepIndex];
        children.push(new Paragraph({ heading: HeadingLevel.HEADING_2, keepNext: true, spacing: { before: 300, after: 120 }, children: [new TextRun({ text: `Step ${stepIndex + 1}: ${step.title || 'Procedure step'}`, bold: true, color: '000000', size: 28 })] }));
        if (lines(step.instruction).length) children.push(...lines(step.instruction).map((instruction) => new Paragraph({ spacing: { after: 140, line: 320 }, children: [new TextRun({ text: resolveFigureReferences(instruction), color: '000000', size: 24 })] })));
        for (const ref of (step.figureRefs || [])) {
          const normalizedRef = typeof ref === 'string' ? { shotId: ref, includeImage: false } : ref;
          const figure = (model.figureCatalog || []).find((item) => item.shotId === normalizedRef.shotId);
          if (!figure) continue;
          children.push(new Paragraph({ spacing: { after: 100 }, children: [new TextRun({ text: `See ${figure.figureId}.`, size: 22, color: '17365D' })] }));
          if (normalizedRef.includeImage) {
            let referencedShot = null;
            for (const modelStep of (stepSectionForReferences || [])) {
              referencedShot = (modelStep.screenshots || []).find((shot) => shot.id === normalizedRef.shotId);
              if (referencedShot) break;
            }
            if (referencedShot) {
              const limits = referencedShot.size === 'small' ? [280, 200] : referencedShot.size === 'large' ? [540, 390] : [410, 295];
              const dimensions = await getImageDimensions(referencedShot.data, limits[0], limits[1]);
              children.push(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 80, after: 50 }, children: [new ImageRun({ data: await imageBytes(referencedShot.data), transformation: dimensions })] }));
              children.push(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 60 }, children: [new TextRun({ text: figure.figureId, bold: true, color: '17365D', size: 19 }), new TextRun({ text: referencedShot.caption ? `: ${referencedShot.caption}` : '', italics: true, color: '666666', size: 18 })] }));
              if ((referencedShot.description || '').trim()) children.push(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 70 }, children: [new TextRun({ text: referencedShot.description.trim(), color: '475569', size: 19 })] }));
              const referencedCallouts = [...new Set((((referencedShot.editorState || {}).objects) || []).filter((object) => object.type === 'callout').map((object) => Number(object.number)).filter(Number.isFinite))].sort((a, b) => a - b).map((number) => ({ number, label: String((referencedShot.calloutLabels || {})[number] || '').trim() })).filter((item) => item.label);
              if (referencedCallouts.length) {
                children.push(new Paragraph({ spacing: { before: 40, after: 30 }, children: [new TextRun({ text: 'Callout legend', bold: true, color: '17365D', size: 20 })] }));
                for (const callout of referencedCallouts) children.push(new Paragraph({ spacing: { after: 40 }, children: [new TextRun({ text: `${callout.number}. ${callout.label}`, size: 19 })] }));
              }
            }
          }
        }
        const shots = Array.isArray(step.screenshots) ? step.screenshots : (step.image ? [{ data: step.image, caption: step.caption || '', size: 'medium' }] : []);
        const prepared = [];
        for (let imageIndex = 0; imageIndex < shots.length; imageIndex += 1) { const shot=shots[imageIndex],limits=shot.size==='small'?[280,200]:shot.size==='large'?[540,390]:[410,295]; prepared.push({shot,imageIndex,dimensions:await getImageDimensions(shot.data,limits[0],limits[1])}); }
        const figureName = (imageIndex) => {
          let n = imageIndex + 1;
          let letter = '';
          while (n > 0) {
            n -= 1;
            letter = String.fromCharCode(65 + (n % 26)) + letter;
            n = Math.floor(n / 26);
          }
          return `Figure ${stepIndex + 1}${letter}`;
        };
        const metadataParagraphs = (item) => {
          const shot = item.shot;
          const figure = figureName(item.imageIndex);
          const result = [
            new Paragraph({
              alignment: AlignmentType.CENTER,
              spacing: { after: 70 },
              children: [
                new TextRun({ text: figure, bold: true, color: '17365D', size: 19 }),
                new TextRun({ text: shot.caption ? `: ${shot.caption}` : '', italics: true, color: '666666', size: 18 })
              ]
            })
          ];
          if ((shot.description || '').trim()) {
            result.push(new Paragraph({
              alignment: AlignmentType.CENTER,
              spacing: { after: 80 },
              children: [new TextRun({ text: shot.description.trim(), color: '475569', size: 19 })]
            }));
          }
          const callouts = [...new Set(
            ((shot.editorState || {}).objects || [])
              .filter((object) => object.type === 'callout')
              .map((object) => Number(object.number))
              .filter(Number.isFinite)
          )]
            .sort((a, b) => a - b)
            .map((number) => ({ number, label: String((shot.calloutLabels || {})[number] || '').trim() }))
            .filter((item) => item.label);
          if (callouts.length) {
            result.push(new Paragraph({
              spacing: { before: 50, after: 35 },
              children: [new TextRun({ text: 'Callout legend', bold: true, color: '17365D', size: 20 })]
            }));
            for (const callout of callouts) {
              result.push(new Paragraph({
                spacing: { after: 45 },
                children: [new TextRun({ text: `${callout.number}. ${callout.label}`, size: 19 })]
              }));
            }
          }
          return result;
        };
        const columns = step.imageLayout === 'two-column' || (step.imageLayout === 'auto' && prepared.length > 1);
        if (columns) {
          for (let i = 0; i < prepared.length; i += 2) {
            const batch = prepared.slice(i, i + 2);
            const cells = [];
            for (const item of batch) {
              cells.push(new TableCell({
                children: [
                  new Paragraph({ alignment: AlignmentType.CENTER, children: [new ImageRun({ data: await imageBytes(item.shot.data), transformation: item.dimensions })] }),
                  ...metadataParagraphs(item)
                ]
              }));
            }
            if (cells.length === 1) cells.push(new TableCell({ children: [new Paragraph('')] }));
            children.push(new Table({ width: { size: 100, type: WidthType.PERCENTAGE }, rows: [new TableRow({ children: cells })] }));
          }
        } else {
          for (const item of prepared) {
            children.push(
              new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 140, after: 60 }, children: [new ImageRun({ data: await imageBytes(item.shot.data), transformation: item.dimensions })] }),
              ...metadataParagraphs(item)
            );
          }
        }
        if((step.check||'').trim())children.push(new Paragraph({spacing:{before:100,after:180,line:300},shading:{fill:'EAF7F1',type:ShadingType.CLEAR},children:[new TextRun({text:'Completion check: ',bold:true,color:'087A55',size:22}),new TextRun({text:resolveFigureReferences(step.check),color:'000000',size:22})]}));
      }
    }
  }

  const figures = model.figureCatalog || [];
  if (figures.length) {
    children.push(new Paragraph({ heading: HeadingLevel.HEADING_1, spacing: { before: 300, after: 120 }, children: [new TextRun({ text: 'Screenshot Index', bold: true, color: '17365D', size: 30 })] }));
    children.push(new Table({ width: { size: 100, type: WidthType.PERCENTAGE }, rows: [
      new TableRow({ children: ['Figure', 'Step', 'Caption or description'].map((value) => new TableCell({ shading: { fill: 'D9EAF7', type: ShadingType.CLEAR }, children: [new Paragraph({ children: [new TextRun({ text: value, bold: true, size: 20 })] })] })) }),
      ...figures.map((item) => new TableRow({ children: [item.figureId, `Step ${item.stepIndex + 1}`, item.caption || item.description || ''].map((value) => new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: String(value), size: 19 })] })] })) }))
    ] }));
  }
  const appendices = model.appendixSections || [];
  for (let i = 0; i < appendices.length; i += 1) {
    const appendix = appendices[i];
    const letter = String.fromCharCode(65 + i);
    children.push(new Paragraph({ heading: HeadingLevel.HEADING_1, spacing: { before: 300, after: 120 }, children: [new TextRun({ text: `Appendix ${letter}: ${appendix.title}`, bold: true, color: '17365D', size: 30 })] }));
    for (const item of appendix.items) {
      const runs = item.url ? [new ExternalHyperlink({ link: item.url, children: [new TextRun({ text: item.label, style: 'Hyperlink' })] })] : [new TextRun({ text: item.label, size: 22 })];
      children.push(new Paragraph({ bullet: { level: 0 }, children: runs }));
    }
  }
  const wordDocument = new Document({
    features: { updateFields: true },
    styles: {
      default: {
        document: {
          run: {
            font: 'Arial',
            size: 24,
            color: '000000'
          },
          paragraph: {
            spacing: { line: 300, after: 120 }
          }
        }
      }
    },
    sections: [
      {
        properties: {
          page: {
            margin: {
              top: 720,
              right: 720,
              bottom: 720,
              left: 720
            }
          }
        },
        children
      }
    ]
  });

  const blob = await Packer.toBlob(wordDocument);
  const downloadUrl = URL.createObjectURL(blob);
  const link = window.document.createElement('a');
  link.href = downloadUrl;
  link.download = `${safeName(model.title)}.docx`;
  window.document.body.appendChild(link);
  link.click();
  link.remove();
  setTimeout(() => URL.revokeObjectURL(downloadUrl), 1000);
}
