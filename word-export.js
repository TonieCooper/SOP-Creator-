export async function exportWord(model) {
  const docx = await import('https://cdn.jsdelivr.net/npm/docx@9.7.1/+esm');
  const {
    AlignmentType,
    Document,
    HeadingLevel,
    ImageRun,
    Packer,
    Paragraph,
    ShadingType,
    Table,
    TableCell,
    TableRow,
    TextRun,
    WidthType
  } = docx;

  const lines = (text) => String(text || '')
    .split(/\n+/)
    .map((item) => item.trim())
    .filter(Boolean);

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
              text: section.value,
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
                text: item,
                color: '000000',
                size: 24
              })
            ]
          })
        )
      );
    }

    if (section.kind === 'revision') {
      children.push(
        new Paragraph({
          spacing: { after: 120, line: 300 },
          children: [
            new TextRun({
              text: `${section.value || 'Draft 0.1'}: `,
              bold: true,
              color: '000000',
              size: 22
            }),
            new TextRun({
              text: 'Initial draft.',
              color: '000000',
              size: 22
            })
          ]
        })
      );
    }

    if (section.kind === 'steps') {
      for (let stepIndex = 0; stepIndex < section.value.length; stepIndex += 1) {
        const step = section.value[stepIndex];

        children.push(
          new Paragraph({
            keepNext: true,
            spacing: { before: 300, after: 120 },
            children: [
              new TextRun({
                text: `Step ${stepIndex + 1}: ${step.title || 'Procedure step'}`,
                bold: true,
                color: '000000',
                size: 28
              })
            ]
          })
        );

        if (lines(step.instruction).length) {
          children.push(
            ...lines(step.instruction).map((instruction) =>
              new Paragraph({
                spacing: { after: 140, line: 320 },
                children: [
                  new TextRun({
                    text: instruction,
                    color: '000000',
                    size: 24
                  })
                ]
              })
            )
          );
        }

        if (step.image) {
          const dimensions = await getImageDimensions(step.image);

          children.push(
            new Paragraph({
              alignment: AlignmentType.CENTER,
              spacing: { before: 160, after: 80 },
              children: [
                new ImageRun({
                  data: await imageBytes(step.image),
                  transformation: dimensions
                })
              ]
            }),
            new Paragraph({
              alignment: AlignmentType.CENTER,
              spacing: { after: 160 },
              children: [
                new TextRun({
                  text: step.caption || `Figure ${stepIndex + 1}. Screenshot supporting this step.`,
                  italics: true,
                  color: '666666',
                  size: 19
                })
              ]
            })
          );
        }

        if ((step.check || '').trim()) {
          children.push(
            new Paragraph({
              spacing: { before: 100, after: 180, line: 300 },
              shading: { fill: 'EAF7F1', type: ShadingType.CLEAR },
              children: [
                new TextRun({
                  text: 'Completion check: ',
                  bold: true,
                  color: '087A55',
                  size: 22
                }),
                new TextRun({
                  text: step.check,
                  color: '000000',
                  size: 22
                })
              ]
            })
          );
        }
      }
    }
  }

  const document = new Document({
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

  const blob = await Packer.toBlob(document);
  const link = document.createElement('a');
  link.href = URL.createObjectURL(blob);
  link.download = `${safeName(model.title)}.docx`;
  document.body.appendChild(link);
  link.click();
  link.remove();
  setTimeout(() => URL.revokeObjectURL(link.href), 1000);
}
