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
      for (let stepIndex=0;stepIndex<section.value.length;stepIndex+=1){const step=section.value[stepIndex];children.push(new Paragraph({keepNext:true,spacing:{before:300,after:120},children:[new TextRun({text:`Step ${stepIndex+1}: ${step.title||'Procedure step'}`,bold:true,color:'000000',size:28})]}));if(lines(step.instruction).length)children.push(...lines(step.instruction).map(instruction=>new Paragraph({spacing:{after:140,line:320},children:[new TextRun({text:instruction,color:'000000',size:24})]})));const shots=step.screenshots||[];for(let imageIndex=0;imageIndex<shots.length;imageIndex+=1){const shot=shots[imageIndex],figure=`Figure ${stepIndex+1}${String.fromCharCode(65+imageIndex)}`,limits=shot.size==='small'?[280,200]:shot.size==='large'?[540,390]:[410,295],dimensions=await getImageDimensions(shot.data,limits[0],limits[1]);children.push(new Paragraph({alignment:AlignmentType.CENTER,spacing:{before:140,after:60},children:[new ImageRun({data:await imageBytes(shot.data),transformation:dimensions})]}),new Paragraph({alignment:AlignmentType.CENTER,spacing:{after:80},children:[new TextRun({text:figure,bold:true,color:'17365D',size:19}),new TextRun({text:shot.caption?`: ${shot.caption}`:'',italics:true,color:'666666',size:18})]}));if((shot.description||'').trim())children.push(new Paragraph({alignment:AlignmentType.CENTER,spacing:{after:90},children:[new TextRun({text:shot.description,color:'475569',size:19})]}));const calls=[...new Set(((shot.editorState||{}).objects||[]).filter(o=>o.type==='callout').map(o=>Number(o.number)).filter(Number.isFinite))].sort((a,b)=>a-b).map(number=>({number,label:String((shot.calloutLabels||{})[number]||'').trim()})).filter(item=>item.label);if(calls.length){children.push(new Paragraph({spacing:{before:60,after:40},children:[new TextRun({text:'Callout legend',bold:true,color:'17365D',size:20})]}));for(const item of calls)children.push(new Paragraph({spacing:{after:50},children:[new TextRun({text:`${item.number}. ${item.label}`,size:19})]}));}}if((step.check||'').trim())children.push(new Paragraph({spacing:{before:100,after:180,line:300},shading:{fill:'EAF7F1',type:ShadingType.CLEAR},children:[new TextRun({text:'Completion check: ',bold:true,color:'087A55',size:22}),new TextRun({text:step.check,color:'000000',size:22})]}));}
    }
  }
  const screenshotRows=[];const stepSection=model.sections.find(section=>section.kind==='steps');if(stepSection)stepSection.value.forEach((step,si)=>(step.screenshots||[]).forEach((shot,ii)=>screenshotRows.push([`Figure ${si+1}${String.fromCharCode(65+ii)}`,`Step ${si+1}`,shot.caption||shot.description||''])));if(screenshotRows.length){children.push(new Paragraph({heading:HeadingLevel.HEADING_1,spacing:{before:300,after:120},children:[new TextRun({text:'Screenshot Index',bold:true,color:'17365D',size:30})]}));children.push(new Table({width:{size:100,type:WidthType.PERCENTAGE},rows:[new TableRow({children:['Figure','Step','Caption or description'].map(v=>new TableCell({shading:{fill:'D9EAF7',type:ShadingType.CLEAR},children:[new Paragraph({children:[new TextRun({text:v,bold:true,size:20})]})]}))}),...screenshotRows.map(row=>new TableRow({children:row.map(v=>new TableCell({children:[new Paragraph({children:[new TextRun({text:String(v),size:19})]})]}))}))]}));}

  const wordDocument = new Document({
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
