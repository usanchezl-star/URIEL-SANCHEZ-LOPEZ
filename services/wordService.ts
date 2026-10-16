
import { 
  Document, 
  Packer, 
  Paragraph, 
  TextRun, 
  Table, 
  TableRow, 
  TableCell, 
  WidthType, 
  AlignmentType, 
  HeadingLevel,
  BorderStyle,
  VerticalAlign,
  PageBreak,
  ImageRun,
  TabStopType,
  TabStopPosition,
  LeaderType,
  Footer,
  Header,
  ExternalHyperlink,
  BookmarkStart,
  BookmarkEnd,
  PageReference
} from "docx";
import * as FileSaver from "file-saver";
import { ERFDocument } from "../types";

export const downloadERFWord = async (docData: ERFDocument) => {
  const { identificacion, antecedentes, requerimiento, reglasNegocio, criteriosAceptacion, pasosProceso, metadata, controlCambios, anexos, hipervinculos } = docData;
  const saveAs = (FileSaver as any).saveAs || (FileSaver as any).default || FileSaver;

  const LIVERPOOL_PINK = "E61B7D";
  const BLACK = "000000";
  const SLATE_BG = "F8FAFC";

  const base64ToArrayBuffer = (base64String: string) => {
    try {
      const base64 = base64String.includes(",") ? base64String.split(",")[1] : base64String;
      const binaryString = window.atob(base64);
      const len = binaryString.length;
      const bytes = new Uint8Array(len);
      for (let i = 0; i < len; i++) bytes[i] = binaryString.charCodeAt(i);
      return bytes.buffer;
    } catch (e) {
      console.error("Error decoding image base64", e);
      return null;
    }
  };

  const createChapterHeader = (text: string, bookmarkId?: string, bookmarkNum?: number) => {
    return new Paragraph({
      shading: { fill: BLACK },
      spacing: { before: 600, after: 300 },
      children: [
        ...(bookmarkId && bookmarkNum !== undefined ? [new BookmarkStart(bookmarkId, bookmarkNum)] : []),
        new TextRun({
          text: text,
          bold: true,
          color: "FFFFFF",
          size: 28,
        }),
        ...(bookmarkNum !== undefined ? [new BookmarkEnd(bookmarkNum)] : []),
      ],
    });
  };

  const createSectionHeader = (text: string, bookmarkId?: string, bookmarkNum?: number) => {
    return new Paragraph({
      spacing: { before: 400, after: 200 },
      children: [
        ...(bookmarkId && bookmarkNum !== undefined ? [new BookmarkStart(bookmarkId, bookmarkNum)] : []),
        new TextRun({
          text: text,
          bold: true,
          color: LIVERPOOL_PINK,
          size: 24,
        }),
        ...(bookmarkNum !== undefined ? [new BookmarkEnd(bookmarkNum)] : []),
      ],
    });
  };

  const doc = new Document({
    sections: [
      {
        properties: {
          titlePage: true,
        },
        headers: {
          default: new Header({
            children: [
              new Paragraph({
                alignment: AlignmentType.CENTER,
                children: [
                  new TextRun({
                    text: "Eficiencia de Procesos C.C.",
                    color: "E2E8F0",
                    size: 40,
                    bold: true,
                  }),
                ],
              }),
            ],
          }),
          first: new Header({
            children: [
              new Paragraph({
                alignment: AlignmentType.CENTER,
                children: [
                  new TextRun({
                    text: "Eficiencia de Procesos C.C.",
                    color: "E2E8F0",
                    size: 40,
                    bold: true,
                  }),
                ],
              }),
              // Watermark only for the first page
              new Paragraph({
                children: [
                  new ImageRun({
                    type: "png",
                    data: await (async () => {
                      // Create a canvas to draw the watermark text
                      const canvas = document.createElement('canvas');
                      canvas.width = 800;
                      canvas.height = 600;
                      const ctx = canvas.getContext('2d');
                      if (ctx) {
                        ctx.translate(400, 300);
                        ctx.rotate(-Math.PI / 4);
                        ctx.font = 'bold 80px Arial';
                        ctx.fillStyle = 'rgba(200, 200, 200, 0.15)'; // Diluted color
                        ctx.textAlign = 'center';
                        ctx.textBaseline = 'middle';
                        ctx.fillText('CONFIDENCIAL', 0, 0);
                        return await new Promise<ArrayBuffer>((resolve) => {
                          canvas.toBlob((blob) => {
                            if (blob) {
                              const reader = new FileReader();
                              reader.onload = () => {
                                const base64 = (reader.result as string).split(',')[1];
                                const binaryString = window.atob(base64);
                                const len = binaryString.length;
                                const bytes = new Uint8Array(len);
                                for (let i = 0; i < len; i++) bytes[i] = binaryString.charCodeAt(i);
                                resolve(bytes.buffer);
                              };
                              reader.readAsDataURL(blob);
                            } else {
                              resolve(new ArrayBuffer(0));
                            }
                          });
                        });
                      }
                      return new ArrayBuffer(0);
                    })(),
                    transformation: {
                      width: 600,
                      height: 450,
                    },
                    floating: {
                      horizontalPosition: {
                        relative: "page",
                        align: AlignmentType.CENTER,
                      },
                      verticalPosition: {
                        relative: "page",
                        align: VerticalAlign.CENTER,
                      },
                      behindDocument: true,
                    },
                  }),
                ],
              }),
            ],
          }),
        },
        footers: {
          first: new Footer({
            children: [
              new Paragraph({
                border: { top: { style: BorderStyle.SINGLE, size: 6, color: "E2E8F0" } },
                spacing: { before: 200 },
                children: [
                  new TextRun({
                    text: "Documento Confidencial.",
                    bold: true,
                    size: 16,
                  }),
                ],
              }),
              new Paragraph({
                children: [
                  new TextRun({
                    text: "Derechos Reservados. El contenido y formato de este documento son propiedad de Servicios Liverpool, S. A. de C. V. y es confidencial. Cualquier reproducción del todo o en parte está estrictamente prohibida sin el permiso escrito de Servicios Liverpool, S. A. de C. V.",
                    size: 16,
                  }),
                ],
              }),
              new Paragraph({
                children: [
                  new TextRun({
                    text: "Este documento está sujeto a cambios. Comentarios, correcciones o preguntas deberán ser dirigidas al autor.",
                    size: 16,
                  }),
                ],
              }),
            ],
          }),
        },
        children: [
          // 1. PORTADA
          new Table({
            width: { size: 100, type: WidthType.PERCENTAGE },
            borders: {
              top: { style: BorderStyle.NONE },
              bottom: { style: BorderStyle.SINGLE, size: 20, color: BLACK },
              left: { style: BorderStyle.NONE },
              right: { style: BorderStyle.NONE },
              insideHorizontal: { style: BorderStyle.NONE },
              insideVertical: { style: BorderStyle.NONE },
            },
            rows: [
              new TableRow({
                children: [
                  new TableCell({
                    children: [
                      new Paragraph({
                        children: [
                          new TextRun({ text: "Liverpool", color: LIVERPOOL_PINK, bold: true, size: 36 }),
                          new TextRun({ text: "®", color: LIVERPOOL_PINK, size: 18, superScript: true })
                        ],
                      }),
                    ],
                  }),
                  new TableCell({
                    verticalAlign: VerticalAlign.CENTER,
                    children: [
                      new Paragraph({
                        alignment: AlignmentType.RIGHT,
                        children: [
                          new TextRun({ text: "MTi", bold: true, size: 32, italics: true, color: "FB923C" })
                        ],
                      }),
                    ],
                  }),
                ],
              }),
            ],
          }),

          new Paragraph({
            alignment: AlignmentType.CENTER,
            spacing: { before: 2500 },
            children: [
              new TextRun({ text: "Metodología de Trabajo de Informática", bold: true, size: 20, color: "64748B" }),
            ],
          }),
          new Paragraph({
            alignment: AlignmentType.CENTER,
            spacing: { before: 400 },
            children: [
              new TextRun({ text: "ESPECIFICACIÓN DE REQUERIMIENTO FUNCIONAL (ERF)", bold: true, size: 36 }),
            ],
          }),
          new Paragraph({
            alignment: AlignmentType.CENTER,
            shading: { fill: SLATE_BG },
            spacing: { before: 1200, after: 1200 },
            children: [
              new TextRun({ text: (identificacion.proceso || "Nuevo Proceso").toUpperCase(), bold: true, size: 40 }),
            ],
          }),
          new Paragraph({
            alignment: AlignmentType.CENTER,
            spacing: { before: 800 },
            children: [
              new TextRun({ text: `Versión: ${metadata.version}  |  Fecha: ${metadata.fecha}  |  Código: ${metadata.codigo}`, bold: true, size: 18, color: "94A3B8" }),
            ],
          }),
          new Paragraph({ children: [new PageBreak()] }),

          // 2. ÍNDICE
          new Paragraph({
            text: "ÍNDICE DE CONTENIDO",
            heading: HeadingLevel.HEADING_1,
            spacing: { after: 400 },
          }),
          ...[
            { text: "CAPÍTULO I: IDENTIFICACIÓN Y CONTROL", bookmark: "cap1", level: 0 },
            { text: "1.1 Identificación del Requerimiento", bookmark: "sec1", level: 1 },
            { text: "1.2 Control de Cambios", bookmark: "sec2", level: 1 },
            { text: "CAPÍTULO II: DEFINICIÓN DEL NEGOCIO", bookmark: "cap2", level: 0 },
            { text: "2.1 Antecedentes y Justificación", bookmark: "sec3", level: 1 },
            { text: "2.2 Objetivos del Requerimiento", bookmark: "sec3b", level: 1 },
            { text: "CAPÍTULO III: ESPECIFICACIÓN FUNCIONAL", bookmark: "cap3", level: 0 },
            { text: "3.1 Flujo de Proceso Detallado", bookmark: "sec4", level: 1 },
            { text: "3.2 Reglas de Negocio", bookmark: "sec5", level: 1 },
            { text: "3.3 Criterios de Aceptación", bookmark: "sec6", level: 1 },
            { text: "CAPÍTULO IV: COMPLEMENTOS", bookmark: "cap4", level: 0 },
            { text: "4.1 Hipervínculos y Enlaces", bookmark: "sec7", level: 1 },
            { text: "4.2 Anexos Visuales", bookmark: "sec8", level: 1 }
          ].map(item => new Paragraph({
            tabStops: [
              {
                type: TabStopType.RIGHT,
                position: TabStopPosition.MAX,
                leader: LeaderType.DOT,
              },
            ],
            indent: { left: item.level * 360 },
            children: [
              new TextRun({ 
                text: item.text, 
                size: 22, 
                bold: item.level === 0 
              }),
              new TextRun({ text: "\t", size: 22 }),
              new PageReference(item.bookmark),
            ],
            spacing: { before: 150 },
          })),
          new Paragraph({ children: [new PageBreak()] }),

          // CAPÍTULO I
          createChapterHeader("CAPÍTULO I: IDENTIFICACIÓN Y CONTROL", "cap1", 100),
          
          // 1.1 IDENTIFICACIÓN
          createSectionHeader("1.1 Identificación del Requerimiento", "sec1", 1),
          new Table({
            width: { size: 100, type: WidthType.PERCENTAGE },
            rows: [
              ["PEP / FOLIO REMEDY", identificacion.pepRemedy],
              ["SISTEMA / PLATAFORMA", identificacion.sistema],
              ["MÓDULO", identificacion.modulo],
              ["PROCESO", identificacion.proceso],
              ["DESCRIPCIÓN", identificacion.descripcion],
              ["SOLICITADO POR", identificacion.requeridoPor],
              ["FECHA SOLICITUD", identificacion.requeridoFecha],
              ["APROBADO POR", identificacion.aprobadoPor],
              ["FECHA APROBACIÓN", identificacion.aprobadoFecha]
            ].map(([label, value]) => (
              new TableRow({
                children: [
                  new TableCell({ width: { size: 35, type: WidthType.PERCENTAGE }, shading: { fill: SLATE_BG }, children: [new Paragraph({ children: [new TextRun({ text: label || "", bold: true, size: 18 })] })] }),
                  new TableCell({ width: { size: 65, type: WidthType.PERCENTAGE }, children: [new Paragraph({ children: [new TextRun({ text: value || "", size: 18 })] })] }),
                ]
              })
            ))
          }),

          // 1.2 CONTROL DE CAMBIOS
          createSectionHeader("1.2 Control de Cambios", "sec2", 2),
          new Table({
            width: { size: 100, type: WidthType.PERCENTAGE },
            rows: [
              new TableRow({
                children: ["VER.", "DESCRIPCIÓN", "AUTOR", "FECHA"].map(h => 
                  new TableCell({ shading: { fill: SLATE_BG }, children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: h, bold: true, size: 16 })] })] })
                )
              }),
              ...controlCambios.map(c => new TableRow({
                children: [c.version, c.descripcion, c.autor, c.fecha].map(val => 
                  new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: val || "", size: 16 })] })] })
                )
              }))
            ]
          }),
          new Paragraph({ children: [new PageBreak()] }),

          // CAPÍTULO II
          createChapterHeader("CAPÍTULO II: DEFINICIÓN DEL NEGOCIO", "cap2", 200),

          // 2.1 ANTECEDENTES
          createSectionHeader("2.1 Antecedentes y Justificación", "sec3", 3),
          new Paragraph({
            alignment: AlignmentType.BOTH,
            children: [new TextRun({ text: antecedentes || "No se proporcionaron antecedentes.", size: 22 })],
            spacing: { after: 400 }
          }),

          // 2.2 OBJETIVOS
          createSectionHeader("2.2 Objetivos del Requerimiento", "sec3b", 33),
          new Paragraph({ children: [new TextRun({ text: "Objetivo General", bold: true, color: LIVERPOOL_PINK, size: 20 })], spacing: { before: 200, after: 100 } }),
          new Paragraph({ children: [new TextRun({ text: requerimiento.objetivo || "", size: 22 })], spacing: { after: 300 } }),
          new Paragraph({ children: [new TextRun({ text: "Necesidad de Negocio", bold: true, color: LIVERPOOL_PINK, size: 20 })], spacing: { before: 200, after: 100 } }),
          new Paragraph({ children: [new TextRun({ text: requerimiento.necesidad || "", size: 22 })] }),
          new Paragraph({ children: [new PageBreak()] }),

          // CAPÍTULO III
          createChapterHeader("CAPÍTULO III: ESPECIFICACIÓN FUNCIONAL", "cap3", 300),

          // 3.1 FLUJO DE PROCESO
          createSectionHeader("3.1 Flujo de Proceso Detallado", "sec4", 4),
          ...pasosProceso.flatMap(p => {
            const stepBuffer = p.pantallaAsociada ? base64ToArrayBuffer(p.pantallaAsociada) : null;
            
            return [
              new Paragraph({
                spacing: { before: 400 },
                children: [
                  new TextRun({ text: `PASO ${p.paso} `, bold: true, size: 20 }),
                  ...((!p.ocultarDetalles && p.timestamp) ? [new TextRun({ text: `(Ref: ${p.timestamp})`, size: 16, color: "64748B" })] : [])
                ]
              }),
              new Table({
                width: { size: 100, type: WidthType.PERCENTAGE },
                rows: [
                  new TableRow({
                    children: [
                      new TableCell({
                        children: [
                          new Paragraph({ children: [new TextRun({ text: p.descripcion || "Sin descripción", size: 20 })], spacing: { after: 100 } }),
                          ...((!p.ocultarDetalles && p.resultadoEsperado) ? [
                            new Paragraph({
                              children: [
                                new TextRun({ text: "Resultado Esperado: ", bold: true, size: 16 }),
                                new TextRun({ text: p.resultadoEsperado, size: 16 })
                              ]
                            })
                          ] : []),
                          ...(stepBuffer ? [
                            new Paragraph({
                              alignment: AlignmentType.CENTER,
                              spacing: { before: 200 },
                              children: [
                                new ImageRun({
                                  data: stepBuffer as any,
                                  type: "jpg",
                                  transformation: { width: 450, height: 280 }
                                })
                              ]
                            }),
                            new Paragraph({
                              alignment: AlignmentType.CENTER,
                              children: [
                                new TextRun({
                                  text: `[CAPTURA_DE_PANTALLA: ${p.timestamp}]`,
                                  size: 14,
                                  color: "64748B",
                                  italics: true
                                })
                              ]
                            })
                          ] : [])
                        ]
                      })
                    ]
                  })
                ]
              })
            ];
          }),
          new Paragraph({ children: [new PageBreak()] }),

          // 7. REGLAS DE NEGOCIO
          createSectionHeader("3.2 Reglas de Negocio", "sec5", 5),
          ...(reglasNegocio.length > 0 
            ? reglasNegocio.map((rn, idx) => new Paragraph({
                spacing: { before: 200 },
                children: [
                  new TextRun({ text: `RN-${idx + 1}: `, bold: true, color: LIVERPOOL_PINK, size: 20 }),
                  new TextRun({ text: rn, size: 22 })
                ]
              }))
            : []
          ),
          new Paragraph({ children: [new PageBreak()] }),

          // 8. CRITERIOS DE ACEPTACIÓN
          createSectionHeader("3.3 Criterios de Aceptación", "sec6", 6),
          ...(criteriosAceptacion.length > 0 
            ? criteriosAceptacion.map((ca, idx) => new Paragraph({
                spacing: { before: 200 },
                children: [
                  new TextRun({ text: `CA-${idx + 1}: `, bold: true, color: LIVERPOOL_PINK, size: 20 }),
                  new TextRun({ text: ca, size: 22 })
                ]
              }))
            : []
          ),
          new Paragraph({ children: [new PageBreak()] }),

          // CAPÍTULO IV
          createChapterHeader("CAPÍTULO IV: COMPLEMENTOS", "cap4", 400),

          // 9. ANEXOS - HIPERVINCULOS
          createSectionHeader("4.1 Hipervínculos y Enlaces", "sec7", 7),
          ...(hipervinculos.length > 0 
            ? [
                ...hipervinculos.map(h => new Paragraph({
                  spacing: { before: 100 },
                  children: [
                    new TextRun({ text: `${h.etiqueta}: `, bold: true, size: 18 }),
                    new ExternalHyperlink({
                      children: [
                        new TextRun({
                          text: h.url,
                          style: "Hyperlink",
                          color: "0000FF",
                          underline: {},
                          size: 18
                        }),
                      ],
                      link: h.url,
                    }),
                  ]
                }))
              ]
            : [new Paragraph({ children: [new TextRun({ text: "No se proporcionaron hipervínculos.", size: 18, italics: true })] })]
          ),

          // 9. ANEXOS - VISUALES
          createSectionHeader("4.2 Anexos Visuales", "sec8", 8),
          ...(anexos.length > 0 
            ? anexos.flatMap((a, idx) => {
                const annexBuffer = a.contenido ? base64ToArrayBuffer(a.contenido) : null;
                
                return [
                  new Paragraph({
                    spacing: { before: 200 },
                    children: [
                      new TextRun({ text: `ANEXO ${idx + 1}: `, bold: true, color: LIVERPOOL_PINK, size: 20 }),
                      new TextRun({ text: a.nombre, size: 22, italics: true })
                    ]
                  }),
                  ...(annexBuffer ? [
                    new Paragraph({
                      alignment: AlignmentType.CENTER,
                      spacing: { before: 200, after: 200 },
                      children: [
                        new ImageRun({
                          data: annexBuffer as any,
                          type: "jpg",
                          transformation: { width: 500, height: 350 }
                        })
                      ]
                    })
                  ] : [
                    new Paragraph({
                      shading: { fill: SLATE_BG },
                      border: {
                        top: { style: BorderStyle.SINGLE, size: 4, color: "E2E8F0" },
                        bottom: { style: BorderStyle.SINGLE, size: 4, color: "E2E8F0" },
                        left: { style: BorderStyle.SINGLE, size: 4, color: "E2E8F0" },
                        right: { style: BorderStyle.SINGLE, size: 4, color: "E2E8F0" },
                      },
                      spacing: { before: 100, after: 200 },
                      children: [
                        new TextRun({ text: `[INSERTAR DOCUMENTO COMPLETO: ${a.nombre}]`, size: 18, color: "64748B" })
                      ]
                    })
                  ])
                ];
              })
            : []
          )
        ],
      },
    ],
  });

  try {
    const blob = await Packer.toBlob(doc);
    const fileName = `ERF_${(identificacion.sistema || "SYS").toUpperCase()}_${(identificacion.proceso || "PROCESO").replace(/\s+/g, "_")}.docx`;
    saveAs(blob, fileName);
  } catch (error) {
    console.error("Error generating Word file", error);
    alert("Hubo un error al generar el archivo Word. Por favor intenta de nuevo.");
  }
};
