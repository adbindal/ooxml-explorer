export interface ReferenceDoc {
  tag: string;
  namespace: string;
  domain: 'docx' | 'xlsx' | 'pptx' | 'shared';
  definition: string;
  attributes: string[];
  parents: string[];
}

export const KNOWLEDGE_BASE: ReferenceDoc[] = [
  // --- DOCX (WordprocessingML) ---
  {
    tag: 'document',
    namespace: 'w',
    domain: 'docx',
    definition: 'Main Document Part. The root element of the primary word processing document.',
    attributes: ['xmlns:w'],
    parents: []
  },
  {
    tag: 'body',
    namespace: 'w',
    domain: 'docx',
    definition: 'Document Body. Contains all block-level contents of the document (paragraphs, tables, etc.).',
    attributes: [],
    parents: ['document']
  },
  {
    tag: 'p',
    namespace: 'w',
    domain: 'docx',
    definition: 'Paragraph. The basic block-level container for text and inline elements in Word.',
    attributes: ['rsidR', 'rsidRDefault'],
    parents: ['body', 'tc']
  },
  {
    tag: 'r',
    namespace: 'w',
    domain: 'docx',
    definition: 'Run. An inline container that groups a series of text characters with a common set of formatting properties.',
    attributes: ['rsidR'],
    parents: ['p']
  },
  {
    tag: 't',
    namespace: 'w',
    domain: 'docx',
    definition: 'Text. Contains the actual text characters within a run.',
    attributes: ['xml:space'],
    parents: ['r']
  },
  {
    tag: 'tbl',
    namespace: 'w',
    domain: 'docx',
    definition: 'Table. Defines a tabular layout in the document.',
    attributes: [],
    parents: ['body', 'tc']
  },
  {
    tag: 'tr',
    namespace: 'w',
    domain: 'docx',
    definition: 'Table Row. Defines a row within a table.',
    attributes: ['rsidR'],
    parents: ['tbl']
  },
  {
    tag: 'tc',
    namespace: 'w',
    domain: 'docx',
    definition: 'Table Cell. Defines a single cell within a table row.',
    attributes: [],
    parents: ['tr']
  },
  {
    tag: 'cantSplit',
    namespace: 'w',
    domain: 'docx',
    definition: 'Table Row Cannot Split. Specifies that the contents of this row must not be split across multiple pages. If a page break occurs, the entire row is moved to the next page.',
    attributes: [],
    parents: ['trPr']
  },
  {
    tag: 'tblHeader',
    namespace: 'w',
    domain: 'docx',
    definition: 'Table Header Row. Specifies that this row is repeated at the top of each page if the table spans multiple pages.',
    attributes: [],
    parents: ['trPr']
  },

  // --- XLSX (SpreadsheetML) ---
  {
    tag: 'worksheet',
    namespace: 'r',
    domain: 'xlsx',
    definition: 'Worksheet. The root element for a single spreadsheet tab containing grid data.',
    attributes: ['xmlns'],
    parents: []
  },
  {
    tag: 'sheetData',
    namespace: 'r',
    domain: 'xlsx',
    definition: 'Sheet Data. The grid container for all rows and cells in the worksheet.',
    attributes: [],
    parents: ['worksheet']
  },
  {
    tag: 'row',
    namespace: 'r',
    domain: 'xlsx',
    definition: 'Spreadsheet Row. Defines a single row of cells in the grid.',
    attributes: ['r', 'spans', 'ht', 'customHeight'],
    parents: ['sheetData']
  },
  {
    tag: 'c',
    namespace: 'r',
    domain: 'xlsx',
    definition: 'Cell. Defines a single grid cell. Can contain values, formulas, or references to the shared strings table.',
    attributes: ['r', 't', 's'],
    parents: ['row']
  },
  {
    tag: 'v',
    namespace: 'r',
    domain: 'xlsx',
    definition: 'Cell Value. Contains the raw value of the cell (number, boolean, or index of a shared string).',
    attributes: [],
    parents: ['c']
  },
  {
    tag: 'f',
    namespace: 'r',
    domain: 'xlsx',
    definition: 'Cell Formula. Specifies the spreadsheet formula used to calculate the cell value.',
    attributes: ['t', 'ref'],
    parents: ['c']
  },
  {
    tag: 'sst',
    namespace: 'r',
    domain: 'xlsx',
    definition: 'Shared String Table. The root element containing all unique text strings used across the spreadsheet, optimizing file size.',
    attributes: ['count', 'uniqueCount'],
    parents: []
  },
  {
    tag: 'si',
    namespace: 'r',
    domain: 'xlsx',
    definition: 'Shared String Item. Represents a single text string within the Shared String Table.',
    attributes: [],
    parents: ['sst']
  },

  // --- PPTX (PresentationML) ---
  {
    tag: 'presentation',
    namespace: 'p',
    domain: 'pptx',
    definition: 'Presentation. The root element of a PowerPoint presentation, specifying slide list, sizes, and global settings.',
    attributes: ['xmlns:p'],
    parents: []
  },
  {
    tag: 'sld',
    namespace: 'p',
    domain: 'pptx',
    definition: 'Slide. The root element representing a single slide within the presentation.',
    attributes: [],
    parents: []
  },
  {
    tag: 'sldLayout',
    namespace: 'p',
    domain: 'pptx',
    definition: 'Slide Layout. The root element specifying the layout and placeholders for a slide.',
    attributes: ['type'],
    parents: []
  },
  {
    tag: 'sldMaster',
    namespace: 'p',
    domain: 'pptx',
    definition: 'Slide Master. The root element defining the master slide, formatting, background, and default theme for slides.',
    attributes: [],
    parents: []
  },
  {
    tag: 'sp',
    namespace: 'p',
    domain: 'pptx',
    definition: 'Shape. Represents a drawing shape (rectangle, circle, text box, etc.) on a slide.',
    attributes: [],
    parents: ['spTree', 'grpSp']
  },
  {
    tag: 'txBody',
    namespace: 'p',
    domain: 'pptx',
    definition: 'Text Body. Contains the text paragraphs and list properties within a slide shape.',
    attributes: [],
    parents: ['sp']
  },

  // --- SHARED (Packaging & Relationships) ---
  {
    tag: 'Relationships',
    namespace: 'r',
    domain: 'shared',
    definition: 'Relationships Part. The root element of a .rels file, containing all relationship mappings for a part.',
    attributes: ['xmlns'],
    parents: []
  },
  {
    tag: 'Relationship',
    namespace: 'r',
    domain: 'shared',
    definition: 'Relationship Definition. Specifies a link between a source part and a target resource (e.g., slide to slide layout, document to image).',
    attributes: ['Id', 'Type', 'Target', 'TargetMode'],
    parents: ['Relationships']
  },
  {
    tag: 'Types',
    namespace: 'r',
    domain: 'shared',
    definition: 'Content Types. The root element of [Content_Types].xml, registering the mime-type of every part in the package.',
    attributes: ['xmlns'],
    parents: []
  },
  {
    tag: 'Override',
    namespace: 'r',
    domain: 'shared',
    definition: 'Content Type Override. Specifies the content type of a specific part inside the ZIP package, overriding defaults.',
    attributes: ['PartName', 'ContentType'],
    parents: ['Types']
  },
  {
    tag: 'Default',
    namespace: 'r',
    domain: 'shared',
    definition: 'Default Content Type. Specifies the default content type for all parts with a specific file extension (e.g. .xml, .png).',
    attributes: ['Extension', 'ContentType'],
    parents: ['Types']
  }
];
