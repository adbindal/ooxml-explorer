/**
 * The analyzer registry — what this engine can check, as data rather than as control flow.
 *
 * Before this, knowing what the engine could do meant reading `ANALYSIS_TARGETS` in a
 * React component and following six imports. Two consequences, both of which showed up
 * as real gaps:
 *
 *   - The **validate** surface ran exactly one analyzer (`checkPackageIntegrity`), so
 *     "is this file correct?" checked content types and relationships and nothing else.
 *   - The **compare** surface ran the structural diff and no analyzer at all, so a file
 *     that had lost an OLE embedding or a bookmark's end marker reported the part change
 *     without reporting that anything was broken.
 *
 * Both are the same bug: there was no way to say "run everything that applies here."
 * A registry makes that one line, and makes adding the next analyzer a data change.
 *
 * THE HONEST HALF.
 *
 * Every entry declares `cannotDetermine` alongside `determines`. This is not
 * documentation — it is the input to the capability ledger, and it extends the honesty
 * property from per-fact to per-capability. The badge already refuses to call a claim
 * Verified without provenance; the same discipline says the engine should be able to
 * answer "why can't you check this?" from data rather than from a model's guess about
 * its own abilities. An analyzer that lists nothing it cannot do is almost certainly
 * lying.
 */

import type { Finding } from './findings';
import type { ComputedEvidence } from './aiService';
import type { PackageParts } from './packageIntegrity';
import { checkPackageIntegrity } from './packageIntegrity';
import { readBookmarks } from './wordBookmarks';
import { readComments, COMMENT_PART_PATHS } from './wordComments';
import { readFields, crossCheckFieldTargets, computeFieldEvidenceForMarkup } from './wordFields';
import { readRevisions, checkRevisionVisibility, readRevisionSettings, computeRevisionEvidenceForMarkup } from './wordRevisions';
import { tableGridFindings, computeTableEvidenceForMarkup } from './wordTableGrid';
import { mediaFindings, computeMediaEvidenceForMarkup, MEDIA_HOST_PART } from './pptMedia';
import { formulaFindings, computeFormulaEvidenceForMarkup, FORMULA_HOST_PART } from './excelFormulas';
import { contentControlFindings, computeContentControlEvidenceForMarkup, SDT_HOST_PART } from './wordContentControls';
import { styleReferenceFindings, computeStyleReferenceEvidenceForMarkup, hasStyleReferences } from './styleReferences';
import { noteFindings, computeNoteEvidenceForMarkup, hasNotes } from './wordNotes';
import { animationFindings, computeAnimationEvidenceForMarkup, ANIMATION_HOST_PART } from './pptAnimation';
import { externalLinkFindings, computeExternalLinkEvidenceForMarkup, EXTERNAL_LINK_HOST_PART } from './excelExternalLinks';
import { hyperlinkFindings, computeHyperlinkEvidenceForMarkup, HYPERLINK_HOST_PART } from './hyperlinks';
import { readOleObjects } from './oleObjects';
import { readPivotTables, computePivotEvidenceForMarkup } from './excelPivotTables';
import { normaliseParts, detectConformance, conformanceFindings, toTransitionalXml } from './conformance';
import { computeBookmarkEvidenceForMarkup } from './wordBookmarks';
import { computeCommentEvidenceForMarkup } from './wordComments';
import { computeOleEvidenceForMarkup } from './oleObjects';
import { computeEvidenceForMarkup } from './wordFormattingAnalysis';
import { computeExcelEvidenceForMarkup } from './excelFormattingAnalysis';
import { computePowerpointEvidenceForMarkup } from './powerpointFormattingAnalysis';
import { computeChartEvidenceForMarkup, chartFindings, CHART_HOST_PART } from './chartSemantics';

export type OoxmlFormat = 'docx' | 'xlsx' | 'pptx';

/**
 * How an analyzer participates in "explain this element".
 *
 * Separate from `analyze` because the two answer different questions. `analyze` finds
 * what is WRONG with a package; `explain` describes what a selected element IS and how
 * it resolves — the six-layer Word cascade, the text a bookmark covers, where a chart's
 * numbers come from. Most analyzers do one or the other; a few do both.
 */
export interface AnalyzerExplain {
  /** True when this analyzer has something to say about the open part. */
  matches: (partPath: string) => boolean;
  /** Named sibling parts to fetch alongside the open one. */
  siblings?: readonly string[];
  /** Or a pattern applied to the package's part list, when the set is not fixed. */
  siblingPattern?: RegExp;
  compute: (parts: PackageParts, rawXml: string) => ComputedEvidence | null;
}

export interface Analyzer {
  /** Matches the namespace of the codes it emits, e.g. `bookmark` → `bookmark/…`. */
  id: string;
  /** Short human-readable name, for the capability ledger. */
  title: string;
  formats: readonly OoxmlFormat[];
  /** What this analyzer establishes, phrased as the questions it answers. */
  determines: readonly string[];
  /**
   * What it explicitly does NOT establish, so the engine can say so rather than
   * leaving a reader to assume silence means "fine".
   */
  cannotDetermine: readonly string[];
  /**
   * True when the package contains anything for this analyzer to look at. Keeps the
   * ledger honest about what actually ran versus what merely could have.
   */
  appliesTo: (parts: PackageParts) => boolean;
  /** Absent when this analyzer only explains and never reports faults. */
  analyze?: (parts: PackageParts) => Finding[];
  /**
   * Set when this analyzer must see the package exactly as written.
   *
   * Every other analyzer receives markup with Strict namespaces already mapped to
   * Transitional (see conformance.ts) — that mapping is what lets them compare
   * namespaces by exact equality. The conformance analyzer is the one that reports
   * *whether that mapping happened*, so handing it normalised markup would leave it
   * permanently convinced every package is Transitional.
   */
  readsRawMarkup?: boolean;
  /** Absent when this analyzer only validates and has nothing to say about one element. */
  explain?: AnalyzerExplain;
}

/** Word story parts, which is where bookmarks and comment anchors live. */
const WORD_BODY = /^word\/(?:document\d*|header[^/]*|footer[^/]*|footnotes\d*|endnotes\d*)\.xml$/;
/** Parts that can host an OLE object, in any of the three formats. */
const OLE_HOST =
  /^(?:word\/(?:document\d*|header[^/]*|footer[^/]*|footnotes\d*|endnotes\d*)\.xml|xl\/worksheets\/[^/]+\.xml|ppt\/slides\/[^/]+\.xml)$/;

const parse = (xml: string | undefined): Document | null => {
  if (xml === undefined) return null;
  const doc = new DOMParser().parseFromString(xml, 'application/xml');
  return doc.getElementsByTagName('parsererror').length > 0 ? null : doc;
};

const matching = (parts: PackageParts, pattern: RegExp) => Object.keys(parts).filter(p => pattern.test(p));

export const ANALYZERS: readonly Analyzer[] = [
  {
    id: 'conformance',
    title: 'Conformance class',
    formats: ['docx', 'xlsx', 'pptx'],
    determines: [
      'whether the package is written in ISO Strict or in Transitional',
      'that Strict namespaces were mapped to Transitional so the other checks could read the file at all'
    ],
    cannotDetermine: [
      'the differences the two conformance classes genuinely disagree on — this maps namespaces, it does not convert',
      'whether a Strict-only construct was misread by a check built against Transitional markup'
    ],
    readsRawMarkup: true,
    appliesTo: parts => detectConformance(parts) === 'strict',
    analyze: parts => conformanceFindings(parts)
  },
  {
    id: 'package',
    title: 'Package integrity',
    formats: ['docx', 'xlsx', 'pptx'],
    determines: [
      'whether every part is declared in [Content_Types].xml',
      'whether every relationship resolves to a part that exists',
      'whether implicit relationships a consumer needs are present and unambiguous'
    ],
    cannotDetermine: [
      'whether a part that exists contains valid content for its type',
      'whether the ZIP container itself is well-formed — that is checked when the file is opened'
    ],
    // Always applicable: every OPC package has content types and relationships.
    appliesTo: () => true,
    analyze: parts => checkPackageIntegrity(parts)
  },
  {
    id: 'bookmark',
    title: 'Bookmarks and the shared markup id space',
    formats: ['docx'],
    determines: [
      'whether every bookmark range opens and closes',
      'whether any w:id is reused across bookmarks, tracked changes or permissions',
      'what text a bookmark covers'
    ],
    cannotDetermine: [
      'whether a bookmark is referenced from a different part — each part is read on its own',
      'whether a field instruction that names a bookmark would actually resolve at render time'
    ],
    appliesTo: parts => matching(parts, WORD_BODY).length > 0,
    analyze: parts =>
      matching(parts, WORD_BODY).flatMap(path => {
        const doc = parse(parts[path]);
        return doc ? readBookmarks(doc, path).problems : [];
      }),
    explain: {
      // Ids are unique per part, not per package, so nothing outside the open file is
      // needed to pair the ranges.
      matches: path => WORD_BODY.test(path) || /^word\/comments\d*\.xml$/.test(path),
      compute: computeBookmarkEvidenceForMarkup
    }
  },
  {
    id: 'comment',
    title: 'Comments, anchoring and threading',
    formats: ['docx'],
    determines: [
      'whether every comment anchor has a body and every body an anchor',
      'what text a comment is attached to',
      'whether a comment is a reply, and whether it is resolved — when the side-car is present'
    ],
    cannotDetermine: [
      'threading or resolved-state when word/commentsExtended.xml is absent; unknown is reported as unknown, never as "not a reply"',
      'whether Word would display a comment that has no commentReference — the schema makes all three markers optional'
    ],
    appliesTo: parts => parts[COMMENT_PART_PATHS.comments] !== undefined || matching(parts, WORD_BODY).length > 0,
    analyze: parts => {
      const body = matching(parts, WORD_BODY)[0];
      const document = parse(body ? parts[body] : undefined);
      if (!document) return [];
      return readComments({
        document,
        comments: parse(parts[COMMENT_PART_PATHS.comments]),
        commentsExtended: parse(parts[COMMENT_PART_PATHS.extended]),
        commentsIds: parse(parts['word/commentsIds.xml'])
      }).problems;
    },
    explain: {
      matches: path => WORD_BODY.test(path),
      siblings: ['word/comments.xml', 'word/commentsExtended.xml', 'word/commentsIds.xml'],
      compute: computeCommentEvidenceForMarkup
    }
  },
  {
    id: 'field',
    title: 'Fields and their cached results',
    formats: ['docx'],
    determines: [
      'whether a cross-reference points at a bookmark that still exists',
      'whether the text a field displays will be recalculated, or is being presented as current',
      'whether every field opens and closes, including nested ones'
    ],
    cannotDetermine: [
      'whether a page-dependent field (TOC, PAGE, NUMPAGES) shows the right number — that needs the document laid out',
      'references to a bookmark in a different part; each story is checked against its own bookmarks',
      'what a field would evaluate to if recalculated — nothing here executes a field'
    ],
    appliesTo: parts => matching(parts, WORD_BODY).length > 0,
    analyze: parts =>
      matching(parts, WORD_BODY).flatMap(path => {
        const doc = parse(parts[path]);
        if (!doc) return [];
        return [...readFields(doc, path).problems, ...crossCheckFieldTargets(doc, path)];
      }),
    explain: {
      matches: path => WORD_BODY.test(path),
      compute: computeFieldEvidenceForMarkup
    }
  },
  {
    id: 'revision',
    title: 'Tracked changes',
    formats: ['docx'],
    determines: [
      'what the document says with every change accepted, versus rejected — two different texts',
      'whether both halves of a tracked move still pair up by name',
      'whether a revision id collides with a bookmark or permission id'
    ],
    cannotDetermine: [
      'what Word displays for markup that is out of schema, such as live text inside a deletion — the fault is reported, the rendering is not predicted',
      'whether two adjacent paragraph-mark revisions by different authors interact; each mark is resolved independently'
    ],
    appliesTo: parts => matching(parts, WORD_BODY).length > 0,
    analyze: parts => {
      const settings = parse(parts['word/settings.xml']);
      return matching(parts, WORD_BODY).flatMap(path => {
        const doc = parse(parts[path]);
        if (!doc) return [];
        const { revisions, problems } = readRevisions(doc, path);
        // Settings live in a different part, so the visibility check takes both and is
        // a separate call rather than folded into readRevisions.
        return [
          ...problems,
          ...(settings ? checkRevisionVisibility(readRevisionSettings(settings), revisions.length, path) : [])
        ];
      });
    },
    explain: {
      matches: path => WORD_BODY.test(path),
      siblings: ['word/settings.xml'],
      compute: computeRevisionEvidenceForMarkup
    }
  },
  {
    id: 'hyperlink',
    title: 'Hyperlinks and their destinations',
    formats: ['docx', 'xlsx', 'pptx'],
    determines: [
      'whether an internal link’s anchor names a bookmark that exists',
      'whether a link references a relationship the part actually declares',
      'whether a hyperlink relationship is declared that nothing points at'
    ],
    cannotDetermine: [
      'whether an external URL resolves — no network request is ever made, so an external target is reported as unverifiable rather than as working or broken',
      'which destination wins when a Word hyperlink carries both r:id and w:anchor; that is reported as ambiguous and resolved in neither direction',
      'the full ppaction:// verb vocabulary, which the schema does not enumerate — unknown verbs are reported verbatim and never judged'
    ],
    appliesTo: parts => Object.keys(parts).some(p => HYPERLINK_HOST_PART.test(p)),
    analyze: parts =>
      Object.keys(parts)
        .filter(p => HYPERLINK_HOST_PART.test(p))
        .flatMap(path => hyperlinkFindings(parts, path)),
    explain: {
      matches: path => HYPERLINK_HOST_PART.test(path),
      siblingPattern: /^(?:word|xl|ppt)\/(?:.*\/)?_rels\/[^/]+\.rels$|^xl\/workbook\.xml$|^ppt\/slides\/[^/]+\.xml$/,
      compute: computeHyperlinkEvidenceForMarkup
    }
  },
  {
    id: 'contentControl',
    title: 'Content controls and data bindings',
    formats: ['docx'],
    determines: [
      'whether a control’s data binding names a custom XML part that exists in the package',
      'whether a control is displaying its placeholder rather than data',
      'whether two controls share an id, so that driving the document by id reaches only one'
    ],
    cannotDetermine: [
      'whether a binding’s XPath selects anything inside the part it names — that needs a namespace-aware XPath engine and the binding’s prefixMappings',
      'whether a control’s stored text is current with respect to the data it is bound to; the text is stored, not computed'
    ],
    appliesTo: parts => Object.keys(parts).some(p => SDT_HOST_PART.test(p)),
    analyze: parts => contentControlFindings(parts),
    explain: {
      matches: path => SDT_HOST_PART.test(path),
      siblingPattern: /^customXml\/(?:item|itemProps)\d*\.xml$/,
      compute: computeContentControlEvidenceForMarkup
    }
  },
  {
    id: 'table',
    title: 'Table grid and merge geometry',
    formats: ['docx'],
    determines: [
      'whether each row’s gridSpan values sum to the declared column grid',
      'whether every vMerge continuation has a restart above it in the same column'
    ],
    cannotDetermine: [
      'how Word will actually repair a table whose geometry does not add up — it does repair them, and the result is not specified',
      'rendered column widths, which depend on layout rules this does not evaluate'
    ],
    appliesTo: parts => matching(parts, WORD_BODY).length > 0,
    analyze: parts =>
      matching(parts, WORD_BODY).flatMap(path => {
        const doc = parse(parts[path]);
        return doc ? tableGridFindings(doc, path) : [];
      }),
    explain: {
      matches: path => WORD_BODY.test(path),
      compute: computeTableEvidenceForMarkup
    }
  },
  {
    id: 'media',
    title: 'Audio and video',
    formats: ['pptx'],
    determines: [
      'whether a slide’s media is embedded in the package or linked to a file outside it',
      'whether the media data is present behind the poster frame that renders in its place'
    ],
    cannotDetermine: [
      'whether a linked media file resolves — its target is outside the package by definition',
      'whether the media stream is playable; the bytes are located but never decoded'
    ],
    appliesTo: parts => Object.keys(parts).some(p => MEDIA_HOST_PART.test(p)),
    analyze: parts =>
      Object.keys(parts)
        .filter(p => MEDIA_HOST_PART.test(p))
        .flatMap(path => mediaFindings(parts, path)),
    explain: {
      matches: path => MEDIA_HOST_PART.test(path),
      siblingPattern: /^ppt\/(?:_rels\/[^/]+\.rels|slides\/_rels\/[^/]+\.rels|media\/[^/]+)$/,
      compute: computeMediaEvidenceForMarkup
    }
  },
  {
    id: 'ole',
    title: 'Embedded objects and their preview images',
    formats: ['docx', 'xlsx', 'pptx'],
    determines: [
      'whether an embedded object still has its data, or only the preview that renders in front of it',
      'whether the declared binding agrees with how the relationship is actually targeted'
    ],
    cannotDetermine: [
      'whether the embedded binary is valid for the progId it claims — the compound file is never opened',
      'whether a linked object resolves, since its target is outside the package by definition'
    ],
    appliesTo: parts => matching(parts, OLE_HOST).length > 0,
    analyze: parts => matching(parts, OLE_HOST).flatMap(path => readOleObjects(parts, path).flatMap(o => o.problems)),
    explain: {
      matches: path => OLE_HOST.test(path),
      // Needs the .rels to resolve the object, and the target part itself to see
      // whether it is actually there - which is the whole point.
      siblingPattern: /^(?:word|xl|ppt)\/(?:.*\/)?(?:_rels\/[^/]+\.rels|embeddings\/[^/]+|media\/[^/]+)$/,
      compute: computeOleEvidenceForMarkup
    }
  },
  {
    id: 'formula',
    title: 'Formulas and their cached values',
    formats: ['xlsx'],
    determines: [
      'whether a shared-formula follower still has the master that holds its formula text',
      'what the workbook says about its own numbers — full recalculation on load, manual mode, an unfinished calculation',
      'which cells store an error, reach outside the package, or depend on when they were calculated'
    ],
    cannotDetermine: [
      'whether a cached value equals what its formula would now produce — that needs a calculation engine, which this is not',
      'the contents of a referenced external workbook, which lives outside this package',
      'which of the 218 documented [MS-OI29500] variations for this clause a given workbook relies on'
    ],
    appliesTo: parts => Object.keys(parts).some(p => FORMULA_HOST_PART.test(p)),
    analyze: parts => formulaFindings(parts),
    explain: {
      matches: path => FORMULA_HOST_PART.test(path),
      siblings: ['xl/workbook.xml'],
      compute: computeFormulaEvidenceForMarkup
    }
  },
  {
    id: 'pivot',
    title: 'Pivot tables and the cache chain',
    formats: ['xlsx'],
    determines: [
      'which hop of the pivot cache chain is broken, when one is',
      'whether a pivot field index falls outside the cache field count'
    ],
    cannotDetermine: [
      'whether the cached records still agree with their source range — staleness needs the source recomputed',
      'which of the 67 documented [MS-OI29500] variations for this clause a given workbook actually relies on — the count is verified, the per-file impact is not'
    ],
    appliesTo: parts => Object.keys(parts).some(p => p.startsWith('xl/pivotTables/') || p.startsWith('xl/pivotCache/')),
    analyze: parts => readPivotTables(parts).flatMap(t => [...t.chain.problems, ...t.problems]),
    explain: {
      matches: path => /^xl\/worksheets\/[^/]+\.xml$/.test(path),
      siblingPattern:
        /^xl\/(?:workbook\.xml|_rels\/workbook\.xml\.rels|worksheets\/_rels\/[^/]+\.rels|pivotTables\/(?:_rels\/)?[^/]+|pivotCache\/(?:_rels\/)?[^/]+)$/,
      compute: computePivotEvidenceForMarkup
    }
  },
  {
    id: 'note',
    title: 'Footnotes and endnotes',
    formats: ['docx'],
    determines: [
      'whether every note reference has a note behind it, and every note is referenced',
      'whether the separator notes that draw the rule above the notes are present'
    ],
    cannotDetermine: [
      'references from a body part not supplied with the request, which would make a note look orphaned when it is not'
    ],
    appliesTo: hasNotes,
    analyze: parts => noteFindings(parts),
    explain: {
      matches: path => /^word\/(?:document\d*|header[^/]*|footer[^/]*|footnotes\d*|endnotes\d*)\.xml$/.test(path),
      siblings: ['word/footnotes.xml', 'word/endnotes.xml'],
      compute: computeNoteEvidenceForMarkup
    }
  },
  {
    id: 'animation',
    title: 'Slide animations and timing',
    formats: ['pptx'],
    determines: [
      'whether an animation targets a shape that still exists — a dead animation never fires and never warns',
      'how many of a slide’s animations will never run'
    ],
    cannotDetermine: [
      'whether an animation that resolves is the RIGHT animation; only that its target exists',
      'targets using the Office 2007 string form of @spid, which is left unjudged',
      'p:subSp and p:charRg targets, and paragraph ranges on frames that are not p:sp'
    ],
    appliesTo: parts => Object.keys(parts).some(p => ANIMATION_HOST_PART.test(p)),
    analyze: parts =>
      Object.keys(parts)
        .filter(p => ANIMATION_HOST_PART.test(p))
        .flatMap(path => {
          const doc = parse(parts[path]);
          return doc ? animationFindings(doc, path) : [];
        }),
    explain: {
      matches: path => ANIMATION_HOST_PART.test(path),
      compute: computeAnimationEvidenceForMarkup
    }
  },
  {
    id: 'externalLink',
    title: 'Links to other workbooks',
    formats: ['xlsx'],
    determines: [
      'whether each external reference resolves through the workbook relationships to a link part that exists',
      'whether a formula’s [N] index names a reference that is declared'
    ],
    cannotDetermine: [
      'whether the linked workbook itself exists or is current — it is outside this package and is never fetched',
      'whether the cached values still match the source, which is the whole reason they are cached'
    ],
    appliesTo: parts => Object.keys(parts).some(p => EXTERNAL_LINK_HOST_PART.test(p)),
    analyze: parts => externalLinkFindings(parts),
    explain: {
      matches: path => EXTERNAL_LINK_HOST_PART.test(path) || /^xl\/worksheets\/[^/]+\.xml$/.test(path),
      siblingPattern: /^xl\/(?:workbook\.xml|_rels\/workbook\.xml\.rels|external(?:Links|References)\/(?:_rels\/)?[^/]+)$/,
      compute: computeExternalLinkEvidenceForMarkup
    }
  },
  {
    id: 'styleRef',
    title: 'Style and format references',
    formats: ['docx', 'xlsx'],
    determines: [
      'whether every pStyle, rStyle and tblStyle names a style that is defined',
      'whether every numId names a numbering definition that exists',
      'whether a cell’s format index, and its font, fill, border and custom number format ids, are all in range'
    ],
    cannotDetermine: [
      'whether a style that IS defined produces the intended appearance — only that the reference lands somewhere',
      'references against a stylesheet not supplied with the request; a style defined in an unloaded part would be reported as missing'
    ],
    appliesTo: hasStyleReferences,
    analyze: parts => styleReferenceFindings(parts),
    explain: {
      matches: path =>
        /^word\/(?:document\d*|header[^/]*|footer[^/]*|footnotes\d*|endnotes\d*)\.xml$/.test(path) ||
        /^xl\/worksheets\/[^/]+\.xml$/.test(path),
      siblings: ['word/styles.xml', 'word/numbering.xml', 'xl/styles.xml'],
      compute: computeStyleReferenceEvidenceForMarkup
    }
  },
  {
    id: 'word-formatting',
    title: 'Word formatting cascade',
    formats: ['docx'],
    determines: [
      'the value every formatting property resolves to, and which of the six cascade layers set it',
      'where Word diverges from ECMA-376 — toggle properties, numbering precedence, table style banding'
    ],
    cannotDetermine: [
      'how the resolved formatting actually renders — line breaking, font substitution and pagination are the renderer’s',
      'formatting in word/glossary/, which resolves against its own stylesheet'
    ],
    appliesTo: parts => parts['word/styles.xml'] !== undefined,
    explain: {
      // Every Word story, not just document.xml: a paragraph in a header or footnote
      // resolves against the same styles.xml and numbering.xml.
      matches: path =>
        /^word\/(?:document\d*|header[^/]*|footer[^/]*|footnotes\d*|endnotes\d*|comments\d*)\.xml$/.test(path),
      siblings: ['word/styles.xml', 'word/numbering.xml'],
      compute: computeEvidenceForMarkup
    }
  },
  {
    id: 'excel-formatting',
    title: 'Excel cell formats and values',
    formats: ['xlsx'],
    determines: [
      'the format a cell resolves to through cellXfs, which is a lookup and not a cascade',
      'which date epoch the workbook uses, and what a serial number therefore means'
    ],
    cannotDetermine: [
      'the result of a formula — nothing here evaluates one, only the cached value is read',
      'whether a cached value is still current with respect to its formula'
    ],
    appliesTo: parts => parts['xl/styles.xml'] !== undefined,
    explain: {
      matches: path => /^xl\/worksheets\/[^/]+\.xml$/.test(path),
      siblings: ['xl/styles.xml', 'xl/workbook.xml', 'xl/sharedStrings.xml'],
      compute: computeExcelEvidenceForMarkup
    }
  },
  {
    id: 'powerpoint',
    title: 'Slide placeholder and theme inheritance',
    formats: ['pptx'],
    determines: [
      'what a shape inherits from its layout and master, matched on @idx or @type',
      'which theme colour a colour-map reference resolves to'
    ],
    cannotDetermine: [
      'layout-to-master placeholder matching, which the specification does not define — Office’s behaviour here is undocumented',
      'the rendered position of a shape inside a group beyond the transform arithmetic'
    ],
    appliesTo: parts => Object.keys(parts).some(p => p.startsWith('ppt/slideLayouts/')),
    explain: {
      // Every hop is an implicit relationship, so the .rels parts are as load-bearing
      // as the content parts and have to be fetched alongside them.
      matches: path => /^ppt\/slides\/[^/]+\.xml$/.test(path),
      siblingPattern: /^ppt\/(slides|slideLayouts|slideMasters|theme)\/(_rels\/)?[^/]+$/,
      compute: computePowerpointEvidenceForMarkup
    }
  },
  {
    id: 'chart',
    title: 'Chart structure for translation',
    formats: ['docx', 'xlsx', 'pptx'],
    determines: [
      'where a series’ numbers come from, and whether the cache is the only source',
      'which parts of a chart carry meaning and which are presentation only'
    ],
    cannotDetermine: [
      'whether a cached series still agrees with the workbook it was read from — the values are compared to nothing',
      'how the chart is laid out — axis scaling and label placement are the renderer’s'
    ],
    appliesTo: parts => Object.keys(parts).some(p => CHART_HOST_PART.test(p)),
    analyze: parts =>
      Object.keys(parts)
        .filter(p => CHART_HOST_PART.test(p))
        .flatMap(path => chartFindings(parts, path)),
    explain: {
      // Charts are self-contained: the series cache travels with the part.
      matches: path => /charts\/chart[^/]*\.xml$/.test(path),
      compute: computeChartEvidenceForMarkup
    }
  }
] as const;

export interface AnalysisRun {
  findings: Finding[];
  /** Analyzer ids that found something to look at and ran. */
  ran: string[];
  /** Analyzer ids with nothing in this package to examine. */
  skipped: string[];
}

/**
 * Runs every analyzer that applies to this package.
 *
 * One analyzer throwing must not lose the others' findings — a malformed part in a
 * corner of the package is exactly when the rest of the report matters most.
 */
export function analyzePackage(rawParts: PackageParts, analyzers: readonly Analyzer[] = ANALYZERS): AnalysisRun {
  // Strict packages spell every namespace differently. Mapping them once here is what
  // lets the analyzers keep comparing namespaces by exact equality; see conformance.ts.
  const parts = normaliseParts(rawParts);
  const findings: Finding[] = [];
  const ran: string[] = [];
  const skipped: string[] = [];

  for (const analyzer of analyzers) {
    // The conformance analyzer is the one thing that has to see the file as written.
    const view = analyzer.readsRawMarkup ? rawParts : parts;
    if (!analyzer.appliesTo(view)) {
      skipped.push(analyzer.id);
      continue;
    }
    if (!analyzer.analyze) {
      // Explain-only: it applies to the package but contributes no faults, so counting
      // it as "ran" would overstate what the validate pass actually checked.
      skipped.push(analyzer.id);
      continue;
    }
    ran.push(analyzer.id);
    try {
      findings.push(...analyzer.analyze(view));
    } catch {
      // Deliberately swallowed: a thrown analyzer is a bug in this engine, not a
      // finding about the user's file, and reporting it as one would be a lie.
    }
  }

  return { findings, ran, skipped };
}

export interface CapabilityLedger {
  ran: Array<{ id: string; title: string; determines: readonly string[] }>;
  skipped: Array<{ id: string; title: string }>;
  /** Every limit declared by an analyzer that actually ran. */
  limits: string[];
}

/**
 * What the engine checked, what it skipped, and what it cannot tell you.
 *
 * Computed entirely from the registry — the model never asserts its own capabilities,
 * exactly as it never asserts the evidence tier. `limits` is the part worth showing a
 * user: a clean report means "nothing found by the checks that ran", and the difference
 * between that and "this file is fine" is the whole point.
 */
export function capabilityLedger(run: AnalysisRun, analyzers: readonly Analyzer[] = ANALYZERS): CapabilityLedger {
  const byId = new Map(analyzers.map(a => [a.id, a]));
  return {
    ran: run.ran.map(id => {
      const a = byId.get(id)!;
      return { id, title: a.title, determines: a.determines };
    }),
    skipped: run.skipped.map(id => ({ id, title: byId.get(id)!.title })),
    limits: run.ran.flatMap(id => byId.get(id)!.cannotDetermine)
  };
}

/**
 * Identity of a finding for the purpose of comparing two packages.
 *
 * Deliberately excludes the message: messages interpolate counts and paths, so two
 * genuinely identical faults can word themselves differently. Code, part and subject
 * are what make a finding the *same* finding.
 */
const identityOf = (f: Finding): string =>
  `${f.code}|${f.part}|${JSON.stringify(f.subject ?? {})}`;

export interface FindingsDelta {
  /** Present after, absent before — what this change broke. */
  introduced: Finding[];
  /** Present before, absent after — what it fixed. */
  resolved: Finding[];
  /** Present in both, and therefore not this change's doing. */
  unchanged: Finding[];
}

/**
 * What a change did to the health of a package.
 *
 * This is the question a regression investigation actually asks, and no structural diff
 * can answer it: a dropped OLE embedding shows up as one removed part, which looks like
 * any other removed part. Here it shows up as an introduced finding that says the
 * document still renders correctly and is broken anyway.
 */
export function diffFindings(before: PackageParts, after: PackageParts): FindingsDelta {
  const first = new Map(analyzePackage(before).findings.map(f => [identityOf(f), f]));
  const second = new Map(analyzePackage(after).findings.map(f => [identityOf(f), f]));

  const introduced: Finding[] = [];
  const resolved: Finding[] = [];
  const unchanged: Finding[] = [];

  for (const [id, f] of second) (first.has(id) ? unchanged : introduced).push(f);
  for (const [id, f] of first) if (!second.has(id)) resolved.push(f);

  return { introduced, resolved, unchanged };
}

/**
 * Which analyzers have something to say about one open part.
 *
 * Returned in registry order so the evidence bundle is stable between runs — a bundle
 * that reshuffles makes two otherwise-identical answers look different.
 */
export const explainersFor = (partPath: string, analyzers: readonly Analyzer[] = ANALYZERS): Analyzer[] =>
  analyzers.filter(a => a.explain?.matches(partPath));

/**
 * Every sibling part the matching analyzers want, resolved against what the package
 * actually has.
 *
 * Unioned so a part wanted by three analyzers is fetched once. `available` is the list
 * of part paths the caller can supply; a named sibling that is not there is dropped
 * rather than requested, because the analyses report an absent part themselves.
 */
export const siblingsFor = (
  partPath: string,
  available: readonly string[],
  analyzers: readonly Analyzer[] = ANALYZERS
): string[] => {
  const wanted = new Set<string>();
  for (const analyzer of explainersFor(partPath, analyzers)) {
    const explain = analyzer.explain!;
    if (explain.siblingPattern) {
      for (const path of available) {
        if (path !== partPath && explain.siblingPattern.test(path)) wanted.add(path);
      }
    }
    for (const path of explain.siblings ?? []) {
      if (available.includes(path)) wanted.add(path);
    }
  }
  return [...wanted];
};

/**
 * Runs every matching explainer and merges what they produce.
 *
 * One analyzer throwing must not suppress the rest, and one returning nothing is not a
 * failure — most parts are only interesting to two or three of them.
 *
 * Returns null when nothing matched at all, which is the signal the caller uses to
 * record a coverage gap: the honest answer there is that no check applies, not that
 * the markup is fine.
 */
export function explainPart(
  rawParts: PackageParts,
  partPath: string,
  rawXml: string,
  registry: readonly Analyzer[] = ANALYZERS
): { evidence: ComputedEvidence; contributors: string[] } | null {
  const analyzers = explainersFor(partPath, registry);
  if (analyzers.length === 0) return null;

  // The selected element's markup is normalised too - it came from the same file, and
  // a Strict-namespaced rawXml would not match anything the analyzers look for.
  const parts = normaliseParts(rawParts);
  rawXml = toTransitionalXml(rawXml);

  const lines: string[] = [];
  const unresolved: string[] = [];
  const contributors: string[] = [];

  for (const analyzer of analyzers) {
    let result: ComputedEvidence | null = null;
    try {
      result = analyzer.explain!.compute(parts, rawXml);
    } catch {
      continue;
    }
    if (!result) continue;
    contributors.push(analyzer.id);
    lines.push(...result.lines);
    unresolved.push(...result.unresolved);
  }

  if (lines.length === 0 && unresolved.length === 0) return null;

  // De-duplicated because analyzers legitimately overlap - bookmarks and the Word
  // cascade both read document.xml and can reach the same conclusion about it.
  return {
    evidence: { lines: [...new Set(lines)], unresolved: [...new Set(unresolved)] },
    contributors
  };
}
