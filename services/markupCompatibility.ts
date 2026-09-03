/**
 * Markup Compatibility and Extensibility (ECMA-376 Part 3).
 *
 * This is a required preprocessing pass, not an optional nicety. Every modern Office
 * file is simultaneously several documents: when Word 2010+ writes a shape it emits it
 * **twice** - once as DrawingML inside `mc:Choice`, once as legacy VML inside
 * `mc:Fallback` - so that older consumers still render something. A reader that does
 * not resolve this either:
 *
 *   - **double-counts**, walking both branches and reporting every text box, image and
 *     shape twice; or
 *   - **zero-counts**, looking for `w:drawing`/`w:pict` at the expected depth, finding
 *     an `mc:AlternateContent` wrapper instead, and reporting nothing.
 *
 * Both are silent. Neither throws.
 *
 * Which branch wins depends on what the *consumer* understands, which is why the
 * understood namespace set is a parameter rather than a constant. Resolving the same
 * file as a 2007-era consumer and as a modern one is a legitimate and useful thing to
 * do - it is precisely how you answer "why does this look different in an older Word?"
 *
 * Namespace URIs below were taken from the Open XML SDK's published `namespaces.json`
 * rather than from memory. Note that the year in an extension URI is *not* the product
 * year: `w15` is `.../office/word/2012/wordml` but ships in Office **2013**.
 */

export const MCE_NAMESPACE = 'http://schemas.openxmlformats.org/markup-compatibility/2006';

/** The original ECMA-376 namespaces - roughly what a 2007-era consumer understood. */
export const LEGACY_CONSUMER_NAMESPACES: readonly string[] = [
  'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
  'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
  'http://schemas.openxmlformats.org/presentationml/2006/main',
  'http://schemas.openxmlformats.org/drawingml/2006/main',
  'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
  'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing',
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
  'urn:schemas-microsoft-com:vml',
  'urn:schemas-microsoft-com:office:word'
];

/** Legacy plus the Office 2010+ extension namespaces a current build understands. */
export const MODERN_CONSUMER_NAMESPACES: readonly string[] = [
  ...LEGACY_CONSUMER_NAMESPACES,
  'http://schemas.microsoft.com/office/word/2010/wordml',
  'http://schemas.microsoft.com/office/word/2012/wordml',
  'http://schemas.microsoft.com/office/word/2010/wordprocessingShape',
  'http://schemas.microsoft.com/office/word/2010/wordprocessingGroup',
  'http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas',
  'http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing'
];

export interface MceSelection {
  /** Element path to the resolved `mc:AlternateContent`, for reporting. */
  path: string;
  /** Which branch was kept. `nothing` means no Choice matched and no Fallback existed. */
  chose: 'choice' | 'fallback' | 'nothing';
  /** The `Requires` value of the winning Choice, when one won. */
  requires?: string;
  /** `Requires` values of the Choices that were passed over, in document order. */
  rejected: string[];
}

export interface MceResult {
  /** The same Document instance, mutated in place. */
  document: Document;
  /** One entry per `mc:AlternateContent` resolved, in the order they were processed. */
  selections: MceSelection[];
}

/** Builds a readable ancestor path such as `w:document/w:body/w:p/w:r`. */
const pathOf = (element: Element): string => {
  const segments: string[] = [];
  let current: Element | null = element.parentElement;
  while (current) {
    segments.unshift(current.nodeName);
    current = current.parentElement;
  }
  return segments.join('/');
};

/**
 * Resolves `Requires` against the namespaces in scope at that element.
 *
 * `Requires` holds namespace *prefixes*, not URIs, and those prefixes are resolved
 * using the declarations in scope where the Choice appears - so the same prefix can
 * legitimately mean different things in different parts of a document.
 */
const choiceIsUnderstood = (choice: Element, understood: Set<string>): boolean => {
  const requires = choice.getAttribute('Requires');
  if (!requires) return false; // Requires is mandatory on mc:Choice
  const prefixes = requires.split(/\s+/).filter(Boolean);
  if (prefixes.length === 0) return false;
  return prefixes.every(prefix => {
    const uri = choice.lookupNamespaceURI(prefix);
    return uri !== null && understood.has(uri);
  });
};

/** Replaces `target` with `replacements`, preserving document order. */
const spliceIn = (target: Element, replacements: Node[]): void => {
  const parent = target.parentNode;
  if (!parent) return;
  for (const node of replacements) {
    parent.insertBefore(node, target);
  }
  parent.removeChild(target);
};

/**
 * Resolves every `mc:AlternateContent` in the document to exactly one branch.
 *
 * Mutates and returns the document. Runs until no `mc:AlternateContent` remains, so
 * branches that themselves contain alternate content are handled - a real pattern,
 * since a `Choice` carrying a modern shape can wrap further conditional markup.
 */
export const resolveAlternateContent = (
  doc: Document,
  understoodNamespaces: Iterable<string> = MODERN_CONSUMER_NAMESPACES
): MceResult => {
  const understood = new Set(understoodNamespaces);
  const selections: MceSelection[] = [];

  // Bounded rather than `while (true)`: splicing children in place could in principle
  // reintroduce an element, and a malformed file should not hang the browser tab.
  const maxIterations = 10_000;
  let iterations = 0;

  for (;;) {
    const pending = doc.getElementsByTagNameNS(MCE_NAMESPACE, 'AlternateContent');
    const element = pending.item(0);
    if (!element) break;
    if (++iterations > maxIterations) break;

    const children = Array.from(element.children);
    const choices = children.filter(
      child => child.namespaceURI === MCE_NAMESPACE && child.localName === 'Choice'
    );
    const fallback = children.find(
      child => child.namespaceURI === MCE_NAMESPACE && child.localName === 'Fallback'
    );

    const selection: MceSelection = { path: pathOf(element), chose: 'nothing', rejected: [] };

    // First matching Choice wins; order is significant and is the author's preference
    // ranking, so this must not be reordered or treated as a set.
    let winner: Element | undefined;
    for (const choice of choices) {
      if (!winner && choiceIsUnderstood(choice, understood)) {
        winner = choice;
        selection.chose = 'choice';
        selection.requires = choice.getAttribute('Requires') ?? undefined;
      } else {
        selection.rejected.push(choice.getAttribute('Requires') ?? '');
      }
    }

    const chosen = winner ?? fallback;
    if (!winner && fallback) selection.chose = 'fallback';

    spliceIn(element, chosen ? Array.from(chosen.childNodes) : []);
    selections.push(selection);
  }

  return { document: doc, selections };
};

/**
 * Collects the namespace URIs declared ignorable anywhere in the document.
 *
 * `mc:Ignorable` names prefixes a consumer may silently drop if it does not recognise
 * them - it is why a modern file opens at all in an older reader instead of being
 * rejected. Reported rather than acted on: knowing a document declares `w14 w15 wp14`
 * ignorable tells you which Office generation wrote it, which is useful context for a
 * reader trying to understand the file.
 */
export const readIgnorableNamespaces = (doc: Document): Set<string> => {
  const ignorable = new Set<string>();
  const walk = (element: Element) => {
    const declared = element.getAttributeNS(MCE_NAMESPACE, 'Ignorable');
    if (declared) {
      for (const prefix of declared.split(/\s+/).filter(Boolean)) {
        const uri = element.lookupNamespaceURI(prefix);
        if (uri) ignorable.add(uri);
      }
    }
    for (const child of Array.from(element.children)) walk(child);
  };
  if (doc.documentElement) walk(doc.documentElement);
  return ignorable;
};

/**
 * Counts `mc:AlternateContent` elements without resolving them.
 *
 * Useful as a cheap "does this part need preprocessing?" probe before committing to
 * the full pass.
 */
export const countAlternateContent = (doc: Document): number =>
  doc.getElementsByTagNameNS(MCE_NAMESPACE, 'AlternateContent').length;
