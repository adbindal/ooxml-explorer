import { describe, it, expect } from 'vitest';
import {
  readOleObjects,
  findSilentlyBrokenOleObjects,
  oleDataIsPresent,
  computeOleEvidenceForMarkup
} from '../services/oleObjects';
import type { PackageParts } from '../services/packageIntegrity';

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
const P = 'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"';
const A = 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"';
const S = 'xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main"';
const O = 'xmlns:o="urn:schemas-microsoft-com:office:office"';
const V = 'xmlns:v="urn:schemas-microsoft-com:vml"';
const R = 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';

const rels = (body: string) =>
  `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${body}</Relationships>`;

const rel = (id: string, type: string, target: string, external = false) =>
  `<Relationship Id="${id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/${type}" Target="${target}"${
    external ? ' TargetMode="External"' : ''
  }/>`;

/** A Word document with one embedded object, its data present and a VML preview. */
const wordPackage = (overrides: Partial<PackageParts> = {}): PackageParts => ({
  'word/document.xml': `<?xml version="1.0"?><w:document ${W} ${O} ${V} ${R}><w:body><w:p><w:r><w:object w:dxaOrig="1440" w:dyaOrig="1440">
      <v:shape id="_x0000_i1025" type="#_x0000_t75"><v:imagedata r:id="rId5" o:title=""/></v:shape>
      <o:OLEObject Type="Embed" ProgID="Excel.Sheet.12" ShapeID="_x0000_i1025" DrawAspect="Content" ObjectID="_1234" r:id="rId4"/>
    </w:object></w:r></w:p></w:body></w:document>`,
  'word/_rels/document.xml.rels': rels(
    rel('rId4', 'oleObject', 'embeddings/oleObject1.bin') + rel('rId5', 'image', 'media/image1.emf')
  ),
  'word/embeddings/oleObject1.bin': 'BINARY',
  'word/media/image1.emf': 'IMAGE',
  ...overrides
});

const slidePackage = (overrides: Partial<PackageParts> = {}): PackageParts => ({
  'ppt/slides/slide1.xml': `<?xml version="1.0"?><p:sld ${P} ${A} ${R}><p:cSld><p:spTree>
      <p:graphicFrame><a:graphic><a:graphicData>
        <p:oleObj spid="_x0000_s1026" name="Worksheet" r:id="rId2" imgW="5486400" imgH="3200400" progId="Excel.Sheet.12">
          <p:embed/>
          <p:pic><p:blipFill><a:blip r:embed="rId3"/></p:blipFill></p:pic>
        </p:oleObj>
      </a:graphicData></a:graphic></p:graphicFrame>
    </p:spTree></p:cSld></p:sld>`,
  'ppt/slides/_rels/slide1.xml.rels': rels(
    rel('rId2', 'oleObject', '../embeddings/oleObject1.bin') + rel('rId3', 'image', '../media/image1.emf')
  ),
  'ppt/embeddings/oleObject1.bin': 'BINARY',
  'ppt/media/image1.emf': 'IMAGE',
  ...overrides
});

describe('the preview that hides breakage', () => {
  it('reports a healthy Word object with no problems', () => {
    const [object] = readOleObjects(wordPackage(), 'word/document.xml');

    expect(object.binding).toBe('embedded');
    expect(object.progId).toBe('Excel.Sheet.12');
    expect(object.dataTarget).toBe('word/embeddings/oleObject1.bin');
    expect(object.dataPartExists).toBe(true);
    expect(object.preview?.partExists).toBe(true);
    expect(object.problems).toEqual([]);
  });

  it('catches a missing embedding whose preview still renders', () => {
    // The headline case. The .bin is gone, the .emf is not, so the page looks
    // pixel-identical and the file opens without a warning.
    const parts = wordPackage();
    delete parts['word/embeddings/oleObject1.bin'];
    const [object] = readOleObjects(parts, 'word/document.xml');

    const problem = object.problems.find(p => p.code === 'ole/data-part-missing');
    expect(problem?.silent).toBe(true);
    expect(problem?.message).toContain('double-clicks');
    expect(object.preview?.partExists).toBe(true);
    expect(oleDataIsPresent(object)).toBe(false);
  });

  it('catches the same failure in a slide', () => {
    const parts = slidePackage();
    delete parts['ppt/embeddings/oleObject1.bin'];
    const [object] = readOleObjects(parts, 'ppt/slides/slide1.xml');

    expect(object.problems.map(p => p.code)).toContain('ole/data-part-missing');
    expect(object.preview?.partExists).toBe(true);
  });

  it('separates silent breakage from visible breakage', () => {
    // A missing preview is visible — something draws wrong. A missing embedding is
    // not. Only the second belongs on a "looks fine, is broken" list.
    const noPreview = readOleObjects(
      {
        ...wordPackage(),
        'word/document.xml': `<?xml version="1.0"?><w:document ${W} ${O} ${R}><w:body><w:p><w:r><w:object>
          <o:OLEObject Type="Embed" ProgID="Excel.Sheet.12" r:id="rId4"/>
        </w:object></w:r></w:p></w:body></w:document>`
      },
      'word/document.xml'
    );

    expect(noPreview[0].problems.find(p => p.code === 'ole/no-preview')?.silent).toBe(false);
    expect(findSilentlyBrokenOleObjects(noPreview)).toEqual([]);
  });

  it('lists only the objects that render correctly and are broken anyway', () => {
    const parts = wordPackage();
    delete parts['word/embeddings/oleObject1.bin'];
    const broken = findSilentlyBrokenOleObjects(readOleObjects(parts, 'word/document.xml'));

    expect(broken).toHaveLength(1);
  });

  it('does not treat a missing progId alone as silent breakage', () => {
    // Unhelpful, but the object still works. Reporting it as broken would bury the
    // failures that actually lose data.
    const parts = wordPackage({
      'word/document.xml': `<?xml version="1.0"?><w:document ${W} ${O} ${V} ${R}><w:body><w:p><w:r><w:object>
        <v:shape><v:imagedata r:id="rId5"/></v:shape>
        <o:OLEObject Type="Embed" r:id="rId4"/>
      </w:object></w:r></w:p></w:body></w:document>`
    });
    const objects = readOleObjects(parts, 'word/document.xml');

    expect(objects[0].problems.map(p => p.code)).toEqual(['ole/no-prog-id']);
    expect(findSilentlyBrokenOleObjects(objects)).toEqual([]);
  });
});

describe('three formats, three ways to say the same thing', () => {
  it('reads the Word binding from an attribute', () => {
    const [object] = readOleObjects(wordPackage(), 'word/document.xml');

    expect(object.binding).toBe('embedded');
    expect(object.bindingEvidence).toBe('o:OLEObject/@Type');
  });

  it('reads o:OLEObject attributes as unprefixed and PascalCase', () => {
    // Against the OOXML-wide convention: these are in no namespace and capitalised.
    // Looking for `type` in the office namespace finds nothing, twice over.
    const [object] = readOleObjects(wordPackage(), 'word/document.xml');

    expect(object.progId).toBe('Excel.Sheet.12');
    expect(object.shapeId).toBe('_x0000_i1025');
  });

  it('reads the PowerPoint binding from a child element', () => {
    const [object] = readOleObjects(slidePackage(), 'ppt/slides/slide1.xml');

    expect(object.binding).toBe('embedded');
    expect(object.bindingEvidence).toBe('p:embed / p:link child element');
  });

  it('infers the Excel binding from whether @link exists', () => {
    const sheet = (attrs: string) => ({
      'xl/worksheets/sheet1.xml': `<?xml version="1.0"?><x:worksheet ${S} ${R}><x:oleObjects>
        <x:oleObject progId="Word.Document.12" shapeId="1025" r:id="rId1" ${attrs}/>
      </x:oleObjects></x:worksheet>`,
      'xl/worksheets/_rels/sheet1.xml.rels': rels(rel('rId1', 'oleObject', '../embeddings/oleObject1.bin')),
      'xl/embeddings/oleObject1.bin': 'BINARY'
    });

    const embedded = readOleObjects(sheet(''), 'xl/worksheets/sheet1.xml');
    const linked = readOleObjects(sheet('link="Sheet1!R1C1"'), 'xl/worksheets/sheet1.xml');

    expect(embedded[0].binding).toBe('embedded');
    expect(linked[0].binding).toBe('linked');
    expect(linked[0].bindingEvidence).toBe('@link attribute presence');
  });

  it('reads the PowerPoint link binding', () => {
    const parts = slidePackage({
      'ppt/slides/slide1.xml': `<?xml version="1.0"?><p:sld ${P} ${A} ${R}><p:cSld><p:spTree>
        <p:oleObj r:id="rId2" progId="Excel.Sheet.12"><p:link updateAutomatic="1"/>
        <p:pic><p:blipFill><a:blip r:embed="rId3"/></p:blipFill></p:pic></p:oleObj>
      </p:spTree></p:cSld></p:sld>`,
      'ppt/slides/_rels/slide1.xml.rels': rels(
        rel('rId2', 'oleObject', 'file:///C:/data/book.xlsx', true) + rel('rId3', 'image', '../media/image1.emf')
      )
    });
    const [object] = readOleObjects(parts, 'ppt/slides/slide1.xml');

    expect(object.binding).toBe('linked');
    expect(object.dataIsExternal).toBe(true);
    expect(object.problems).toEqual([]);
  });
});

describe('binding and relationship must agree', () => {
  it('flags a linked object whose target is inside the package', () => {
    const parts = wordPackage({
      'word/document.xml': `<?xml version="1.0"?><w:document ${W} ${O} ${V} ${R}><w:body><w:p><w:r><w:object>
        <v:shape><v:imagedata r:id="rId5"/></v:shape>
        <o:OLEObject Type="Link" ProgID="Excel.Sheet.12" r:id="rId4"/>
      </w:object></w:r></w:p></w:body></w:document>`
    });
    const [object] = readOleObjects(parts, 'word/document.xml');

    expect(object.problems.find(p => p.code === 'ole/binding-mismatch')?.message).toContain('TargetMode');
  });

  it('flags an embedded object whose data is external to the package', () => {
    const parts = wordPackage({
      'word/_rels/document.xml.rels': rels(
        rel('rId4', 'oleObject', 'file:///C:/Users/someone/book.xlsx', true) + rel('rId5', 'image', 'media/image1.emf')
      )
    });
    const [object] = readOleObjects(parts, 'word/document.xml');

    const problem = object.problems.find(p => p.code === 'ole/binding-mismatch');
    expect(problem?.message).toContain('not in the package');
    expect(oleDataIsPresent(object)).toBeNull();
  });

  it('reports an unresolvable external link as unknown, not as missing', () => {
    // "We cannot check" and "it is missing" are different answers and only one is a
    // defect. A linked object's target lives outside the package by definition.
    const parts = wordPackage({
      'word/document.xml': `<?xml version="1.0"?><w:document ${W} ${O} ${V} ${R}><w:body><w:p><w:r><w:object>
        <v:shape><v:imagedata r:id="rId5"/></v:shape>
        <o:OLEObject Type="Link" ProgID="Excel.Sheet.12" r:id="rId4"/>
      </w:object></w:r></w:p></w:body></w:document>`,
      'word/_rels/document.xml.rels': rels(
        rel('rId4', 'oleObject', 'file:///C:/book.xlsx', true) + rel('rId5', 'image', 'media/image1.emf')
      )
    });
    const [object] = readOleObjects(parts, 'word/document.xml');

    expect(object.dataPartExists).toBeNull();
    expect(oleDataIsPresent(object)).toBeNull();
    expect(object.problems).toEqual([]);
  });
});

describe('undetermined bindings', () => {
  it('flags a Word object with no Type rather than assuming embedded', () => {
    const parts = wordPackage({
      'word/document.xml': `<?xml version="1.0"?><w:document ${W} ${O} ${V} ${R}><w:body><w:p><w:r><w:object>
        <v:shape><v:imagedata r:id="rId5"/></v:shape>
        <o:OLEObject ProgID="Excel.Sheet.12" r:id="rId4"/>
      </w:object></w:r></w:p></w:body></w:document>`
    });
    const [object] = readOleObjects(parts, 'word/document.xml');

    expect(object.binding).toBe('unknown');
    expect(object.problems.map(p => p.code)).toContain('ole/unknown-binding');
  });

  it('flags a slide object declaring both p:embed and p:link', () => {
    const parts = slidePackage({
      'ppt/slides/slide1.xml': `<?xml version="1.0"?><p:sld ${P} ${A} ${R}><p:cSld><p:spTree>
        <p:oleObj r:id="rId2" progId="Excel.Sheet.12"><p:embed/><p:link/>
        <p:pic><p:blipFill><a:blip r:embed="rId3"/></p:blipFill></p:pic></p:oleObj>
      </p:spTree></p:cSld></p:sld>`
    });
    const [object] = readOleObjects(parts, 'ppt/slides/slide1.xml');

    expect(object.binding).toBe('unknown');
    expect(object.problems.find(p => p.code === 'ole/unknown-binding')?.message).toContain('alternatives');
  });

  it('flags a slide object declaring neither', () => {
    const parts = slidePackage({
      'ppt/slides/slide1.xml': `<?xml version="1.0"?><p:sld ${P} ${A} ${R}><p:cSld><p:spTree>
        <p:oleObj r:id="rId2" progId="Excel.Sheet.12">
        <p:pic><p:blipFill><a:blip r:embed="rId3"/></p:blipFill></p:pic></p:oleObj>
      </p:spTree></p:cSld></p:sld>`
    });
    const [object] = readOleObjects(parts, 'ppt/slides/slide1.xml');

    expect(object.binding).toBe('unknown');
  });
});

describe('broken references', () => {
  it('reports an object that names no relationship', () => {
    const parts = wordPackage({
      'word/document.xml': `<?xml version="1.0"?><w:document ${W} ${O} ${V} ${R}><w:body><w:p><w:r><w:object>
        <v:shape><v:imagedata r:id="rId5"/></v:shape>
        <o:OLEObject Type="Embed" ProgID="Excel.Sheet.12"/>
      </w:object></w:r></w:p></w:body></w:document>`
    });
    const [object] = readOleObjects(parts, 'word/document.xml');

    expect(object.problems.map(p => p.code)).toContain('ole/no-data-reference');
  });

  it('reports a relationship id absent from the rels part', () => {
    const parts = wordPackage({
      'word/_rels/document.xml.rels': rels(rel('rId5', 'image', 'media/image1.emf'))
    });
    const [object] = readOleObjects(parts, 'word/document.xml');

    expect(object.problems.find(p => p.code === 'ole/relationship-missing')?.message).toContain('rId4');
  });

  it('reports a part with no relationships at all', () => {
    const parts = wordPackage();
    delete parts['word/_rels/document.xml.rels'];
    const [object] = readOleObjects(parts, 'word/document.xml');

    expect(object.problems.map(p => p.code)).toContain('ole/relationship-missing');
    expect(object.preview?.target).toBeNull();
  });
});

describe('tolerating input', () => {
  it('returns nothing for a part with no OLE objects', () => {
    expect(readOleObjects(wordPackage({ 'word/document.xml': `<?xml version="1.0"?><w:document ${W}><w:body/></w:document>` }), 'word/document.xml')).toEqual([]);
  });

  it('returns nothing for a part that is not in the package', () => {
    expect(readOleObjects(wordPackage(), 'word/nope.xml')).toEqual([]);
  });

  it('returns nothing rather than throwing on malformed XML', () => {
    expect(readOleObjects({ 'word/document.xml': '<w:document><unclosed>' }, 'word/document.xml')).toEqual([]);
  });

  it('finds objects in a Strict-namespace Word package', () => {
    // Strict does not merely swap the host — it DROPS THE YEAR:
    //   Transitional  http://schemas.openxmlformats.org/wordprocessingml/2006/main
    //   Strict        http://purl.oclc.org/ooxml/wordprocessingml/main
    // An earlier version of this test invented "purl.oclc.org/.../2006/main", which
    // does not exist. It passed against a matcher that only handled Transitional, so
    // the test asserted Strict support the code did not have.
    const parts: PackageParts = {
      'word/document.xml': `<?xml version="1.0"?><w:document xmlns:w="http://purl.oclc.org/ooxml/wordprocessingml/main" ${O} xmlns:r="http://purl.oclc.org/ooxml/officeDocument/relationships"><w:body><w:p><w:r><w:object>
        <o:OLEObject Type="Embed" ProgID="Excel.Sheet.12" r:id="rId4"/>
      </w:object></w:r></w:p></w:body></w:document>`,
      'word/_rels/document.xml.rels': rels(rel('rId4', 'oleObject', 'embeddings/oleObject1.bin')),
      'word/embeddings/oleObject1.bin': 'BINARY'
    };
    const [object] = readOleObjects(parts, 'word/document.xml');

    expect(object.binding).toBe('embedded');
    expect(object.dataPartExists).toBe(true);
  });

  it('finds objects in a Strict-namespace Excel package', () => {
    const parts: PackageParts = {
      'xl/worksheets/sheet1.xml': `<?xml version="1.0"?><x:worksheet xmlns:x="http://purl.oclc.org/ooxml/spreadsheetml/main" xmlns:r="http://purl.oclc.org/ooxml/officeDocument/relationships"><x:oleObjects>
        <x:oleObject progId="Word.Document.12" r:id="rId1"/>
      </x:oleObjects></x:worksheet>`,
      'xl/worksheets/_rels/sheet1.xml.rels': rels(rel('rId1', 'oleObject', '../embeddings/oleObject1.bin')),
      'xl/embeddings/oleObject1.bin': 'BINARY'
    };
    const [object] = readOleObjects(parts, 'xl/worksheets/sheet1.xml');

    expect(object.binding).toBe('embedded');
    expect(object.dataPartExists).toBe(true);
  });

  it('finds objects in a Strict-namespace PowerPoint package, preview included', () => {
    const parts: PackageParts = {
      'ppt/slides/slide1.xml': `<?xml version="1.0"?><p:sld xmlns:p="http://purl.oclc.org/ooxml/presentationml/main" xmlns:a="http://purl.oclc.org/ooxml/drawingml/main" xmlns:r="http://purl.oclc.org/ooxml/officeDocument/relationships"><p:cSld><p:spTree>
        <p:oleObj r:id="rId2" progId="Excel.Sheet.12"><p:embed/>
        <p:pic><p:blipFill><a:blip r:embed="rId3"/></p:blipFill></p:pic></p:oleObj>
      </p:spTree></p:cSld></p:sld>`,
      'ppt/slides/_rels/slide1.xml.rels': rels(
        rel('rId2', 'oleObject', '../embeddings/oleObject1.bin') + rel('rId3', 'image', '../media/image1.emf')
      ),
      'ppt/embeddings/oleObject1.bin': 'BINARY',
      'ppt/media/image1.emf': 'IMAGE'
    };
    const [object] = readOleObjects(parts, 'ppt/slides/slide1.xml');

    expect(object.binding).toBe('embedded');
    // The a:blip lookup is year-pinned too, so this pins the drawingml suffix as well.
    expect(object.preview?.partExists).toBe(true);
  });
});

describe('computeOleEvidenceForMarkup — panel wiring', () => {
  it('returns null when the part has no OLE objects', () => {
    const parts = wordPackage({
      'word/document.xml': `<?xml version="1.0"?><w:document ${W}><w:body/></w:document>`
    });

    expect(computeOleEvidenceForMarkup(parts)).toBeNull();
  });

  it('returns null when no part in the bundle can host an OLE object', () => {
    expect(computeOleEvidenceForMarkup({ 'word/styles.xml': '<w:styles/>' })).toBeNull();
  });

  it('names the owning application and where the binding was read from', () => {
    const evidence = computeOleEvidenceForMarkup(wordPackage());

    expect(evidence!.lines.some(l => l.includes('Excel.Sheet.12') && l.includes('o:OLEObject/@Type'))).toBe(true);
  });

  it('calls out objects that render correctly and are broken anyway', () => {
    const parts = wordPackage();
    delete parts['word/embeddings/oleObject1.bin'];
    const evidence = computeOleEvidenceForMarkup(parts);

    expect(evidence!.lines.some(l => l.includes('render exactly as intended and are broken anyway'))).toBe(true);
  });

  it('does not claim silent breakage when everything resolves', () => {
    const evidence = computeOleEvidenceForMarkup(wordPackage());

    expect(evidence!.lines.some(l => l.includes('broken anyway'))).toBe(false);
  });

  it('caps the claim: a present binary is not a validated binary', () => {
    // The relationship resolves and the part exists. Whether the bytes are a valid
    // compound file for the declared progId is a different question, and asserting
    // it would be exactly the kind of unearned confidence the tier exists to prevent.
    const evidence = computeOleEvidenceForMarkup(wordPackage());

    expect(evidence!.unresolved.some(u => u.includes('not opened'))).toBe(true);
  });

  it('sends an unresolvable external target to unresolved, not to problems', () => {
    const parts = wordPackage({
      'word/document.xml': `<?xml version="1.0"?><w:document ${W} ${O} ${V} ${R}><w:body><w:p><w:r><w:object>
        <v:shape><v:imagedata r:id="rId5"/></v:shape>
        <o:OLEObject Type="Link" ProgID="Excel.Sheet.12" r:id="rId4"/>
      </w:object></w:r></w:p></w:body></w:document>`,
      'word/_rels/document.xml.rels': rels(
        rel('rId4', 'oleObject', 'file:///C:/book.xlsx', true) + rel('rId5', 'image', 'media/image1.emf')
      )
    });
    const evidence = computeOleEvidenceForMarkup(parts);

    expect(evidence!.unresolved.some(u => u.includes('cannot be checked from inside the package'))).toBe(true);
  });

  it('works for a slide as well as a document', () => {
    const evidence = computeOleEvidenceForMarkup(slidePackage());

    expect(evidence!.lines[0]).toContain('ppt/slides/slide1.xml');
    expect(evidence!.lines.some(l => l.includes('p:embed / p:link child element'))).toBe(true);
  });
});
