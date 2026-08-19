import { describe, it, expect } from 'vitest';
import {
  readMedia,
  readMediaTriggers,
  findSilentlyBrokenMedia,
  mediaDataIsPresent,
  mediaFindings,
  computeMediaEvidenceForMarkup
} from '../services/pptMedia';
import type { PackageParts } from '../services/packageIntegrity';

const P = 'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"';
const A = 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"';
const R = 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
const P14 = 'xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main"';

const rels = (body: string) =>
  `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${body}</Relationships>`;

const rel = (id: string, type: string, target: string, external = false) =>
  `<Relationship Id="${id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/${type}" Target="${target}"${
    external ? ' TargetMode="External"' : ''
  }/>`;

const contentTypes = (body = '<Default Extension="mp4" ContentType="video/mp4"/>') =>
  `<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
    <Default Extension="png" ContentType="image/png"/>${body}</Types>`;

/**
 * A slide with one video: the poster frame, the classic `a:videoFile r:link`, and the
 * Office 2010 `p14:media r:embed` PowerPoint writes beside it for the same clip.
 */
const slide = (body: string) =>
  `<?xml version="1.0"?><p:sld ${P} ${A} ${R} ${P14}><p:cSld><p:spTree>${body}</p:spTree></p:cSld></p:sld>`;

const videoPic = (opts: { link?: string; embed?: string; poster?: string; contentType?: string; id?: string } = {}) => {
  const link = opts.link ?? 'rId2';
  const embed = opts.embed ?? 'rId3';
  const poster = opts.poster ?? 'rId4';
  const ct = opts.contentType === undefined ? ' contentType="video/mp4"' : ` contentType="${opts.contentType}"`;
  return `<p:pic>
      <p:nvPicPr>
        <p:cNvPr id="${opts.id ?? '5'}" name="Keynote.mp4"/><p:cNvPicPr/>
        <p:nvPr>
          <a:videoFile r:link="${link}"${ct}/>
          <p:extLst><p:ext uri="{DAA4B4D4-6D71-4841-9C94-3DE7FCFB9230}"><p14:media r:embed="${embed}"/></p:ext></p:extLst>
        </p:nvPr>
      </p:nvPicPr>
      <p:blipFill><a:blip r:embed="${poster}"/></p:blipFill>
    </p:pic>`;
};

const deck = (overrides: Partial<PackageParts> = {}): PackageParts => ({
  '[Content_Types].xml': contentTypes(),
  'ppt/slides/slide1.xml': slide(videoPic()),
  'ppt/slides/_rels/slide1.xml.rels': rels(
    rel('rId2', 'video', '../media/media1.mp4') +
      rel('rId3', 'media', '../media/media1.mp4') +
      rel('rId4', 'image', '../media/image1.png')
  ),
  'ppt/media/media1.mp4': 'MP4',
  'ppt/media/image1.png': 'PNG',
  ...overrides
});

describe('the poster frame that hides a missing clip', () => {
  it('reports a healthy embedded video with no problems', () => {
    const [clip] = readMedia(deck(), 'ppt/slides/slide1.xml');

    expect(clip.kind).toBe('video');
    expect(clip.shapeId).toBe('5');
    expect(clip.shapeName).toBe('Keynote.mp4');
    expect(clip.references.map(r => r.element)).toEqual(['a:videoFile', 'p14:media']);
    expect(clip.poster?.partExists).toBe(true);
    expect(mediaDataIsPresent(clip)).toBe(true);
    expect(clip.problems).toEqual([]);
  });

  it('catches a missing media part whose poster still renders', () => {
    // The headline case. The .mp4 is gone, the .png is not, so the slide is
    // pixel-identical and the deck opens without a warning.
    const parts = deck();
    delete parts['ppt/media/media1.mp4'];
    const [clip] = readMedia(parts, 'ppt/slides/slide1.xml');

    const problem = clip.problems.find(p => p.code === 'media/media-part-missing');
    expect(problem?.silent).toBe(true);
    expect(problem?.severity).toBe('error');
    expect(problem?.message).toContain('presses play');
    expect(clip.poster?.partExists).toBe(true);
    expect(mediaDataIsPresent(clip)).toBe(false);
  });

  it('separates the silent break from the visible one', () => {
    // A missing poster is visible — the shape draws as empty space. A missing clip is
    // not. Only the second belongs on a "looks fine, is broken" list.
    const noPoster = readMedia(
      deck({ 'ppt/slides/slide1.xml': slide(videoPic().replace(/<p:blipFill>.*<\/p:blipFill>/s, '')) }),
      'ppt/slides/slide1.xml'
    );

    expect(noPoster[0].poster).toBeNull();
    expect(noPoster[0].problems.find(p => p.code === 'media/no-poster-frame')?.silent).toBe(false);
    expect(findSilentlyBrokenMedia(noPoster)).toEqual([]);
  });

  it('reports a poster whose image part is gone as visible, not silent', () => {
    const parts = deck();
    delete parts['ppt/media/image1.png'];
    const [clip] = readMedia(parts, 'ppt/slides/slide1.xml');

    const problem = clip.problems.find(p => p.code === 'media/poster-part-missing');
    expect(problem?.silent).toBe(false);
    expect(clip.poster?.partExists).toBe(false);
    // The clip itself is fine — the break is only in the still.
    expect(mediaDataIsPresent(clip)).toBe(true);
  });

  it('lists only the clips that render correctly and are broken anyway', () => {
    const parts = deck();
    delete parts['ppt/media/media1.mp4'];

    expect(findSilentlyBrokenMedia(readMedia(parts, 'ppt/slides/slide1.xml'))).toHaveLength(1);
  });
});

describe('linked media is a relationship property, not an attribute name', () => {
  it('reports an externally linked clip as an error even though @r:link was used either way', () => {
    // a:videoFile has ONLY r:link in the schema, so the attribute name says nothing
    // about whether the clip is in the package. TargetMode does.
    const external = deck({
      'ppt/slides/slide1.xml': slide(videoPic({ embed: '' })),
      'ppt/slides/_rels/slide1.xml.rels': rels(
        rel('rId2', 'video', 'file:///C:/Users/someone/Keynote.mp4', true) + rel('rId4', 'image', '../media/image1.png')
      )
    });
    const [clip] = readMedia(external, 'ppt/slides/slide1.xml');

    const problem = clip.problems.find(p => p.code === 'media/external-link');
    expect(problem?.severity).toBe('error');
    expect(problem?.silent).toBe(true);
    expect(problem?.message).toContain('not in the package');
    expect(clip.references[0].attribute).toBe('r:link');
    expect(clip.references[0].isExternal).toBe(true);
  });

  it('does not call an internal r:link external', () => {
    // The healthy fixture uses r:link for a packaged clip, which is exactly what
    // PowerPoint 2010+ writes. Reporting that would fire on almost every real deck.
    const [clip] = readMedia(deck(), 'ppt/slides/slide1.xml');

    expect(clip.references[0].attribute).toBe('r:link');
    expect(clip.references[0].isExternal).toBe(false);
    expect(clip.problems.map(p => p.code)).not.toContain('media/external-link');
  });

  it('downgrades an external link when a packaged copy on the same picture resolves', () => {
    // p14:media r:embed resolves inside the package, so PowerPoint 2010+ plays that and
    // the deck is still self-contained. Calling this an error would be over-claiming.
    const parts = deck({
      'ppt/slides/_rels/slide1.xml.rels': rels(
        rel('rId2', 'video', 'file:///C:/Keynote.mp4', true) +
          rel('rId3', 'media', '../media/media1.mp4') +
          rel('rId4', 'image', '../media/image1.png')
      )
    });
    const [clip] = readMedia(parts, 'ppt/slides/slide1.xml');

    expect(clip.problems.map(p => p.code)).toEqual(['media/external-link-shadowed']);
    expect(clip.problems[0].severity).toBe('note');
    expect(mediaDataIsPresent(clip)).toBe(true);
  });

  it('does not downgrade when the packaged copy is the one that is missing', () => {
    // Same markup as above, but the embedded part is gone: nothing resolves inside the
    // package any more, so the external link is back to being the whole story.
    const parts = deck({
      'ppt/slides/_rels/slide1.xml.rels': rels(
        rel('rId2', 'video', 'file:///C:/Keynote.mp4', true) +
          rel('rId3', 'media', '../media/media1.mp4') +
          rel('rId4', 'image', '../media/image1.png')
      )
    });
    delete parts['ppt/media/media1.mp4'];
    const codes = readMedia(parts, 'ppt/slides/slide1.xml')[0].problems.map(p => p.code);

    expect(codes).toContain('media/external-link');
    expect(codes).toContain('media/media-part-missing');
    expect(codes).not.toContain('media/external-link-shadowed');
  });

  it('treats an external clip as unknowable rather than missing', () => {
    const parts = deck({
      'ppt/slides/slide1.xml': slide(videoPic({ embed: '' })),
      'ppt/slides/_rels/slide1.xml.rels': rels(
        rel('rId2', 'video', 'https://example.invalid/clip.mp4', true) + rel('rId4', 'image', '../media/image1.png')
      )
    });
    const [clip] = readMedia(parts, 'ppt/slides/slide1.xml');

    expect(clip.references[0].partExists).toBeNull();
    expect(mediaDataIsPresent(clip)).toBeNull();
    expect(clip.problems.map(p => p.code)).not.toContain('media/media-part-missing');
  });
});

describe('the media family', () => {
  const pic = (inner: string, blipFill = '<p:blipFill><a:blip r:embed="rId4"/></p:blipFill>') =>
    `<p:pic><p:nvPicPr><p:cNvPr id="7" name="Clip"/><p:cNvPicPr/><p:nvPr>${inner}</p:nvPr></p:nvPicPr>${blipFill}</p:pic>`;

  const soundDeck = (inner: string) =>
    deck({
      'ppt/slides/slide1.xml': slide(pic(inner)),
      'ppt/slides/_rels/slide1.xml.rels': rels(
        rel('rId2', 'audio', '../media/audio1.wav') + rel('rId4', 'image', '../media/image1.png')
      ),
      'ppt/media/audio1.wav': 'WAV'
    });

  it('reads a:audioFile as audio through r:link', () => {
    const [clip] = readMedia(soundDeck('<a:audioFile r:link="rId2"/>'), 'ppt/slides/slide1.xml');

    expect(clip.kind).toBe('audio');
    expect(clip.references[0].attribute).toBe('r:link');
    expect(clip.problems).toEqual([]);
  });

  it('reads a:wavAudioFile through r:embed, the one member that uses it', () => {
    const [clip] = readMedia(soundDeck('<a:wavAudioFile r:embed="rId2"/>'), 'ppt/slides/slide1.xml');

    expect(clip.kind).toBe('audio');
    expect(clip.references[0].attribute).toBe('r:embed');
    expect(clip.references[0].partExists).toBe(true);
  });

  it('reads a:quickTimeFile as video', () => {
    const [clip] = readMedia(soundDeck('<a:quickTimeFile r:link="rId2"/>'), 'ppt/slides/slide1.xml');

    expect(clip.kind).toBe('video');
  });

  it('finds a p14:media clip with no classic sibling element', () => {
    const [clip] = readMedia(
      soundDeck('<p:extLst><p:ext uri="{X}"><p14:media r:embed="rId2"/></p:ext></p:extLst>'),
      'ppt/slides/slide1.xml'
    );

    expect(clip.references.map(r => r.element)).toEqual(['p14:media']);
    expect(clip.references[0].attribute).toBe('r:embed');
  });

  it('reads p14:media/@r:link as a link, not an embed', () => {
    const [clip] = readMedia(
      soundDeck('<p:extLst><p:ext uri="{X}"><p14:media r:link="rId2"/></p:ext></p:extLst>'),
      'ppt/slides/slide1.xml'
    );

    expect(clip.references[0].attribute).toBe('r:link');
  });

  it('ignores an ordinary picture that carries no media at all', () => {
    expect(readMedia(soundDeck('<p:ph idx="1"/>'), 'ppt/slides/slide1.xml')).toEqual([]);
  });

  it('ignores an OLE preview picture, which has no p:nvPicPr', () => {
    const parts = deck({
      'ppt/slides/slide1.xml': slide(
        '<p:graphicFrame><a:graphic><a:graphicData><p:oleObj r:id="rId9"><p:embed/>' +
          '<p:pic><p:blipFill><a:blip r:embed="rId4"/></p:blipFill></p:pic>' +
          '</p:oleObj></a:graphicData></a:graphic></p:graphicFrame>'
      )
    });

    expect(readMedia(parts, 'ppt/slides/slide1.xml')).toEqual([]);
  });
});

describe('broken references', () => {
  it('reports a relationship id the rels part never declares', () => {
    const parts = deck({
      'ppt/slides/_rels/slide1.xml.rels': rels(rel('rId3', 'media', '../media/media1.mp4') + rel('rId4', 'image', '../media/image1.png'))
    });
    const problem = readMedia(parts, 'ppt/slides/slide1.xml')[0].problems.find(
      p => p.code === 'media/relationship-missing'
    );

    expect(problem?.message).toContain('rId2');
    expect(problem?.silent).toBe(true);
  });

  it('reports a media element that names no relationship at all', () => {
    const parts = deck({
      'ppt/slides/slide1.xml': slide(
        '<p:pic><p:nvPicPr><p:cNvPr id="5" name="Clip"/><p:cNvPicPr/><p:nvPr><a:videoFile/></p:nvPr></p:nvPicPr>' +
          '<p:blipFill><a:blip r:embed="rId4"/></p:blipFill></p:pic>'
      )
    });
    const codes = readMedia(parts, 'ppt/slides/slide1.xml')[0].problems.map(p => p.code);

    expect(codes).toContain('media/no-media-reference');
    expect(codes).not.toContain('media/relationship-missing');
  });

  it('reports every reference when the whole rels part is absent', () => {
    const parts = deck();
    delete parts['ppt/slides/_rels/slide1.xml.rels'];
    const [clip] = readMedia(parts, 'ppt/slides/slide1.xml');

    expect(clip.problems.filter(p => p.code === 'media/relationship-missing')).toHaveLength(2);
    expect(clip.poster?.partExists).toBe(false);
  });
});

describe('declared content type against the package', () => {
  it('flags markup that disagrees with [Content_Types].xml', () => {
    const parts = deck({ '[Content_Types].xml': contentTypes('<Default Extension="mp4" ContentType="video/quicktime"/>') });
    const problem = readMedia(parts, 'ppt/slides/slide1.xml')[0].problems.find(
      p => p.code === 'media/content-type-mismatch'
    );

    expect(problem?.message).toContain('video/quicktime');
    expect(problem?.silent).toBe(true);
  });

  it('accepts a case difference and a codec parameter as agreement', () => {
    const parts = deck({ '[Content_Types].xml': contentTypes('<Default Extension="mp4" ContentType="VIDEO/MP4; codecs=avc1"/>') });

    expect(readMedia(parts, 'ppt/slides/slide1.xml')[0].problems).toEqual([]);
  });

  it('prefers an Override to the Default for the same part', () => {
    const parts = deck({
      '[Content_Types].xml': contentTypes(
        '<Default Extension="mp4" ContentType="video/mp4"/><Override PartName="/ppt/media/media1.mp4" ContentType="video/x-msvideo"/>'
      )
    });

    expect(readMedia(parts, 'ppt/slides/slide1.xml')[0].problems.map(p => p.code)).toContain(
      'media/content-type-mismatch'
    );
  });

  it('says nothing when the markup declares no content type', () => {
    // @contentType is optional, and absent on a:quickTimeFile by schema. Silence is the
    // correct answer, not a guess from the file extension.
    const parts = deck({
      'ppt/slides/slide1.xml': slide(videoPic({ contentType: undefined }).replace(' contentType="video/mp4"', '')),
      '[Content_Types].xml': contentTypes('<Default Extension="mp4" ContentType="video/quicktime"/>')
    });

    expect(readMedia(parts, 'ppt/slides/slide1.xml')[0].problems).toEqual([]);
  });

  it('says nothing when the package declares no content type for the part', () => {
    // An undeclared part is packageIntegrity's finding, not this one's. Reporting it
    // here would double-count it in the panel.
    const parts = deck({ '[Content_Types].xml': contentTypes('') });

    expect(readMedia(parts, 'ppt/slides/slide1.xml')[0].problems).toEqual([]);
  });
});

describe('timing triggers that point at nothing', () => {
  const withTiming = (spid: string, picId = '5') =>
    deck({
      'ppt/slides/slide1.xml': slide(videoPic({ id: picId })).replace(
        '</p:cSld>',
        `</p:cSld><p:timing><p:tnLst><p:par><p:cTn><p:childTnLst>
            <p:video><p:cMediaNode><p:cTn/><p:tgtEl><p:spTgt spid="${spid}"/></p:tgtEl></p:cMediaNode></p:video>
          </p:childTnLst></p:cTn></p:par></p:tnLst></p:timing>`
      )
    });

  it('binds a trigger to the picture it starts', () => {
    const [clip] = readMedia(withTiming('5'), 'ppt/slides/slide1.xml');

    expect(clip.triggers).toHaveLength(1);
    expect(clip.triggers[0].label).toBe('p:video');
    expect(clip.triggers[0].targetExists).toBe(true);
    expect(clip.triggers[0].problems).toEqual([]);
  });

  it('flags a trigger whose target shape id is not on the slide', () => {
    const [trigger] = readMediaTriggers(withTiming('99'), 'ppt/slides/slide1.xml');

    expect(trigger.targetExists).toBe(false);
    const problem = trigger.problems.find(p => p.code === 'media/dangling-trigger');
    expect(problem?.silent).toBe(true);
    expect(problem?.subject?.spid).toBe('99');
  });

  it('does not match a shape id by prefix or by numeric coincidence', () => {
    // spid="5" and a shape with id="55" are different shapes. A substring or
    // startsWith comparison would call this trigger live.
    const [trigger] = readMediaTriggers(withTiming('5', '55'), 'ppt/slides/slide1.xml');

    expect(trigger.targetExists).toBe(false);
  });

  it('declines to judge the Office 2007 string form of @spid', () => {
    // @spid is StringValidator in Office2007 and numeric ST_DrawingElementId from
    // Office2010 on, while p:cNvPr/@id is always UInt32. A 2007 deck writing
    // spid="_x0000_s1026" matches no shape id BY DESIGN, so reporting it as dangling
    // would be a confident wrong answer on every deck of that vintage.
    const [trigger] = readMediaTriggers(withTiming('_x0000_s1026'), 'ppt/slides/slide1.xml');

    expect(trigger.targetShapeId).toBe('_x0000_s1026');
    expect(trigger.targetExists).toBeNull();
    expect(trigger.problems).toEqual([]);
  });

  it('finds p:audio triggers as well as p:video', () => {
    const parts = withTiming('99');
    parts['ppt/slides/slide1.xml'] = parts['ppt/slides/slide1.xml']
      .replace('<p:video>', '<p:audio>')
      .replace('</p:video>', '</p:audio>');
    const [trigger] = readMediaTriggers(parts, 'ppt/slides/slide1.xml');

    expect(trigger.label).toBe('p:audio');
    expect(trigger.problems.map(p => p.code)).toEqual(['media/dangling-trigger']);
  });

  it('ignores a p:video element that is not inside p:timing', () => {
    // Only the timing tree starts playback; anything else called p:video is not a trigger.
    const parts = deck({
      'ppt/slides/slide1.xml': slide(
        `${videoPic()}<p:video><p:cMediaNode><p:tgtEl><p:spTgt spid="99"/></p:tgtEl></p:cMediaNode></p:video>`
      )
    });

    expect(readMediaTriggers(parts, 'ppt/slides/slide1.xml')).toEqual([]);
  });

  it('unions media and trigger findings for the registry', () => {
    const codes = mediaFindings(withTiming('99'), 'ppt/slides/slide1.xml').map(f => f.code);

    expect(codes).toEqual(['media/dangling-trigger']);
  });
});

describe('tolerating input', () => {
  it('returns nothing for a slide with no media', () => {
    expect(readMedia(deck({ 'ppt/slides/slide1.xml': slide('') }), 'ppt/slides/slide1.xml')).toEqual([]);
  });

  it('returns nothing for a part that is not in the package', () => {
    expect(readMedia(deck(), 'ppt/slides/slide9.xml')).toEqual([]);
    expect(readMediaTriggers(deck(), 'ppt/slides/slide9.xml')).toEqual([]);
  });

  it('returns nothing rather than throwing on malformed XML', () => {
    expect(readMedia({ 'ppt/slides/slide1.xml': '<p:sld><unclosed>' }, 'ppt/slides/slide1.xml')).toEqual([]);
    expect(readMediaTriggers({ 'ppt/slides/slide1.xml': '<p:sld><unclosed>' }, 'ppt/slides/slide1.xml')).toEqual([]);
  });

  it('survives a rels part that does not parse', () => {
    const parts = deck({ 'ppt/slides/_rels/slide1.xml.rels': '<Relationships><oops>' });
    const [clip] = readMedia(parts, 'ppt/slides/slide1.xml');

    expect(clip.problems.map(p => p.code)).toContain('media/relationship-missing');
  });

  it('ignores a media element in the wrong namespace', () => {
    // p:videoFile is not a:videoFile. Namespaces are compared by exact equality because
    // conformance.ts has already mapped ISO Strict to Transitional.
    const parts = deck({
      'ppt/slides/slide1.xml': slide(
        '<p:pic><p:nvPicPr><p:cNvPr id="5" name="Clip"/><p:cNvPicPr/><p:nvPr><p:videoFile r:link="rId2"/></p:nvPr></p:nvPicPr>' +
          '<p:blipFill><a:blip r:embed="rId4"/></p:blipFill></p:pic>'
      )
    });

    expect(readMedia(parts, 'ppt/slides/slide1.xml')).toEqual([]);
  });

  it('reads a Strict package once conformance has normalised it', () => {
    // The pipeline rewrites purl.oclc.org URIs before any analyzer runs, so what arrives
    // here is Transitional. p14 is a Microsoft extension namespace and is unaffected by
    // that mapping — it is spelled the same in both conformance classes.
    const [clip] = readMedia(deck(), 'ppt/slides/slide1.xml');

    expect(clip.references.map(r => r.element)).toContain('p14:media');
  });

  it('finds several clips on one slide and keeps their problems apart', () => {
    const parts = deck({
      'ppt/slides/slide1.xml': slide(
        videoPic({ id: '5' }) + videoPic({ id: '6', link: 'rId7', embed: 'rId8', poster: 'rId4' })
      )
    });
    const clips = readMedia(parts, 'ppt/slides/slide1.xml');

    expect(clips).toHaveLength(2);
    expect(clips[0].problems).toEqual([]);
    expect(clips[1].problems.filter(p => p.code === 'media/relationship-missing')).toHaveLength(2);
  });
});

describe('computeMediaEvidenceForMarkup — panel wiring', () => {
  it('returns null when no part in the bundle can host media', () => {
    expect(computeMediaEvidenceForMarkup({ 'ppt/presentation.xml': '<p:presentation/>' })).toBeNull();
  });

  it('returns null when the slide has no media', () => {
    expect(computeMediaEvidenceForMarkup(deck({ 'ppt/slides/slide1.xml': slide('') }))).toBeNull();
  });

  it('skips past a host part that has no media to one that does', () => {
    // Key order is insertion order, so taking the first host part blind would report
    // "no media" for this deck purely because the layout was bundled first.
    const parts: PackageParts = {
      'ppt/slideLayouts/slideLayout1.xml': slide(''),
      ...deck()
    };
    const evidence = computeMediaEvidenceForMarkup(parts);

    expect(evidence!.lines[0]).toContain('ppt/slides/slide1.xml');
  });

  it('names the shape and how the clip is referenced', () => {
    const evidence = computeMediaEvidenceForMarkup(deck());

    expect(evidence!.lines.some(l => l.includes('Keynote.mp4') && l.includes('a:videoFile/r:link'))).toBe(true);
    expect(evidence!.lines.some(l => l.includes('poster frame stands in front of it'))).toBe(true);
  });

  it('calls out clips that render correctly and are broken anyway', () => {
    const parts = deck();
    delete parts['ppt/media/media1.mp4'];
    const evidence = computeMediaEvidenceForMarkup(parts);

    expect(evidence!.lines.some(l => l.includes('render exactly as intended and are broken anyway'))).toBe(true);
  });

  it('does not claim silent breakage when everything resolves', () => {
    expect(computeMediaEvidenceForMarkup(deck())!.lines.some(l => l.includes('broken anyway'))).toBe(false);
  });

  it('caps the claim: a present media part is not a playable one', () => {
    expect(computeMediaEvidenceForMarkup(deck())!.unresolved.some(u => u.includes('never decoded'))).toBe(true);
  });

  it('sends an unresolvable external target to unresolved rather than claiming it is missing', () => {
    const parts = deck({
      'ppt/slides/slide1.xml': slide(videoPic({ embed: '' })),
      'ppt/slides/_rels/slide1.xml.rels': rels(
        rel('rId2', 'video', 'file:///C:/Keynote.mp4', true) + rel('rId4', 'image', '../media/image1.png')
      )
    });
    const evidence = computeMediaEvidenceForMarkup(parts);

    expect(evidence!.unresolved.some(u => u.includes('cannot be checked from inside the package'))).toBe(true);
    expect(evidence!.unresolved.some(u => u.includes('never decoded'))).toBe(false);
  });

  it('states the Office 2007 @spid limit instead of silently skipping it', () => {
    const parts = deck({
      'ppt/slides/slide1.xml': slide(videoPic()).replace(
        '</p:cSld>',
        `</p:cSld><p:timing><p:tnLst><p:par><p:cTn><p:childTnLst>
          <p:video><p:cMediaNode><p:tgtEl><p:spTgt spid="_x0000_s1026"/></p:tgtEl></p:cMediaNode></p:video>
        </p:childTnLst></p:cTn></p:par></p:tnLst></p:timing>`
      )
    });
    const evidence = computeMediaEvidenceForMarkup(parts);

    expect(evidence!.unresolved.some(u => u.includes('Office 2007 string form'))).toBe(true);
  });

  it('reports a dangling trigger in the panel even though no clip owns it', () => {
    const parts = deck({
      'ppt/slides/slide1.xml': slide(videoPic()).replace(
        '</p:cSld>',
        `</p:cSld><p:timing><p:tnLst><p:par><p:cTn><p:childTnLst>
          <p:video><p:cMediaNode><p:tgtEl><p:spTgt spid="99"/></p:tgtEl></p:cMediaNode></p:video>
        </p:childTnLst></p:cTn></p:par></p:tnLst></p:timing>`
      )
    });
    const evidence = computeMediaEvidenceForMarkup(parts);

    expect(evidence!.lines.some(l => l.includes('shape id 99'))).toBe(true);
  });
});
