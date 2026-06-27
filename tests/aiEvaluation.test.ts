import { describe, it, expect, beforeEach } from '../services/browserTestRunner';
import { explainElement, getAiClient } from '../services/aiService';
import { useAppStore } from '../store/appStore';

/* eslint-disable @typescript-eslint/no-explicit-any */
const { mockKnowledgeBase } = vi.hoisted(() => {
  return {
    mockKnowledgeBase: [
      { tag: 'tcW', namespace: 'w', domain: 'docx', definition: 'Table Cell Width. Specifies the width of the table cell.', attributes: [], parents: [] },
      { tag: 'cantSplit', namespace: 'w', domain: 'docx', definition: 'Table Row Cannot Split. Specifies that the row must not be split across pages.', attributes: [], parents: [] },
      { tag: 'sheetData', namespace: 'r', domain: 'xlsx', definition: 'Sheet Data. The grid container for all rows and cells in the worksheet.', attributes: [], parents: [] },
      { tag: 'sld', namespace: 'p', domain: 'pptx', definition: 'Slide. The root element representing a single slide within the presentation.', attributes: [], parents: [] },
      { tag: 'Relationship', namespace: 'r', domain: 'shared', definition: 'Relationship Definition. Specifies a link between a source and target.', attributes: [], parents: [] }
    ]
  };
});

vi.mock('../services/storageService', () => {
  return {
    querySchemaFromStorage: vi.fn(async (tag: string, domain: string) => {
      const matches = mockKnowledgeBase.filter((m: any) => m.tag === tag);
      const bestMatch = matches.find((m: any) => m.domain === domain || m.domain === 'shared');
      return bestMatch || null;
    }),
    searchSchemasInStorage: vi.fn(async (keyword: string, domain: string) => {
      const cleanKeyword = keyword.toLowerCase().trim();
      return mockKnowledgeBase.filter((m: any) => {
        const inDomain = m.domain === domain || m.domain === 'shared';
        if (!inDomain) return false;
        const matchesTag = m.tag.toLowerCase().includes(cleanKeyword);
        const matchesDef = m.definition.toLowerCase().includes(cleanKeyword);
        return matchesTag || matchesDef;
      });
    })
  };
});
/* eslint-enable @typescript-eslint/no-explicit-any */

interface GoldenCase {
  tagName: string;
  rawXml: string;
  fileType: 'docx' | 'xlsx' | 'pptx';
  expectedKeywords: string[];
}

/**
 * The Golden Dataset for OOXML Explainer AI Evaluation.
 * Used to verify prompt completeness and keyword coverage.
 */
const GOLDEN_DATASET: GoldenCase[] = [
  {
    tagName: 'cantSplit',
    rawXml: '<w:cantSplit />',
    fileType: 'docx',
    expectedKeywords: ['split', 'row', 'page']
  },
  {
    tagName: 'sheetData',
    rawXml: '<sheetData />',
    fileType: 'xlsx',
    expectedKeywords: ['grid', 'cell', 'row', 'worksheet']
  },
  {
    tagName: 'sld',
    rawXml: '<p:sld />',
    fileType: 'pptx',
    expectedKeywords: ['slide', 'presentation']
  }
];

describe('AI EVAL Pipeline (Layer 5)', () => {
  beforeEach(() => {
    useAppStore.setState({
      ui: {
        ...useAppStore.getState().ui,
        aiProvider: 'gemini-cloud',
        dlpMode: false
      }
    });
    localStorage.setItem('ooxml_explorer_api_key', 'mock-eval-key');
  });

  it('runs the golden dataset through the prompt pipeline and verifies RAG injection', async () => {
    const ai = getAiClient('mock-eval-key');
    
    for (const testCase of GOLDEN_DATASET) {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const generateSpy = vi.spyOn(ai.models, 'generateContent').mockImplementation(async (args: any) => {
        // Assert that the prompt contains the tag name
        expect(args.contents).toContain(testCase.tagName);
        
        // Assert that the prompt contains the RAG injected keywords
        for (const keyword of testCase.expectedKeywords) {
          const lowerContents = args.contents.toLowerCase();
          expect(lowerContents.includes(keyword) || args.config?.systemInstruction?.toLowerCase().includes(keyword)).toBe(true);
        }

        return {
          text: `Verified explanation for ${testCase.tagName} containing: ${testCase.expectedKeywords.join(', ')}`
        };
      });

      const explanation = await explainElement(testCase.tagName, testCase.rawXml, testCase.fileType);
      
      // Verify that the returned explanation contains the expected evaluation keywords
      for (const keyword of testCase.expectedKeywords) {
        expect(explanation.toLowerCase()).toContain(keyword);
      }

      generateSpy.mockRestore?.();
    }

    localStorage.removeItem('ooxml_explorer_api_key');
  });
});
