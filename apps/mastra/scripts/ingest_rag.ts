import { mastra } from '../src';

// Define the reference tags to ingest during calibration
const REFERENCE_TAGS = [
  { tag: 'document', namespace: 'w', domain: 'docx' },
  { tag: 'body', namespace: 'w', domain: 'docx' },
  { tag: 'p', namespace: 'w', domain: 'docx' },
  { tag: 'r', namespace: 'w', domain: 'docx' },
  { tag: 't', namespace: 'w', domain: 'docx' },
  { tag: 'tbl', namespace: 'w', domain: 'docx' },
  { tag: 'tr', namespace: 'w', domain: 'docx' },
  { tag: 'tc', namespace: 'w', domain: 'docx' },
  { tag: 'cantSplit', namespace: 'w', domain: 'docx' },
  { tag: 'tblHeader', namespace: 'w', domain: 'docx' },
  { tag: 'worksheet', namespace: 'r', domain: 'xlsx' },
  { tag: 'sheetData', namespace: 'r', domain: 'xlsx' },
  { tag: 'row', namespace: 'r', domain: 'xlsx' },
  { tag: 'c', namespace: 'r', domain: 'xlsx' },
  { tag: 'v', namespace: 'r', domain: 'xlsx' },
  { tag: 'f', namespace: 'r', domain: 'xlsx' },
  { tag: 'sst', namespace: 'r', domain: 'xlsx' },
  { tag: 'si', namespace: 'r', domain: 'xlsx' },
  { tag: 'presentation', namespace: 'p', domain: 'pptx' },
  { tag: 'sld', namespace: 'p', domain: 'pptx' },
  { tag: 'sldLayout', namespace: 'p', domain: 'pptx' },
  { tag: 'sldMaster', namespace: 'p', domain: 'pptx' },
  { tag: 'sp', namespace: 'p', domain: 'pptx' },
  { tag: 'txBody', namespace: 'p', domain: 'pptx' },
  { tag: 'Relationships', namespace: 'r', domain: 'shared' },
  { tag: 'Relationship', namespace: 'r', domain: 'shared' },
  { tag: 'Types', namespace: 'r', domain: 'shared' },
  { tag: 'Override', namespace: 'r', domain: 'shared' },
  { tag: 'Default', namespace: 'r', domain: 'shared' }
];

const main = async () => {
  console.log(`[Ingest] Starting Mastra RAG Ingestion Workflow for ${REFERENCE_TAGS.length} calibration tags...`);

  try {
    // Explicitly initialize storage to ensure SQLite tables exist in script context
    const storage = mastra.getStorage();
    if (storage) {
      console.log('[Ingest] Initializing persistent storage...');
      await storage.init();
    }

    const workflow = mastra.getWorkflow('ooxml-rag-ingestion');
    const run = await workflow.createRun();
    const response = await run.start({ inputData: { tags: REFERENCE_TAGS } });
    
    if (response.status === 'success') {
      const result = response.result as { success: boolean; count: number };
      console.log(`[Ingest] Ingestion and compilation completed successfully!`);
      console.log(`  - Success: ${result.success}`);
      console.log(`  - Processed Tags Count: ${result.count}`);
    } else if (response.status === 'failed') {
      console.error(`[Ingest] Ingestion workflow failed:`, response.error);
    } else {
      console.error(`[Ingest] Ingestion workflow ended with status: ${response.status}`, response);
    }
  } catch (e) {
    console.error(`[Ingest] Ingestion workflow failed:`, e);
  }
};

main().catch(console.error);
