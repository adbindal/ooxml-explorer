import { ReferenceDoc } from './staticKnowledgeBase';

const DB_NAME = 'ooxml_explorer_db';
const DB_VERSION = 1;
const STORE_NAME = 'rag_schemas';

let dbInstance: IDBDatabase | null = null;

/**
 * Initializes the IndexedDB database.
 */
const getDB = (): Promise<IDBDatabase> => {
  if (dbInstance) return Promise.resolve(dbInstance);

  return new Promise((resolve, reject) => {
    const request = indexedDB.open(DB_NAME, DB_VERSION);

    request.onupgradeneeded = () => {
      const db = request.result;
      if (!db.objectStoreNames.contains(STORE_NAME)) {
        const store = db.createObjectStore(STORE_NAME, { keyPath: 'id' });
        store.createIndex('tag', 'tag', { unique: false });
        store.createIndex('domain', 'domain', { unique: false });
      }
    };

    request.onsuccess = () => {
      dbInstance = request.result;
      resolve(request.result);
    };

    request.onerror = () => {
      reject(request.error);
    };
  });
};

/**
 * Feeds the initial RAG data from public/rag-data.json into IndexedDB.
 */
export const initStorageService = async (): Promise<void> => {
  const db = await getDB();
  
  // Check if store is already populated
  const count = await new Promise<number>((resolve, reject) => {
    const transaction = db.transaction(STORE_NAME, 'readonly');
    const store = transaction.objectStore(STORE_NAME);
    const request = store.count();
    request.onsuccess = () => resolve(request.result);
    request.onerror = () => reject(request.error);
  });

  if (count > 0) {
    return; // Already populated
  }

  console.log('[StorageService] IndexedDB empty. Initializing RAG schema database...');
  try {
    const response = await fetch('/rag-data.json');
    if (!response.ok) {
      throw new Error(`Failed to fetch schema data: ${response.statusText}`);
    }
    const data: ReferenceDoc[] = await response.json();

    const transaction = db.transaction(STORE_NAME, 'readwrite');
    const store = transaction.objectStore(STORE_NAME);

    for (const doc of data) {
      // Create a unique composite key for the store
      const record = {
        ...doc,
        id: `${doc.domain}:${doc.tag}`
      };
      store.put(record);
    }

    await new Promise<void>((resolve, reject) => {
      transaction.oncomplete = () => resolve();
      transaction.onerror = () => reject(transaction.error);
    });
    console.log(`[StorageService] Loaded ${data.length} schemas into IndexedDB.`);
  } catch (error) {
    console.error('[StorageService] Failed to initialize schema database:', error);
  }
};

/**
 * Queries a single tag schema from IndexedDB by tag name and domain.
 */
export const querySchemaFromStorage = async (
  tag: string,
  domain: 'docx' | 'xlsx' | 'pptx' | 'shared'
): Promise<ReferenceDoc | null> => {
  const db = await getDB();

  return new Promise((resolve, reject) => {
    const transaction = db.transaction(STORE_NAME, 'readonly');
    const store = transaction.objectStore(STORE_NAME);
    const index = store.index('tag');
    const request = index.getAll(tag);

    request.onsuccess = () => {
      const matches = request.result as (ReferenceDoc & { id: string })[];
      // Filter by domain or shared
      const bestMatch = matches.find(m => m.domain === domain || m.domain === 'shared');
      resolve(bestMatch || null);
    };

    request.onerror = () => {
      reject(request.error);
    };
  });
};

/**
 * Performs a keyword search on tag name and definition across IndexedDB.
 */
export const searchSchemasInStorage = async (
  keyword: string,
  domain: 'docx' | 'xlsx' | 'pptx' | 'shared'
): Promise<ReferenceDoc[]> => {
  const db = await getDB();
  const cleanKeyword = keyword.toLowerCase().trim();

  return new Promise((resolve, reject) => {
    const transaction = db.transaction(STORE_NAME, 'readonly');
    const store = transaction.objectStore(STORE_NAME);
    const request = store.openCursor();
    const results: ReferenceDoc[] = [];

    request.onsuccess = (event) => {
      const cursor = (event.target as IDBRequest<IDBCursorWithValue | null>).result;
      if (cursor) {
        const doc = cursor.value as ReferenceDoc;
        const inDomain = doc.domain === domain || doc.domain === 'shared';
        
        if (inDomain) {
          const matchesTag = doc.tag.toLowerCase().includes(cleanKeyword);
          const matchesDef = doc.definition.toLowerCase().includes(cleanKeyword);
          if (matchesTag || matchesDef) {
            results.push(doc);
          }
        }
        cursor.continue();
      } else {
        resolve(results);
      }
    };

    request.onerror = () => {
      reject(request.error);
    };
  });
};
