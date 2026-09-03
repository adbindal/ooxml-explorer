import { ReferenceDoc } from './staticKnowledgeBase';

const DB_NAME = 'ooxml_explorer_db';
/**
 * v2 re-keys records on `domain:namespace:tag`.
 *
 * v1 keyed on `domain:tag`, which cannot distinguish elements that share a local
 * name across namespaces - and 103 tags in the corpus do. `<a:bottom>` (a DrawingML
 * border) and `<w:bottom>` (a paragraph border) are different elements with different
 * attributes and parents, and v1 could only store one of them per domain.
 */
const DB_VERSION = 2;
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
      // The key scheme changed in v2, so an existing store holds records under keys
      // that can no longer be derived. Drop and repopulate rather than migrate -
      // the data is a regenerable cache of /rag-data.json, not user content.
      if (db.objectStoreNames.contains(STORE_NAME)) {
        db.deleteObjectStore(STORE_NAME);
      }
      const store = db.createObjectStore(STORE_NAME, { keyPath: 'id' });
      store.createIndex('tag', 'tag', { unique: false });
      store.createIndex('domain', 'domain', { unique: false });
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
      // Namespace is part of the key: the same local name means different elements
      // in different namespaces (`a:bottom` vs `w:bottom` vs `x:bottom`).
      const record = {
        ...doc,
        id: `${doc.domain}:${doc.namespace}:${doc.tag}`
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
 * Selects the right record for a tag among same-named candidates.
 *
 * Exported for testing and reuse by the bundled offline fallback, which must apply
 * exactly the same rule so that an offline user and an online user cannot get
 * different answers for the same element.
 *
 * The namespace prefix is decisive when it matches. When it doesn't, this returns a
 * record only if the choice is unambiguous - because handing back a same-named
 * element from the wrong namespace would present a confidently wrong answer under a
 * "Grounded" badge, which is worse than admitting the tag isn't covered.
 */
/**
 * Does this record match a keyword? Pure, so it can be tested without IndexedDB.
 *
 * ⚠️ `definition` is OPTIONAL and usually absent — 1,870 of the 1,899 records are
 * generated from the SDK schema, which supplies structure and no prose. An earlier
 * version read `doc.definition.toLowerCase()` directly, which held only while the corpus
 * was the 29 hand-written entries. Once generated records arrived, the first one threw
 * and took the entire natural-language search down with it — in the browser, uncaught.
 *
 * The bug survived because the only coverage of this search was mocks that reimplemented
 * the predicate rather than calling it. Extracted here so a test can reach the real one.
 */
export const matchesKeyword = (doc: ReferenceDoc, cleanKeyword: string): boolean =>
  doc.tag.toLowerCase().includes(cleanKeyword) ||
  (doc.definition ?? '').toLowerCase().includes(cleanKeyword);

export const selectBestMatch = (
  candidates: ReferenceDoc[],
  domain: 'docx' | 'xlsx' | 'pptx' | 'shared',
  namespace?: string
): ReferenceDoc | null => {
  const inScope = candidates.filter(m => m.domain === domain || m.domain === 'shared');
  if (inScope.length === 0) return null;

  if (namespace) {
    const exact = inScope.filter(m => m.namespace === namespace);
    if (exact.length > 0) {
      // Prefer the record from the document's own domain over a `shared` one.
      return exact.find(m => m.domain === domain) ?? exact[0];
    }
    // The prefix did not match any candidate. A document may legitimately bind an
    // unconventional prefix, so fall through - but only when there is nothing to
    // choose between.
    return inScope.length === 1 ? inScope[0] : null;
  }

  return inScope.find(m => m.domain === domain) ?? inScope[0];
};

/**
 * Queries a single tag schema from IndexedDB by tag name, domain and namespace prefix.
 *
 * `namespace` is optional only so callers that genuinely do not know it (a keyword
 * search hit) can omit it. Callers that have it must pass it: 103 tags in the corpus
 * exist under more than one namespace.
 */
export const querySchemaFromStorage = async (
  tag: string,
  domain: 'docx' | 'xlsx' | 'pptx' | 'shared',
  namespace?: string
): Promise<ReferenceDoc | null> => {
  const db = await getDB();

  return new Promise((resolve, reject) => {
    const transaction = db.transaction(STORE_NAME, 'readonly');
    const store = transaction.objectStore(STORE_NAME);
    const index = store.index('tag');
    const request = index.getAll(tag);

    request.onsuccess = () => {
      const matches = request.result as (ReferenceDoc & { id: string })[];
      resolve(selectBestMatch(matches, domain, namespace));
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
          if (matchesKeyword(doc, cleanKeyword)) {
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
