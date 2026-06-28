import JSZip, { JSZipObject } from 'jszip';
import { FileNode, DiffNode } from '../types';

interface ExtendedZipObject extends JSZipObject {
    crc32?: number;
    _data?: {
        crc32?: number;
        compressionMethod?: number;
        compression?: { magic?: Uint8Array };
    };
}

export const loadZipFile = async (file: File): Promise<{ zip: JSZip; tree: FileNode; flat: Record<string, JSZipObject> }> => {
  const zip = await new JSZip().loadAsync(file);
  
  if (!zip.file("[Content_Types].xml")) {
    throw new Error("Invalid OOXML file: Missing [Content_Types].xml");
  }

  const root: FileNode = { name: 'root', path: '', isFolder: true, children: {} };
  const flat: Record<string, JSZipObject> = {};

  zip.forEach((relativePath, zipEntry) => {
    // Sanitize and neutralize path traversal attacks
    if (relativePath.includes('../') || relativePath.includes('..\\') || relativePath.includes('./') || relativePath.includes('.\\')) {
        console.warn(`[Security] Ignored zip entry with path traversal sequence: ${relativePath}`);
        return;
    }
    
    flat[relativePath] = zipEntry;
    const parts = relativePath.split('/');
    let currentLevel = root;
    
    parts.forEach((part, index) => {
      const isLast = index === parts.length - 1;
      if (part === "" && isLast) return;

      if (!currentLevel.children[part]) {
        const isEntryDir = zipEntry.dir;
        const isFolder = !isLast || isEntryDir;

        currentLevel.children[part] = {
          name: part,
          path: isFolder ? (currentLevel.path ? `${currentLevel.path}/${part}` : part) : relativePath,
          isFolder: isFolder,
          children: {},
          zipEntry: isFolder ? undefined : zipEntry
        };
      }
      currentLevel = currentLevel.children[part];
    });
  });

  return { zip, tree: root, flat };
};

export const generateDiffTree = (
  flatA: Record<string, JSZipObject>, 
  flatB: Record<string, JSZipObject>
): DiffNode => {
  const root: DiffNode = { name: 'root', path: '', isFolder: true, children: {}, hasChange: false };
  const allPaths = new Set([...Object.keys(flatA), ...Object.keys(flatB)]);
  const sortedPaths = Array.from(allPaths).sort();

  sortedPaths.forEach(path => {
    const entryA = flatA[path] as ExtendedZipObject | undefined;
    const entryB = flatB[path] as ExtendedZipObject | undefined;

    let status: 'added' | 'deleted' | 'modified' | 'unchanged' = 'unchanged';
    
    if (entryA && !entryB) {
      if (!entryA.dir) status = 'deleted';
    } else if (!entryA && entryB) {
      if (!entryB.dir) status = 'added';
    } else if (entryA && entryB && !entryA.dir && !entryB.dir) {
      // Robust CRC check for JSZip internal structures
      const getCrc = (entry: ExtendedZipObject) => {
          if (entry.crc32 !== undefined) return entry.crc32;
          if (entry._data && entry._data.crc32 !== undefined) return entry._data.crc32;
          return null;
      };

      const crcA = getCrc(entryA);
      const crcB = getCrc(entryB);
      
      if (crcA !== null && crcB !== null && crcA !== crcB) {
        status = 'modified';
      }
    }

    if ( (entryA && entryA.dir) || (entryB && entryB.dir) ) {
        return; 
    }

    const parts = path.split('/');
    let currentLevel = root;

    parts.forEach((part, index) => {
      if (!currentLevel.children[part]) {
        const isLast = index === parts.length - 1;
        currentLevel.children[part] = {
          name: part,
          path: isLast ? path : (currentLevel.path ? `${currentLevel.path}/${part}` : part),
          isFolder: !isLast,
          children: {},
          status: isLast ? status : undefined,
          hasChange: false
        } as DiffNode;
      }
      currentLevel = currentLevel.children[part] as DiffNode;
    });
  });

  const propagateChange = (node: DiffNode): boolean => {
    if (!node.isFolder) {
      return node.status !== 'unchanged';
    }
    let childChanged = false;
    Object.values(node.children).forEach(child => {
      if (propagateChange(child as DiffNode)) {
        childChanged = true;
      }
    });
    node.hasChange = childChanged;
    return childChanged;
  };

  propagateChange(root);
  return root;
};

// Enhanced detection of compression from JSZip internal structures
const detectCompression = (fileEntry: ExtendedZipObject): 'STORE' | 'DEFLATE' | null => {
    if (!fileEntry || !fileEntry._data) return null;
    
    // Check compressionMethod property (0 = STORE, 8 = DEFLATE)
    if (typeof fileEntry._data.compressionMethod === 'number') {
         if (fileEntry._data.compressionMethod === 0) return 'STORE';
         if (fileEntry._data.compressionMethod === 8) return 'DEFLATE';
    }
    
    // Check magic bytes in compression object
    if (fileEntry._data.compression && fileEntry._data.compression.magic) {
        const magic = fileEntry._data.compression.magic; // Uint8Array
        if (magic[0] === 0 && magic[1] === 0) return 'STORE';
        if (magic[0] === 8 && magic[1] === 0) return 'DEFLATE';
    }
    
    return null;
};

/**
 * Creates a new ZIP Blob with pending changes applied.
 * Ensures strict OOXML compliance (mimetype first, uncompressed).
 */
export const createModifiedZip = async (originalZip: JSZip, pendingChanges: Record<string, string>): Promise<Blob> => {
    console.log("📦 Preparing Export (Re-packing mode)...");
    
    // Create a NEW zip to ensure clean structure and control order
    const newZip = new JSZip();

    // 1. Handle mimetype FIRST (Strict OOXML requirement)
    // It must be uncompressed (STORE) and the first entry.
    let mimeContent: string | Uint8Array | null = null;
    
    // Priority: Pending Change > Original File
    if (pendingChanges["mimetype"]) {
        mimeContent = pendingChanges["mimetype"];
    } else if (originalZip.file("mimetype")) {
        mimeContent = await originalZip.file("mimetype")!.async("uint8array");
    }

    if (mimeContent) {
        // Force STORE for mimetype
        newZip.file("mimetype", mimeContent, { compression: "STORE" });
    }

    // 2. Transfer other files
    // Merge paths to handle both existing and potentially new files
    const allPaths = new Set([...Object.keys(originalZip.files), ...Object.keys(pendingChanges)]);
    
    for (const path of allPaths) {
        if (path === "mimetype") continue; // Already handled
        
        // Handle Directories
        if (originalZip.files[path] && originalZip.files[path].dir) {
            newZip.folder(path);
            continue;
        }

        const hasPendingChange = Object.hasOwn(pendingChanges, path);
        
        if (hasPendingChange) {
            // Write NEW content
            // Default to DEFLATE for modified content as it's standard for XML
            newZip.file(path, pendingChanges[path], { compression: "DEFLATE" });
        } else if (originalZip.files[path]) {
            // Transfer UNMODIFIED content
            const originalEntry = originalZip.files[path];
            
            // Read raw content
            const content = await originalEntry.async("uint8array");
            
            // Detect and Preserve Compression
            const detected = detectCompression(originalEntry);
            const compressionMode = (detected === 'STORE') ? 'STORE' : 'DEFLATE';
            
            newZip.file(path, content, { compression: compressionMode });
        }
    }

    // 3. Generate
    console.log("📦 Generating OOXML Package...");
    const blob = await newZip.generateAsync({ 
        type: 'blob',
        compression: "DEFLATE",
        compressionOptions: {
            level: 6 // Standard Office compression
        },
        mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    });
    
    console.log(`📊 Final Size: ${(blob.size / 1024).toFixed(2)} KB`);
    return blob;
};

/**
 * UI Helper to download the modified zip.
 */
export const exportModifiedZip = async (originalZip: JSZip, pendingChanges: Record<string, string>, fileName: string) => {
    const blob = await createModifiedZip(originalZip, pendingChanges);
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `MODIFIED_${fileName}`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
};