import type { PresentationItem } from "./presentation-types";

const DATABASE_NAME = "apresenta-local-v1";
const DATABASE_VERSION = 1;
const DECK_KEY = "current-deck";

export interface StoredMedia {
  id: string;
  name: string;
  mimeType: string;
  size: number;
  blob: Blob;
  storedAt: number;
}

export interface StoredDeck {
  key: string;
  title: string;
  items: PresentationItem[];
  updatedAt: number;
}

let databasePromise: Promise<IDBDatabase> | undefined;

function openDatabase(): Promise<IDBDatabase> {
  if (databasePromise) return databasePromise;
  const promise = new Promise<IDBDatabase>((resolve, reject) => {
    const request = indexedDB.open(DATABASE_NAME, DATABASE_VERSION);
    request.onupgradeneeded = () => {
      const database = request.result;
      if (!database.objectStoreNames.contains("media")) {
        database.createObjectStore("media", { keyPath: "id" });
      }
      if (!database.objectStoreNames.contains("decks")) {
        database.createObjectStore("decks", { keyPath: "key" });
      }
    };
    request.onsuccess = () => {
      const database = request.result;
      database.onversionchange = () => database.close();
      resolve(database);
    };
    request.onerror = () => reject(request.error ?? new Error("Não foi possível abrir o armazenamento local."));
    request.onblocked = () => reject(new Error("Feche as outras abas do Apresenta e tente novamente."));
  }).catch((error) => {
    databasePromise = undefined;
    throw error;
  });
  databasePromise = promise;
  return promise;
}

export function createId(): string {
  return globalThis.crypto?.randomUUID?.() ?? `${Date.now()}-${Math.random().toString(36).slice(2)}`;
}

export async function readDeck(): Promise<StoredDeck | undefined> {
  const database = await openDatabase();
  return new Promise((resolve, reject) => {
    const request = database.transaction("decks", "readonly").objectStore("decks").get(DECK_KEY);
    request.onsuccess = () => resolve(request.result as StoredDeck | undefined);
    request.onerror = () => reject(request.error ?? new Error("Não foi possível abrir a apresentação local."));
  });
}

export async function writeDeck(title: string, items: PresentationItem[]): Promise<void> {
  const database = await openDatabase();
  return new Promise((resolve, reject) => {
    const transaction = database.transaction(["decks", "media"], "readwrite");
    transaction.objectStore("decks").put({
      key: DECK_KEY,
      title,
      items,
      updatedAt: Date.now(),
    } satisfies StoredDeck);
    const referencedAssets = new Set(items.flatMap((item) => item.assetId ? [item.assetId] : []));
    const mediaStore = transaction.objectStore("media");
    const storedKeys = mediaStore.getAllKeys();
    storedKeys.onsuccess = () => {
      for (const key of storedKeys.result) {
        if (!referencedAssets.has(String(key))) mediaStore.delete(key);
      }
    };
    storedKeys.onerror = () => transaction.abort();
    transaction.oncomplete = () => resolve();
    transaction.onerror = () => reject(transaction.error ?? new Error("Não foi possível salvar a sequência."));
    transaction.onabort = () => reject(transaction.error ?? new Error("O salvamento local foi cancelado."));
  });
}

export async function writeMedia(media: StoredMedia): Promise<void> {
  const database = await openDatabase();
  return new Promise((resolve, reject) => {
    const transaction = database.transaction("media", "readwrite");
    transaction.objectStore("media").put(media);
    transaction.oncomplete = () => resolve();
    transaction.onerror = () => reject(transaction.error ?? new Error("Não foi possível guardar o arquivo no navegador."));
    transaction.onabort = () => reject(transaction.error ?? new Error("O armazenamento do arquivo foi cancelado."));
  });
}

export async function readMedia(id: string): Promise<StoredMedia | undefined> {
  const database = await openDatabase();
  return new Promise((resolve, reject) => {
    const request = database.transaction("media", "readonly").objectStore("media").get(id);
    request.onsuccess = () => resolve(request.result as StoredMedia | undefined);
    request.onerror = () => reject(request.error ?? new Error("Não foi possível abrir este arquivo local."));
  });
}

export async function estimateLocalStorage(): Promise<{ used?: number; quota?: number }> {
  try {
    const estimate = await navigator.storage?.estimate();
    return { used: estimate?.usage, quota: estimate?.quota };
  } catch {
    return {};
  }
}
