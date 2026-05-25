export interface User {
  username: string;
  passwordHash: string;
  isPro: boolean;
  dateCreated: string;
  name?: string;
  mobile?: string;
  email?: string;
}

export interface SavedRecord {
  id?: number;
  username: string;
  filename: string;
  size: number;
  mode: "universal" | "logistics" | "shopify";
  rawRows: any[];
  tableHeaders: string[];
  dataProfile: any | null;
  logisticsAnalytics: any | null;
  shopifyAnalytics: any | null;
  outName: string;
  chatHistory: any[];
  timestamp: string;
}

const DB_NAME = "auto_excel_intel_db";
const DB_VERSION = 1;

export function openDB(): Promise<IDBDatabase> {
  return new Promise((resolve, reject) => {
    const request = indexedDB.open(DB_NAME, DB_VERSION);
    request.onupgradeneeded = (event: any) => {
      const db = event.target.result;
      if (!db.objectStoreNames.contains("users")) {
        db.createObjectStore("users", { keyPath: "username" });
      }
      if (!db.objectStoreNames.contains("records")) {
        const recordsStore = db.createObjectStore("records", { keyPath: "id", autoIncrement: true });
        recordsStore.createIndex("username", "username", { unique: false });
      }
    };
    request.onsuccess = (event: any) => resolve(event.target.result);
    request.onerror = (event: any) => reject(event.target.error);
  });
}

export async function dbSaveUser(user: User): Promise<void> {
  const db = await openDB();
  return new Promise((resolve, reject) => {
    const transaction = db.transaction("users", "readwrite");
    const store = transaction.objectStore("users");
    const request = store.put(user);
    request.onsuccess = () => resolve();
    request.onerror = () => reject(request.error);
  });
}

export async function dbGetUser(username: string): Promise<User | null> {
  const db = await openDB();
  return new Promise((resolve, reject) => {
    const transaction = db.transaction("users", "readonly");
    const store = transaction.objectStore("users");
    const request = store.get(username);
    request.onsuccess = () => resolve(request.result || null);
    request.onerror = () => reject(request.error);
  });
}

export async function dbSaveRecord(record: SavedRecord): Promise<number> {
  const db = await openDB();
  return new Promise((resolve, reject) => {
    const transaction = db.transaction("records", "readwrite");
    const store = transaction.objectStore("records");
    const request = store.put(record);
    request.onsuccess = (event: any) => resolve(event.target.result);
    request.onerror = () => reject(request.error);
  });
}

export async function dbGetRecords(username: string): Promise<SavedRecord[]> {
  const db = await openDB();
  return new Promise((resolve, reject) => {
    const transaction = db.transaction("records", "readonly");
    const store = transaction.objectStore("records");
    const index = store.index("username");
    const request = index.getAll(username);
    request.onsuccess = () => resolve(request.result || []);
    request.onerror = () => reject(request.error);
  });
}

export async function dbDeleteRecord(id: number): Promise<void> {
  const db = await openDB();
  return new Promise((resolve, reject) => {
    const transaction = db.transaction("records", "readwrite");
    const store = transaction.objectStore("records");
    const request = store.delete(id);
    request.onsuccess = () => resolve();
    request.onerror = () => reject(request.error);
  });
}
