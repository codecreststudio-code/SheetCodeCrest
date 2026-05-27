export type OrderRow = Record<string, any>;

export function num(v: any): number {
  if (typeof v === "number") return Number.isFinite(v) ? v : 0;
  const n = parseFloat(String(v ?? "").replace(/[^0-9.-]/g, ""));
  return Number.isFinite(n) ? n : 0;
}

function normalizeKey(key: string): string {
  return key.replace(/[^a-z0-9]/gi, "").toLowerCase();
}

function findField(row: OrderRow, patterns: RegExp[], fallback?: string): string | undefined {
  const keys = Object.keys(row || {});
  return keys.find((key) => patterns.some((pattern) => pattern.test(normalizeKey(key)))) || fallback;
}

function text(row: OrderRow, field: string | undefined, fallback = "Unknown"): string {
  const value = field ? row[field] : undefined;
  const clean = String(value ?? "").trim();
  return clean || fallback;
}

function normalizeStatus(value: string): string {
  const raw = value.trim();
  const key = raw.replace(/[_-]+/g, " ").replace(/\s+/g, " ").toUpperCase();
  if (/^DELIVERED$/.test(key)) return "DELIVERED";
  if (key.includes("RTO") && key.includes("DELIVER")) return "RTO DELIVERED";
  if (key.includes("CANCEL")) return "CANCELED";
  if (key.includes("LOST")) return "LOST";
  if (key.includes("PICKUP")) return "PICKUP";
  if (key.includes("TRANSIT") || key.includes("SHIPPED")) return "IN TRANSIT";
  return key || "UNKNOWN";
}

function normalizePayment(value: string): string {
  const key = value.trim().toLowerCase();
  if (key.includes("cod") || key.includes("cash")) return "cod";
  if (key.includes("prepaid") || key.includes("online") || key.includes("upi") || key.includes("paid")) return "prepaid";
  return key || "unknown";
}

export function mergeMultiSKU(data: OrderRow[]): OrderRow[] {
  const map = new Map<string, OrderRow>();
  data.forEach((row) => {
    const keys = Object.keys(row || {});
    const normalize = (key: string) => key.replace(/\s+/g, "").toLowerCase();
    const findKey = (pred: (k: string) => boolean) => keys.find((k) => !!k && pred(normalize(k))) as string | undefined;
    const idKey = findKey((k) => k.includes("orderid") || (k.includes("order") && k.includes("id")))
      || findKey((k) => k.includes("order") && (k.includes("no") || k.includes("number")))
      || findKey((k) => k.includes("awb") || k.includes("shipment"))
      || keys.find((k) => !!k && /\d/.test(String(row[k] || "")))
      || "Order ID";
    const prodKey = findKey((k) => k.includes("product") && k.includes("name")) || findKey((k) => k.includes("product")) || "Product Name";
    const skuKey = findKey((k) => k.includes("sku")) || findKey((k) => k.includes("product") && k.includes("code")) || "Channel SKU";
    const qtyKey = findKey((k) => k.includes("quantity") || k.includes("qty")) || findKey((k) => k.includes("pcs") || k.includes("count")) || "Product Quantity";

    const id = String(row[idKey] || "").trim();
    if (!id) return;
    if (!map.has(id)) {
      map.set(id, { ...row, "SKU Count": 1, "Is Multi-SKU": "No", _products: [row[prodKey] || ""], _skus: [row[skuKey] || ""] });
    } else {
      const ex = map.get(id)!;
      ex["SKU Count"] = (ex["SKU Count"] || 1) + 1;
      ex["Is Multi-SKU"] = "Yes";
      const pName = row[prodKey] || "";
      if (pName && !ex._products.includes(pName)) ex._products.push(pName);
      const sku = row[skuKey] || "";
      if (sku && !ex._skus.includes(sku)) ex._skus.push(sku);
      ex[qtyKey] = (parseFloat(ex[qtyKey]) || 0) + (parseFloat(row[qtyKey]) || 0);
    }
  });
  const merged: OrderRow[] = [];
  map.forEach((row) => {
    row["Product Name"] = (row._products || []).join(" + ");
    row["Channel SKU"] = (row._skus || []).join(" + ");
    delete row._products; delete row._skus;
    merged.push(row);
  });
  return merged;
}

export type AnalyticsResult = {
  total: number;
  totalRev: number;
  totalQty: number;
  totalCOD: number;
  totalFreight: number;
  delivered: number;
  rto: number;
  deliveryRate: number;
  rtoRate: number;
  statusCounts: Record<string, any>;
  pickupCounts: Record<string, any>;
  qtyCounts: Record<string, any>;
  courierCounts: Record<string, any>;
  zoneCounts: Record<string, any>;
  stateCounts: Record<string, any>;
  ndrCounts: Record<string, number>;
  payCounts: Record<string, any>;
};

export function computeAnalytics(data: OrderRow[]): AnalyticsResult {
  const total = data.length;
  const statusCounts: Record<string, any> = {}, pickupCounts: Record<string, any> = {}, qtyCounts: Record<string, any> = {}, courierCounts: Record<string, any> = {}, zoneCounts: Record<string, any> = {}, stateCounts: Record<string, any> = {}, ndrCounts: Record<string, number> = {}, payCounts: Record<string, any> = {};
  let totalRev = 0, totalQty = 0, totalCOD = 0, totalFreight = 0;

  data.forEach((r) => {
    const statusField = findField(r, [/^status$/, /shipmentstatus/, /orderstatus/], "Status");
    const pickupField = findField(r, [/pickup.*address.*name/, /pickupaddress/, /warehouse/, /pickup/], "Pickup Address Name");
    const qtyField = findField(r, [/product.*qty/, /product.*quantity/, /^qty$/, /^quantity$/], "Product Quantity");
    const courierField = findField(r, [/courier.*company/, /^courier$/, /carrier/, /logisticspartner/], "Courier Company");
    const zoneField = findField(r, [/^zone$/, /shippingzone/], "Zone");
    const stateField = findField(r, [/address.*state/, /^state$/, /customer.*state/], "Address State");
    const ndrField = findField(r, [/ndr.*reason/, /non.*delivery.*reason/], "Latest NDR Reason");
    const payField = findField(r, [/payment.*method/, /paymentmode/, /^payment$/], "Payment Method");
    const revenueField = findField(r, [/order.*total/, /ordertotal/, /totalamount/, /invoiceamount/, /revenue/], "Order Total");
    const codField = findField(r, [/cod.*pay/, /codamount/, /collectableamount/, /collectibleamount/], "COD Payble Amount");
    const freightField = findField(r, [/freight.*total/, /shippingcharge/, /freight/, /couriercharge/], "Freight Total Amount");

    const status = normalizeStatus(text(r, statusField, "UNKNOWN"));
    const pickup = text(r, pickupField);
    const qty = num(qtyField ? r[qtyField] : 0);
    const courier = text(r, courierField);
    const zone = text(r, zoneField);
    const state = text(r, stateField);
    const ndr = text(r, ndrField, "");
    const pay = normalizePayment(text(r, payField, "unknown"));
    const rev = num(revenueField ? r[revenueField] : 0);
    const cod = num(codField ? r[codField] : 0);
    const freight = num(freightField ? r[freightField] : 0);

    totalRev += rev; totalQty += qty; totalCOD += cod; totalFreight += freight;

    [statusCounts, pickupCounts, courierCounts, zoneCounts, stateCounts, payCounts].forEach((map, i) => {
      const key = [status, pickup, courier, zone, state, pay][i];
      if (!map[key]) map[key] = { orders: 0, qty: 0, revenue: 0, delivered: 0, rto: 0, canceled: 0, cod: 0, freight: 0, prepaid: 0, codOrders: 0 };
      map[key].orders++;
      map[key].qty += qty;
      map[key].revenue += rev;
      map[key].cod += cod;
      map[key].freight += freight;
      if (status === "DELIVERED") map[key].delivered++;
      if (status.includes("RTO")) map[key].rto++;
      if (status === "CANCELED") map[key].canceled++;
      if (pay === "prepaid") map[key].prepaid++;
      if (pay === "cod") map[key].codOrders++;
    });

    if (!qtyCounts[qty]) qtyCounts[qty] = { orders: 0, revenue: 0, delivered: 0, rto: 0, cod: 0 };
    qtyCounts[qty].orders++; qtyCounts[qty].revenue += rev; qtyCounts[qty].cod += cod;
    if (status === "DELIVERED") qtyCounts[qty].delivered++;
    if (status.includes("RTO")) qtyCounts[qty].rto++;

    if (ndr) { ndrCounts[ndr] = (ndrCounts[ndr] || 0) + 1; }
  });

  const delivered = statusCounts["DELIVERED"]?.orders || 0;
  let rto = 0;
  Object.keys(statusCounts).forEach((statusName) => {
    if (statusName.includes("RTO")) {
      rto += statusCounts[statusName].orders;
    }
  });

  return { total, totalRev, totalQty, totalCOD, totalFreight, delivered, rto, deliveryRate: total ? delivered / total : 0, rtoRate: total ? rto / total : 0, statusCounts, pickupCounts, qtyCounts, courierCounts, zoneCounts, stateCounts, ndrCounts, payCounts };
}
