export type OrderRow = Record<string, any>;

export function num(v: any): number {
  const n = parseFloat(v);
  return Number.isFinite(n) ? n : 0;
}

export function mergeMultiSKU(data: OrderRow[]): OrderRow[] {
  const map = new Map<string, OrderRow>();
  data.forEach((row) => {
    const keys = Object.keys(row || {});
    const findKey = (pred: (k: string) => boolean) => keys.find((k) => !!k && pred(k)) as string | undefined;
    const idKey = findKey((k) => k.replace(/\s+/g, "").toLowerCase().includes("orderid")) || findKey((k) => /\border\b/i.test(k) && /\bid\b/i.test(k)) || "Order ID";
    const prodKey = findKey((k) => /product.*name/i.test(k)) || "Product Name";
    const skuKey = findKey((k) => /sku/i.test(k)) || "Channel SKU";
    const qtyKey = findKey((k) => /quantity|qty/i.test(k)) || "Product Quantity";

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
    const status = String(r["Status"] || "UNKNOWN").trim();
    const pickup = String(r["Pickup Address Name"] || "Unknown").trim();
    const qty = num(r["Product Quantity"]);
    const courier = String(r["Courier Company"] || "Unknown").trim();
    const zone = String(r["Zone"] || "Unknown").trim();
    const state = String(r["Address State"] || "Unknown").trim();
    const ndr = String(r["Latest NDR Reason"] || "").trim();
    const pay = String(r["Payment Method"] || "unknown").trim();
    const rev = num(r["Order Total"]);
    const cod = num(r["COD Payble Amount"]);
    const freight = num(r["Freight Total Amount"]);

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
      if (status === "RTO DELIVERED") map[key].rto++;
      if (status === "CANCELED") map[key].canceled++;
      if (pay === "prepaid") map[key].prepaid++;
      if (pay === "cod") map[key].codOrders++;
    });

    if (!qtyCounts[qty]) qtyCounts[qty] = { orders: 0, revenue: 0, delivered: 0, rto: 0, cod: 0 };
    qtyCounts[qty].orders++; qtyCounts[qty].revenue += rev; qtyCounts[qty].cod += cod;
    if (status === "DELIVERED") qtyCounts[qty].delivered++;
    if (status === "RTO DELIVERED") qtyCounts[qty].rto++;

    if (ndr) { ndrCounts[ndr] = (ndrCounts[ndr] || 0) + 1; }
  });

  const delivered = statusCounts["DELIVERED"]?.orders || 0;
  const rto = statusCounts["RTO DELIVERED"]?.orders || 0;

  return { total, totalRev, totalQty, totalCOD, totalFreight, delivered, rto, deliveryRate: total ? delivered / total : 0, rtoRate: total ? rto / total : 0, statusCounts, pickupCounts, qtyCounts, courierCounts, zoneCounts, stateCounts, ndrCounts, payCounts };
}
