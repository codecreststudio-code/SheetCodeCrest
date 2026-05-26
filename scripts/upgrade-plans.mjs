// upgrade-plans-v2.mjs - CRLF-aware Node.js plan upgrade script
import { readFileSync, writeFileSync } from 'fs';
import { resolve } from 'path';

const file = resolve('src/App.tsx');
let content = readFileSync(file, 'utf-8');
const origLen = content.length;
console.log(`Original: ${origLen} chars`);

// Helper: convert \n to \r\n for matching Windows files
const crlf = s => s.replace(/\r\n/g, '\n').replace(/\n/g, '\r\n');

let changed = 0;
function replace(old, neo, label) {
  const oldCrlf = crlf(old);
  if (content.includes(oldCrlf)) {
    content = content.replace(oldCrlf, crlf(neo));
    console.log(`✓ ${label}`);
    changed++;
  } else {
    // Try without CRLF normalization
    if (content.includes(old)) {
      content = content.replace(old, neo);
      console.log(`✓ (LF) ${label}`);
      changed++;
    } else {
      console.log(`✗ NOT FOUND: ${label}`);
    }
  }
}

// ── 1. openNewPlan: add extended field resets ──
replace(
`  const openNewPlan = () => {
    setAdminEditPlan(null);
    setAdminPlanName("");
    setAdminPlanPrice(0);
    setAdminPlanPeriod("monthly");
    setAdminPlanFeatures([]);
    setAdminPlanFeatureInput("");
    setAdminPlanActive(true);
    setAdminPlanModalOpen(true);
  };`,
`  const openNewPlan = () => {
    setAdminEditPlan(null);
    setAdminPlanName("");
    setAdminPlanPrice(0);
    setAdminPlanPeriod("monthly");
    setAdminPlanFeatures([]);
    setAdminPlanFeatureInput("");
    setAdminPlanActive(true);
    setAdminPlanDescription("");
    setAdminPlanHighlighted(false);
    setAdminPlanColor("#f59e0b");
    setAdminPlanMaxReports(0);
    setAdminPlanSortOrder(adminPlans.length);
    setAdminPlanModalOpen(true);
  };`,
'openNewPlan'
);

// ── 2. openEditPlan: populate extended fields ──
replace(
`  const openEditPlan = (plan: Plan) => {
    setAdminEditPlan(plan);
    setAdminPlanName(plan.name);
    setAdminPlanPrice(plan.price);
    setAdminPlanPeriod(plan.billingPeriod);
    setAdminPlanFeatures([...plan.features]);
    setAdminPlanFeatureInput("");
    setAdminPlanActive(plan.isActive);
    setAdminPlanModalOpen(true);
  };`,
`  const openEditPlan = (plan: Plan) => {
    setAdminEditPlan(plan);
    setAdminPlanName(plan.name);
    setAdminPlanPrice(plan.price);
    setAdminPlanPeriod(plan.billingPeriod);
    setAdminPlanFeatures([...plan.features]);
    setAdminPlanFeatureInput("");
    setAdminPlanActive(plan.isActive);
    setAdminPlanDescription(plan.description || "");
    setAdminPlanHighlighted(plan.highlighted || false);
    setAdminPlanColor(plan.color || "#f59e0b");
    setAdminPlanMaxReports(plan.maxReports ?? 0);
    setAdminPlanSortOrder(plan.sortOrder ?? 99);
    setAdminPlanModalOpen(true);
  };`,
'openEditPlan'
);

// ── 3. handleSavePlan: include extended fields ──
replace(
`      const plan: Plan = {
        id: adminEditPlan?.id,
        name: adminPlanName.trim(),
        price: adminPlanPrice,
        billingPeriod: adminPlanPeriod,
        features: adminPlanFeatures,
        isActive: adminPlanActive
      };`,
`      const plan: Plan = {
        id: adminEditPlan?.id,
        name: adminPlanName.trim(),
        price: adminPlanPrice,
        billingPeriod: adminPlanPeriod,
        features: adminPlanFeatures,
        isActive: adminPlanActive,
        description: adminPlanDescription.trim(),
        highlighted: adminPlanHighlighted,
        color: adminPlanColor,
        maxReports: adminPlanMaxReports,
        sortOrder: adminPlanSortOrder,
      };`,
'handleSavePlan plan object'
);

// ── 4. Replace entire Plans Tab UI ──
// Use a marker search that accounts for CRLF
const plansTabStart = `{/* \u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500 TAB: PLANS \u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500 */}`;
const plansTabEnd   = `{/* \u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500 TAB: ANALYTICS \u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500 */}`;
const si = content.indexOf(plansTabStart);
const ei = content.indexOf(plansTabEnd);
console.log(`Plans tab markers: si=${si}, ei=${ei}`);

if (si >= 0 && ei > si) {
  const indent = '                  ';
  const newPlansTab = crlf(`{/* \u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500 TAB: PLANS \u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500 */}
${indent}{adminTab === "plans" && (
${indent}  <div>
${indent}    {/* Header */}
${indent}    <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: "1.25rem", flexWrap: "wrap", gap: "8px" }}>
${indent}      <div>
${indent}        <div style={{ fontWeight: 700, fontSize: "14px" }}>📦 Plan Packages</div>
${indent}        <div style={{ fontSize: "12px", color: "var(--slate)", marginTop: "2px" }}>Configure subscription tiers — changes appear instantly in checkout.</div>
${indent}      </div>
${indent}      <button type="button" onClick={openNewPlan}
${indent}        style={{ padding: "9px 18px", borderRadius: "10px", background: "var(--amber)", color: "#000", border: "none", fontWeight: 800, fontSize: "12px", cursor: "pointer" }}>
${indent}        ➕ New Plan
${indent}      </button>
${indent}    </div>

${indent}    {/* Plan Cards */}
${indent}    <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill, minmax(270px, 1fr))", gap: "16px" }}>
${indent}      {adminPlans.map(plan => {
${indent}        const ac = plan.color || "#f59e0b";
${indent}        return (
${indent}          <div key={plan.id || plan.name}
${indent}            style={{ padding: "1.5rem", borderRadius: "16px", border: \`2px solid \${plan.isActive ? ac : "var(--hairline)"}\`, background: \`linear-gradient(135deg, \${ac}0a, transparent 70%)\`, position: "relative" }}>
${indent}            {plan.highlighted && plan.isActive && (
${indent}              <div style={{ position: "absolute", top: "-12px", left: "50%", transform: "translateX(-50%)", background: ac, color: "#000", fontSize: "9px", fontWeight: 800, padding: "3px 12px", borderRadius: "20px", whiteSpace: "nowrap" }}>⭐ RECOMMENDED</div>
${indent}            )}
${indent}            <div style={{ display: "flex", justifyContent: "space-between", marginBottom: "10px" }}>
${indent}              <span style={{ fontSize: "10px", padding: "2px 8px", borderRadius: "20px", fontWeight: 700, background: plan.isActive ? "rgba(16,185,129,0.15)" : "rgba(100,116,139,0.15)", color: plan.isActive ? "#10b981" : "#64748b" }}>
${indent}                {plan.isActive ? "● ACTIVE" : "○ INACTIVE"}
${indent}              </span>
${indent}              <span style={{ fontSize: "10px", color: "var(--slate)", fontFamily: "monospace" }}>#{plan.sortOrder ?? "—"}</span>
${indent}            </div>
${indent}            <div style={{ fontWeight: 800, fontSize: "18px", color: ac }}>{plan.name}</div>
${indent}            {plan.description && <div style={{ fontSize: "11px", color: "var(--slate)", marginBottom: "4px", fontStyle: "italic" }}>{plan.description}</div>}
${indent}            <div style={{ fontSize: "26px", fontWeight: 900, marginTop: "6px", color: plan.price === 0 ? "#10b981" : "var(--text)" }}>
${indent}              {plan.price === 0 ? "Free" : \`₹\${plan.price.toLocaleString()}\`}
${indent}              {plan.price > 0 && <span style={{ fontSize: "12px", fontWeight: 400, color: "var(--slate)", marginLeft: "4px" }}>/{plan.billingPeriod}</span>}
${indent}            </div>
${indent}            <div style={{ fontSize: "11px", color: "var(--slate)", margin: "4px 0 10px" }}>
${indent}              {(plan.maxReports ?? 0) === 0 ? "♾️ Unlimited reports" : \`📊 \${plan.maxReports} reports/\${plan.billingPeriod}\`}
${indent}            </div>
${indent}            <div style={{ display: "flex", flexWrap: "wrap", gap: "5px", marginBottom: "14px" }}>
${indent}              {plan.features.map((f, i) => (
${indent}                <span key={i} style={{ fontSize: "10px", padding: "2px 8px", borderRadius: "20px", background: \`\${ac}18\`, color: ac, fontWeight: 600 }}>✓ {f}</span>
${indent}              ))}
${indent}            </div>
${indent}            <div style={{ display: "flex", gap: "6px" }}>
${indent}              <button type="button" onClick={() => openEditPlan(plan)}
${indent}                style={{ flex: 1, padding: "8px 0", borderRadius: "8px", fontSize: "12px", fontWeight: 700, cursor: "pointer", border: \`1px solid \${ac}\`, background: \`\${ac}18\`, color: ac }}>
${indent}                ✏️ Edit
${indent}              </button>
${indent}              <button type="button" title="Duplicate"
${indent}                onClick={() => {
${indent}                  setAdminEditPlan(null);
${indent}                  setAdminPlanName(plan.name + " (Copy)");
${indent}                  setAdminPlanPrice(plan.price);
${indent}                  setAdminPlanPeriod(plan.billingPeriod);
${indent}                  setAdminPlanFeatures([...plan.features]);
${indent}                  setAdminPlanFeatureInput("");
${indent}                  setAdminPlanActive(false);
${indent}                  setAdminPlanDescription(plan.description || "");
${indent}                  setAdminPlanHighlighted(false);
${indent}                  setAdminPlanColor(plan.color || "#f59e0b");
${indent}                  setAdminPlanMaxReports(plan.maxReports ?? 0);
${indent}                  setAdminPlanSortOrder((plan.sortOrder ?? 99) + 1);
${indent}                  setAdminPlanModalOpen(true);
${indent}                }}
${indent}                style={{ padding: "8px 10px", borderRadius: "8px", fontSize: "12px", fontWeight: 700, cursor: "pointer", border: "1px solid var(--hairline)", background: "rgba(255,255,255,0.04)", color: "var(--slate)" }}>
${indent}                📋
${indent}              </button>
${indent}              <button type="button" onClick={() => handleDeletePlan(plan)}
${indent}                style={{ padding: "8px 10px", borderRadius: "8px", fontSize: "12px", fontWeight: 700, cursor: "pointer", border: "1px solid #ef4444", background: "rgba(239,68,68,0.08)", color: "#ef4444" }}>
${indent}                🗑️
${indent}              </button>
${indent}            </div>
${indent}          </div>
${indent}        );
${indent}      })}
${indent}      {adminPlans.length === 0 && (
${indent}        <div style={{ gridColumn: "1/-1", padding: "3rem", textAlign: "center", color: "var(--slate)", fontSize: "13px", border: "2px dashed var(--hairline)", borderRadius: "16px" }}>
${indent}          <div style={{ fontSize: "3rem", marginBottom: "8px" }}>📦</div>
${indent}          No plans yet. Click <strong>➕ New Plan</strong> to create your first package.
${indent}        </div>
${indent}      )}
${indent}    </div>

${indent}    {/* Plan Editor Side Panel */}
${indent}    {adminPlanModalOpen && (
${indent}      <div style={{ position: "fixed", top: 0, right: 0, width: "440px", height: "100vh", background: "var(--glass-bg)", backdropFilter: "blur(24px)", borderLeft: \`2px solid \${adminPlanColor}\`, zIndex: 9999, display: "flex", flexDirection: "column" }}>
${indent}        <div style={{ padding: "1.25rem 1.5rem", borderBottom: "1px solid var(--hairline)", display: "flex", alignItems: "center", justifyContent: "space-between", background: \`\${adminPlanColor}14\`, flexShrink: 0 }}>
${indent}          <div>
${indent}            <div style={{ fontWeight: 800, fontSize: "15px", color: adminPlanColor }}>{adminEditPlan ? "✏️ Edit Plan" : "➕ Create Plan"}</div>
${indent}            <div style={{ fontSize: "11px", color: "var(--slate)", marginTop: "2px" }}>{adminEditPlan ? \`Editing: \${adminEditPlan.name}\` : "New package"}</div>
${indent}          </div>
${indent}          <button type="button" onClick={() => setAdminPlanModalOpen(false)} style={{ background: "none", border: "none", color: "var(--slate)", fontSize: "20px", cursor: "pointer" }}>✕</button>
${indent}        </div>
${indent}        <div style={{ padding: "1.25rem 1.5rem", display: "flex", flexDirection: "column", gap: "14px", overflowY: "auto", flex: 1 }}>
${indent}          <div className="form-group">
${indent}            <label className="form-label">Plan Name *</label>
${indent}            <input className="form-input" type="text" value={adminPlanName} onChange={e => setAdminPlanName(e.target.value)} placeholder="e.g. Starter, Pro, Enterprise..." />
${indent}          </div>
${indent}          <div className="form-group">
${indent}            <label className="form-label">Tagline / Description</label>
${indent}            <input className="form-input" type="text" value={adminPlanDescription} onChange={e => setAdminPlanDescription(e.target.value)} placeholder="e.g. Perfect for small teams" />
${indent}          </div>
${indent}          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: "10px" }}>
${indent}            <div className="form-group">
${indent}              <label className="form-label">Price (₹) — 0 = Free</label>
${indent}              <input className="form-input" type="number" min={0} value={adminPlanPrice} onChange={e => setAdminPlanPrice(Number(e.target.value))} />
${indent}            </div>
${indent}            <div className="form-group">
${indent}              <label className="form-label">Billing Period</label>
${indent}              <select className="form-input" value={adminPlanPeriod} onChange={e => setAdminPlanPeriod(e.target.value as Plan["billingPeriod"])} style={{ appearance: "none", cursor: "pointer" }}>
${indent}                {(["free","monthly","yearly","lifetime"] as const).map(p => (
${indent}                  <option key={p} value={p}>{p.charAt(0).toUpperCase() + p.slice(1)}</option>
${indent}                ))}
${indent}              </select>
${indent}            </div>
${indent}          </div>
${indent}          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: "10px" }}>
${indent}            <div className="form-group">
${indent}              <label className="form-label">Report Limit (0 = ∞)</label>
${indent}              <input className="form-input" type="number" min={0} value={adminPlanMaxReports} onChange={e => setAdminPlanMaxReports(Number(e.target.value))} />
${indent}            </div>
${indent}            <div className="form-group">
${indent}              <label className="form-label">Sort Order</label>
${indent}              <input className="form-input" type="number" min={0} value={adminPlanSortOrder} onChange={e => setAdminPlanSortOrder(Number(e.target.value))} />
${indent}            </div>
${indent}          </div>
${indent}          <div className="form-group">
${indent}            <label className="form-label">Accent Color</label>
${indent}            <div style={{ display: "flex", gap: "10px", alignItems: "center" }}>
${indent}              <input type="color" value={adminPlanColor} onChange={e => setAdminPlanColor(e.target.value)}
${indent}                style={{ width: "48px", height: "40px", padding: "2px", borderRadius: "8px", border: "1px solid var(--hairline)", cursor: "pointer" }} />
${indent}              <div style={{ flex: 1, height: "40px", borderRadius: "8px", border: \`2px solid \${adminPlanColor}\`, background: \`\${adminPlanColor}14\`, display: "flex", alignItems: "center", justifyContent: "center", fontSize: "13px", fontWeight: 800, color: adminPlanColor }}>
${indent}                {adminPlanName || "Preview"}
${indent}              </div>
${indent}            </div>
${indent}          </div>
${indent}          <div className="form-group">
${indent}            <label className="form-label">Features</label>
${indent}            <div style={{ display: "flex", gap: "6px" }}>
${indent}              <input className="form-input" type="text" value={adminPlanFeatureInput} onChange={e => setAdminPlanFeatureInput(e.target.value)}
${indent}                onKeyDown={e => { if (e.key === "Enter" && adminPlanFeatureInput.trim()) { setAdminPlanFeatures(prev => [...prev, adminPlanFeatureInput.trim()]); setAdminPlanFeatureInput(""); e.preventDefault(); }}}
${indent}                placeholder="Type feature, press Enter..." />
${indent}              <button type="button" onClick={() => { if (adminPlanFeatureInput.trim()) { setAdminPlanFeatures(prev => [...prev, adminPlanFeatureInput.trim()]); setAdminPlanFeatureInput(""); }}}
${indent}                style={{ padding: "0 14px", borderRadius: "8px", border: "none", background: "var(--amber)", color: "#000", fontWeight: 700, cursor: "pointer", fontSize: "18px" }}>+</button>
${indent}            </div>
${indent}            <div style={{ display: "flex", flexWrap: "wrap", gap: "6px", marginTop: "8px" }}>
${indent}              {adminPlanFeatures.map((f, i) => (
${indent}                <span key={i} style={{ fontSize: "11px", padding: "4px 10px", borderRadius: "20px", background: \`\${adminPlanColor}18\`, color: adminPlanColor, display: "flex", alignItems: "center", gap: "5px", fontWeight: 600 }}>
${indent}                  {f}
${indent}                  <button type="button" onClick={() => setAdminPlanFeatures(prev => prev.filter((_, j) => j !== i))}
${indent}                    style={{ background: "none", border: "none", color: "#ef4444", cursor: "pointer", padding: 0, fontSize: "12px", lineHeight: 1 }}>✕</button>
${indent}                </span>
${indent}              ))}
${indent}            </div>
${indent}          </div>
${indent}          <div style={{ display: "flex", gap: "10px" }}>
${indent}            <label style={{ flex: 1, display: "flex", alignItems: "center", gap: "8px", padding: "10px 12px", borderRadius: "8px", cursor: "pointer", border: \`1px solid \${adminPlanActive ? "#10b981" : "var(--hairline)"}\`, background: "rgba(255,255,255,0.03)" }}>
${indent}              <input type="checkbox" id="plan-active-chk" checked={adminPlanActive} onChange={e => setAdminPlanActive(e.target.checked)} style={{ width: "16px", height: "16px", accentColor: "#10b981" }} />
${indent}              <span style={{ fontSize: "12px", fontWeight: 700 }}>Active (visible)</span>
${indent}            </label>
${indent}            <label style={{ flex: 1, display: "flex", alignItems: "center", gap: "8px", padding: "10px 12px", borderRadius: "8px", cursor: "pointer", border: \`1px solid \${adminPlanHighlighted ? adminPlanColor : "var(--hairline)"}\`, background: "rgba(255,255,255,0.03)" }}>
${indent}              <input type="checkbox" id="plan-highlight-chk" checked={adminPlanHighlighted} onChange={e => setAdminPlanHighlighted(e.target.checked)} style={{ width: "16px", height: "16px", accentColor: adminPlanColor }} />
${indent}              <span style={{ fontSize: "12px", fontWeight: 700 }}>⭐ Recommended</span>
${indent}            </label>
${indent}          </div>
${indent}        </div>
${indent}        <div style={{ padding: "1rem 1.5rem", borderTop: "1px solid var(--hairline)", display: "flex", gap: "8px", flexShrink: 0 }}>
${indent}          <button type="button" onClick={() => setAdminPlanModalOpen(false)} style={{ flex: 1, padding: "11px", border: "1px solid var(--hairline)", borderRadius: "10px", background: "transparent", color: "var(--slate)", cursor: "pointer", fontWeight: 700 }}>Cancel</button>
${indent}          <button type="button" onClick={handleSavePlan} disabled={adminLoading}
${indent}            style={{ flex: 2, padding: "11px", border: "none", borderRadius: "10px", background: adminPlanColor, color: "#000", cursor: "pointer", fontWeight: 800, fontSize: "13px", opacity: adminLoading ? 0.6 : 1 }}>
${indent}            {adminLoading ? "Saving..." : (adminEditPlan ? "💾 Save Changes" : "✅ Create Plan")}
${indent}          </button>
${indent}        </div>
${indent}      </div>
${indent}    )}
${indent}  </div>
${indent})}

${indent}`);
  content = content.substring(0, si) + newPlansTab + content.substring(ei);
  console.log('✓ Replaced Plans Tab UI');
  changed++;
} else {
  console.log(`✗ Plans Tab not found`);
}

// ── 6. Update checkout modal payment-summary-box ──
// Find via partial match
const oldCheckoutSummary = `<div className="payment-summary-box">\r\n                <span>🚀 SheetCodeCrest Pro Lifetime</span>\r\n                <strong>?1,599</strong>\r\n              </div>`;
const newCheckoutSummary = `{/* Plan Selection */}
              {checkoutPlans.length > 0 && (
                <div style={{ marginBottom: "1.25rem" }}>
                  <div style={{ fontSize: "11px", fontWeight: 700, textTransform: "uppercase", letterSpacing: "0.06em", color: "var(--slate)", marginBottom: "10px" }}>Choose Your Plan</div>
                  <div style={{ display: "grid", gridTemplateColumns: checkoutPlans.length > 2 ? "repeat(3,1fr)" : checkoutPlans.length === 2 ? "repeat(2,1fr)" : "1fr", gap: "10px" }}>
                    {checkoutPlans.map(plan => {
                      const isSel = selectedPlanId === plan.id;
                      const ac = plan.color || "#f59e0b";
                      return (
                        <button key={plan.id || plan.name} type="button" onClick={() => setSelectedPlanId(plan.id || null)}
                          style={{ position: "relative", padding: "14px 10px", borderRadius: "12px", border: \`2px solid \${isSel ? ac : "var(--hairline)"}\`, background: isSel ? \`\${ac}14\` : "rgba(255,255,255,0.02)", cursor: "pointer", textAlign: "left", transition: "all 0.15s" }}>
                          {plan.highlighted && <div style={{ position: "absolute", top: "-9px", left: "50%", transform: "translateX(-50%)", background: ac, color: "#000", fontSize: "7px", fontWeight: 800, padding: "2px 8px", borderRadius: "20px", whiteSpace: "nowrap" }}>⭐ BEST</div>}
                          <div style={{ fontWeight: 800, fontSize: "13px", color: isSel ? ac : "var(--text)" }}>{plan.name}</div>
                          {plan.description && <div style={{ fontSize: "9px", color: "var(--slate)", marginTop: "1px" }}>{plan.description}</div>}
                          <div style={{ fontWeight: 900, fontSize: "18px", color: ac, marginTop: "4px" }}>₹{plan.price.toLocaleString()}<span style={{ fontSize: "9px", fontWeight: 400, color: "var(--slate)", marginLeft: "2px" }}>/{plan.billingPeriod}</span></div>
                          <div style={{ marginTop: "4px" }}>{plan.features.slice(0,2).map((f,i) => <div key={i} style={{ fontSize: "9px", color: "var(--slate)" }}>✓ {f}</div>)}</div>
                          {isSel && <div style={{ position: "absolute", top: "8px", right: "8px", width: "14px", height: "14px", borderRadius: "50%", background: ac, display: "flex", alignItems: "center", justifyContent: "center", fontSize: "8px", color: "#000", fontWeight: 900 }}>✓</div>}
                        </button>
                      );
                    })}
                  </div>
                </div>
              )}
              <div className="payment-summary-box">
                {(() => {
                  const sel = checkoutPlans.find(p => p.id === selectedPlanId) || checkoutPlans[0];
                  return sel
                    ? <><span>🚀 {sel.name} — {sel.billingPeriod}</span><strong>₹{sel.price.toLocaleString()}</strong></>
                    : <><span>🚀 SheetCodeCrest Pro</span><strong>₹1,599</strong></>;
                })()}
              </div>`;

if (content.includes(oldCheckoutSummary)) {
  content = content.replace(oldCheckoutSummary, crlf(newCheckoutSummary));
  console.log('✓ Updated checkout modal');
  changed++;
} else {
  // Search for the key text to help debug
  const idx = content.indexOf('payment-summary-box');
  if (idx >= 0) {
    console.log('Found payment-summary-box at:', idx);
    console.log('Context:', JSON.stringify(content.substring(idx, idx+150)));
  }
  console.log('✗ Checkout modal summary not matched');
}

// ── Write ──
import { writeFileSync } from 'fs';
writeFileSync(file, content, 'utf-8');
console.log(`\nDone! ${changed} changes applied. New size: ${content.length} chars`);
