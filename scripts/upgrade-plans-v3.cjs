// upgrade-plans-v3.cjs
const fs = require('fs');
const path = require('path');

const file = path.resolve('src/App.tsx');
let c = fs.readFileSync(file, 'utf-8');
console.log('Original:', c.length, 'chars');

const crlf = s => s.replace(/\r\n/g, '\n').replace(/\n/g, '\r\n');
let changed = 0;

function rep(old, neo, label) {
  const oc = crlf(old);
  if (c.includes(oc)) {
    c = c.replace(oc, crlf(neo));
    console.log('v ' + label);
    changed++;
  } else if (c.includes(old)) {
    c = c.replace(old, neo);
    console.log('v(lf) ' + label);
    changed++;
  } else {
    console.log('x NOT FOUND: ' + label);
  }
}

// 1. openNewPlan
rep(
"  const openNewPlan = () => {\n" +
"    setAdminEditPlan(null);\n" +
"    setAdminPlanName(\"\");\n" +
"    setAdminPlanPrice(0);\n" +
"    setAdminPlanPeriod(\"monthly\");\n" +
"    setAdminPlanFeatures([]);\n" +
"    setAdminPlanFeatureInput(\"\");\n" +
"    setAdminPlanActive(true);\n" +
"    setAdminPlanModalOpen(true);\n" +
"  };",
"  const openNewPlan = () => {\n" +
"    setAdminEditPlan(null);\n" +
"    setAdminPlanName(\"\");\n" +
"    setAdminPlanPrice(0);\n" +
"    setAdminPlanPeriod(\"monthly\");\n" +
"    setAdminPlanFeatures([]);\n" +
"    setAdminPlanFeatureInput(\"\");\n" +
"    setAdminPlanActive(true);\n" +
"    setAdminPlanDescription(\"\");\n" +
"    setAdminPlanHighlighted(false);\n" +
"    setAdminPlanColor(\"#f59e0b\");\n" +
"    setAdminPlanMaxReports(0);\n" +
"    setAdminPlanSortOrder(adminPlans.length);\n" +
"    setAdminPlanModalOpen(true);\n" +
"  };",
"openNewPlan"
);

// 2. openEditPlan
rep(
"  const openEditPlan = (plan: Plan) => {\n" +
"    setAdminEditPlan(plan);\n" +
"    setAdminPlanName(plan.name);\n" +
"    setAdminPlanPrice(plan.price);\n" +
"    setAdminPlanPeriod(plan.billingPeriod);\n" +
"    setAdminPlanFeatures([...plan.features]);\n" +
"    setAdminPlanFeatureInput(\"\");\n" +
"    setAdminPlanActive(plan.isActive);\n" +
"    setAdminPlanModalOpen(true);\n" +
"  };",
"  const openEditPlan = (plan: Plan) => {\n" +
"    setAdminEditPlan(plan);\n" +
"    setAdminPlanName(plan.name);\n" +
"    setAdminPlanPrice(plan.price);\n" +
"    setAdminPlanPeriod(plan.billingPeriod);\n" +
"    setAdminPlanFeatures([...plan.features]);\n" +
"    setAdminPlanFeatureInput(\"\");\n" +
"    setAdminPlanActive(plan.isActive);\n" +
"    setAdminPlanDescription(plan.description || \"\");\n" +
"    setAdminPlanHighlighted(plan.highlighted || false);\n" +
"    setAdminPlanColor(plan.color || \"#f59e0b\");\n" +
"    setAdminPlanMaxReports(plan.maxReports ?? 0);\n" +
"    setAdminPlanSortOrder(plan.sortOrder ?? 99);\n" +
"    setAdminPlanModalOpen(true);\n" +
"  };",
"openEditPlan"
);

// 3. handleSavePlan plan object
rep(
"      const plan: Plan = {\n" +
"        id: adminEditPlan?.id,\n" +
"        name: adminPlanName.trim(),\n" +
"        price: adminPlanPrice,\n" +
"        billingPeriod: adminPlanPeriod,\n" +
"        features: adminPlanFeatures,\n" +
"        isActive: adminPlanActive\n" +
"      };",
"      const plan: Plan = {\n" +
"        id: adminEditPlan?.id,\n" +
"        name: adminPlanName.trim(),\n" +
"        price: adminPlanPrice,\n" +
"        billingPeriod: adminPlanPeriod,\n" +
"        features: adminPlanFeatures,\n" +
"        isActive: adminPlanActive,\n" +
"        description: adminPlanDescription.trim(),\n" +
"        highlighted: adminPlanHighlighted,\n" +
"        color: adminPlanColor,\n" +
"        maxReports: adminPlanMaxReports,\n" +
"        sortOrder: adminPlanSortOrder,\n" +
"      };",
"handleSavePlan plan object"
);

// 4. Plans Tab UI — replace between markers
const plansStart = "{/* \u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500 TAB: PLANS \u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500 */}";
const plansEnd   = "{/* \u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500 TAB: ANALYTICS \u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500 */}";
const si = c.indexOf(plansStart);
const ei = c.indexOf(plansEnd);
console.log('Plans tab: si=' + si + ' ei=' + ei);
if (si >= 0 && ei > si) {
  const p = '                  ';
  const newTab =
    plansStart + '\r\n' +
    p + '{adminTab === "plans" && (\r\n' +
    p + '  <div>\r\n' +
    p + '    {/* Header */}\r\n' +
    p + '    <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: "1.25rem", flexWrap: "wrap", gap: "8px" }}>\r\n' +
    p + '      <div>\r\n' +
    p + '        <div style={{ fontWeight: 700, fontSize: "14px" }}>📦 Plan Packages</div>\r\n' +
    p + '        <div style={{ fontSize: "12px", color: "var(--slate)", marginTop: "2px" }}>Manage subscription tiers — changes appear live in checkout.</div>\r\n' +
    p + '      </div>\r\n' +
    p + '      <button type="button" onClick={openNewPlan}\r\n' +
    p + '        style={{ padding: "9px 18px", borderRadius: "10px", background: "var(--amber)", color: "#000", border: "none", fontWeight: 800, fontSize: "12px", cursor: "pointer" }}>\r\n' +
    p + '        ➕ New Plan\r\n' +
    p + '      </button>\r\n' +
    p + '    </div>\r\n' +
    p + '    {/* Plan Cards */}\r\n' +
    p + '    <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill, minmax(270px, 1fr))", gap: "16px" }}>\r\n' +
    p + '      {adminPlans.map(plan => {\r\n' +
    p + '        const ac = plan.color || "#f59e0b";\r\n' +
    p + '        return (\r\n' +
    p + '          <div key={plan.id || plan.name} style={{ padding: "1.5rem", borderRadius: "16px", border: `2px solid ${plan.isActive ? ac : "var(--hairline)"}`, background: `linear-gradient(135deg, ${ac}0a, transparent 70%)`, position: "relative" }}>\r\n' +
    p + '            {plan.highlighted && plan.isActive && (\r\n' +
    p + '              <div style={{ position: "absolute", top: "-12px", left: "50%", transform: "translateX(-50%)", background: ac, color: "#000", fontSize: "9px", fontWeight: 800, padding: "3px 12px", borderRadius: "20px", whiteSpace: "nowrap" }}>⭐ RECOMMENDED</div>\r\n' +
    p + '            )}\r\n' +
    p + '            <div style={{ display: "flex", justifyContent: "space-between", marginBottom: "10px" }}>\r\n' +
    p + '              <span style={{ fontSize: "10px", padding: "2px 8px", borderRadius: "20px", fontWeight: 700, background: plan.isActive ? "rgba(16,185,129,0.15)" : "rgba(100,116,139,0.15)", color: plan.isActive ? "#10b981" : "#64748b" }}>\r\n' +
    p + '                {plan.isActive ? "● ACTIVE" : "○ INACTIVE"}\r\n' +
    p + '              </span>\r\n' +
    p + '              <span style={{ fontSize: "10px", color: "var(--slate)", fontFamily: "monospace" }}>#{plan.sortOrder ?? "—"}</span>\r\n' +
    p + '            </div>\r\n' +
    p + '            <div style={{ fontWeight: 800, fontSize: "18px", color: ac }}>{plan.name}</div>\r\n' +
    p + '            {plan.description && <div style={{ fontSize: "11px", color: "var(--slate)", marginBottom: "4px", fontStyle: "italic" }}>{plan.description}</div>}\r\n' +
    p + '            <div style={{ fontSize: "26px", fontWeight: 900, marginTop: "6px", color: plan.price === 0 ? "#10b981" : "var(--text)" }}>\r\n' +
    p + '              {plan.price === 0 ? "Free" : `\u20b9${plan.price.toLocaleString()}`}\r\n' +
    p + '              {plan.price > 0 && <span style={{ fontSize: "12px", fontWeight: 400, color: "var(--slate)", marginLeft: "4px" }}>/{plan.billingPeriod}</span>}\r\n' +
    p + '            </div>\r\n' +
    p + '            <div style={{ fontSize: "11px", color: "var(--slate)", margin: "4px 0 10px" }}>\r\n' +
    p + '              {(plan.maxReports ?? 0) === 0 ? "\u267e\ufe0f Unlimited reports" : `\ud83d\udcca ${plan.maxReports} reports/${plan.billingPeriod}`}\r\n' +
    p + '            </div>\r\n' +
    p + '            <div style={{ display: "flex", flexWrap: "wrap", gap: "5px", marginBottom: "14px" }}>\r\n' +
    p + '              {plan.features.map((f, i) => (\r\n' +
    p + '                <span key={i} style={{ fontSize: "10px", padding: "2px 8px", borderRadius: "20px", background: `${ac}18`, color: ac, fontWeight: 600 }}>\u2713 {f}</span>\r\n' +
    p + '              ))}\r\n' +
    p + '            </div>\r\n' +
    p + '            <div style={{ display: "flex", gap: "6px" }}>\r\n' +
    p + '              <button type="button" onClick={() => openEditPlan(plan)} style={{ flex: 1, padding: "8px 0", borderRadius: "8px", fontSize: "12px", fontWeight: 700, cursor: "pointer", border: `1px solid ${ac}`, background: `${ac}18`, color: ac }}>\u270f\ufe0f Edit</button>\r\n' +
    p + '              <button type="button" title="Duplicate" onClick={() => { setAdminEditPlan(null); setAdminPlanName(plan.name + " (Copy)"); setAdminPlanPrice(plan.price); setAdminPlanPeriod(plan.billingPeriod); setAdminPlanFeatures([...plan.features]); setAdminPlanFeatureInput(""); setAdminPlanActive(false); setAdminPlanDescription(plan.description || ""); setAdminPlanHighlighted(false); setAdminPlanColor(plan.color || "#f59e0b"); setAdminPlanMaxReports(plan.maxReports ?? 0); setAdminPlanSortOrder((plan.sortOrder ?? 99) + 1); setAdminPlanModalOpen(true); }} style={{ padding: "8px 10px", borderRadius: "8px", fontSize: "12px", fontWeight: 700, cursor: "pointer", border: "1px solid var(--hairline)", background: "rgba(255,255,255,0.04)", color: "var(--slate)" }}>\ud83d\udccb</button>\r\n' +
    p + '              <button type="button" onClick={() => handleDeletePlan(plan)} style={{ padding: "8px 10px", borderRadius: "8px", fontSize: "12px", fontWeight: 700, cursor: "pointer", border: "1px solid #ef4444", background: "rgba(239,68,68,0.08)", color: "#ef4444" }}>\ud83d\uddd1\ufe0f</button>\r\n' +
    p + '            </div>\r\n' +
    p + '          </div>\r\n' +
    p + '        );\r\n' +
    p + '      })}\r\n' +
    p + '      {adminPlans.length === 0 && (\r\n' +
    p + '        <div style={{ gridColumn: "1/-1", padding: "3rem", textAlign: "center", color: "var(--slate)", fontSize: "13px", border: "2px dashed var(--hairline)", borderRadius: "16px" }}>\r\n' +
    p + '          <div style={{ fontSize: "3rem", marginBottom: "8px" }}>\ud83d\udce6</div>\r\n' +
    p + '          No plans yet. Click <strong>\u2795 New Plan</strong> to get started.\r\n' +
    p + '        </div>\r\n' +
    p + '      )}\r\n' +
    p + '    </div>\r\n' +
    p + '    {/* Plan Editor Side Panel */}\r\n' +
    p + '    {adminPlanModalOpen && (\r\n' +
    p + '      <div style={{ position: "fixed", top: 0, right: 0, width: "440px", height: "100vh", background: "var(--glass-bg)", backdropFilter: "blur(24px)", borderLeft: `2px solid ${adminPlanColor}`, zIndex: 9999, display: "flex", flexDirection: "column" }}>\r\n' +
    p + '        <div style={{ padding: "1.25rem 1.5rem", borderBottom: "1px solid var(--hairline)", display: "flex", alignItems: "center", justifyContent: "space-between", background: `${adminPlanColor}14`, flexShrink: 0 }}>\r\n' +
    p + '          <div>\r\n' +
    p + '            <div style={{ fontWeight: 800, fontSize: "15px", color: adminPlanColor }}>{adminEditPlan ? "\u270f\ufe0f Edit Plan" : "\u2795 Create Plan"}</div>\r\n' +
    p + '            <div style={{ fontSize: "11px", color: "var(--slate)", marginTop: "2px" }}>{adminEditPlan ? `Editing: ${adminEditPlan.name}` : "New package"}</div>\r\n' +
    p + '          </div>\r\n' +
    p + '          <button type="button" onClick={() => setAdminPlanModalOpen(false)} style={{ background: "none", border: "none", color: "var(--slate)", fontSize: "20px", cursor: "pointer" }}>\u2715</button>\r\n' +
    p + '        </div>\r\n' +
    p + '        <div style={{ padding: "1.25rem 1.5rem", display: "flex", flexDirection: "column", gap: "14px", overflowY: "auto", flex: 1 }}>\r\n' +
    p + '          <div className="form-group"><label className="form-label">Plan Name *</label><input className="form-input" type="text" value={adminPlanName} onChange={e => setAdminPlanName(e.target.value)} placeholder="e.g. Starter, Pro, Enterprise..." /></div>\r\n' +
    p + '          <div className="form-group"><label className="form-label">Tagline / Description</label><input className="form-input" type="text" value={adminPlanDescription} onChange={e => setAdminPlanDescription(e.target.value)} placeholder="e.g. Perfect for small teams" /></div>\r\n' +
    p + '          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: "10px" }}>\r\n' +
    p + '            <div className="form-group"><label className="form-label">Price (\u20b9) \u2014 0 = Free</label><input className="form-input" type="number" min={0} value={adminPlanPrice} onChange={e => setAdminPlanPrice(Number(e.target.value))} /></div>\r\n' +
    p + '            <div className="form-group"><label className="form-label">Billing Period</label><select className="form-input" value={adminPlanPeriod} onChange={e => setAdminPlanPeriod(e.target.value as Plan["billingPeriod"])} style={{ appearance: "none", cursor: "pointer" }}>{(["free","monthly","yearly","lifetime"] as const).map(per => (<option key={per} value={per}>{per.charAt(0).toUpperCase() + per.slice(1)}</option>))}</select></div>\r\n' +
    p + '          </div>\r\n' +
    p + '          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: "10px" }}>\r\n' +
    p + '            <div className="form-group"><label className="form-label">Report Limit (0 = \u221e)</label><input className="form-input" type="number" min={0} value={adminPlanMaxReports} onChange={e => setAdminPlanMaxReports(Number(e.target.value))} /></div>\r\n' +
    p + '            <div className="form-group"><label className="form-label">Sort Order</label><input className="form-input" type="number" min={0} value={adminPlanSortOrder} onChange={e => setAdminPlanSortOrder(Number(e.target.value))} /></div>\r\n' +
    p + '          </div>\r\n' +
    p + '          <div className="form-group"><label className="form-label">Accent Color</label><div style={{ display: "flex", gap: "10px", alignItems: "center" }}><input type="color" value={adminPlanColor} onChange={e => setAdminPlanColor(e.target.value)} style={{ width: "48px", height: "40px", padding: "2px", borderRadius: "8px", border: "1px solid var(--hairline)", cursor: "pointer" }} /><div style={{ flex: 1, height: "40px", borderRadius: "8px", border: `2px solid ${adminPlanColor}`, background: `${adminPlanColor}14`, display: "flex", alignItems: "center", justifyContent: "center", fontSize: "13px", fontWeight: 800, color: adminPlanColor }}>{adminPlanName || "Preview"}</div></div></div>\r\n' +
    p + '          <div className="form-group"><label className="form-label">Features</label><div style={{ display: "flex", gap: "6px" }}><input className="form-input" type="text" value={adminPlanFeatureInput} onChange={e => setAdminPlanFeatureInput(e.target.value)} onKeyDown={e => { if (e.key === "Enter" && adminPlanFeatureInput.trim()) { setAdminPlanFeatures(prev => [...prev, adminPlanFeatureInput.trim()]); setAdminPlanFeatureInput(""); e.preventDefault(); }}} placeholder="Type feature, press Enter..." /><button type="button" onClick={() => { if (adminPlanFeatureInput.trim()) { setAdminPlanFeatures(prev => [...prev, adminPlanFeatureInput.trim()]); setAdminPlanFeatureInput(""); }}} style={{ padding: "0 14px", borderRadius: "8px", border: "none", background: "var(--amber)", color: "#000", fontWeight: 700, cursor: "pointer", fontSize: "18px" }}>+</button></div><div style={{ display: "flex", flexWrap: "wrap", gap: "6px", marginTop: "8px" }}>{adminPlanFeatures.map((f, i) => (<span key={i} style={{ fontSize: "11px", padding: "4px 10px", borderRadius: "20px", background: `${adminPlanColor}18`, color: adminPlanColor, display: "flex", alignItems: "center", gap: "5px", fontWeight: 600 }}>{f}<button type="button" onClick={() => setAdminPlanFeatures(prev => prev.filter((_, j) => j !== i))} style={{ background: "none", border: "none", color: "#ef4444", cursor: "pointer", padding: 0, fontSize: "12px" }}>\u2715</button></span>))}</div></div>\r\n' +
    p + '          <div style={{ display: "flex", gap: "10px" }}>\r\n' +
    p + '            <label style={{ flex: 1, display: "flex", alignItems: "center", gap: "8px", padding: "10px 12px", borderRadius: "8px", cursor: "pointer", border: `1px solid ${adminPlanActive ? "#10b981" : "var(--hairline)"}`, background: "rgba(255,255,255,0.03)" }}><input id="plan-active-chk" type="checkbox" checked={adminPlanActive} onChange={e => setAdminPlanActive(e.target.checked)} style={{ width: "16px", height: "16px", accentColor: "#10b981" }} /><span style={{ fontSize: "12px", fontWeight: 700 }}>Active</span></label>\r\n' +
    p + '            <label style={{ flex: 1, display: "flex", alignItems: "center", gap: "8px", padding: "10px 12px", borderRadius: "8px", cursor: "pointer", border: `1px solid ${adminPlanHighlighted ? adminPlanColor : "var(--hairline)"}`, background: "rgba(255,255,255,0.03)" }}><input id="plan-highlight-chk" type="checkbox" checked={adminPlanHighlighted} onChange={e => setAdminPlanHighlighted(e.target.checked)} style={{ width: "16px", height: "16px", accentColor: adminPlanColor }} /><span style={{ fontSize: "12px", fontWeight: 700" }}>\u2b50 Recommended</span></label>\r\n' +
    p + '          </div>\r\n' +
    p + '        </div>\r\n' +
    p + '        <div style={{ padding: "1rem 1.5rem", borderTop: "1px solid var(--hairline)", display: "flex", gap: "8px", flexShrink: 0 }}>\r\n' +
    p + '          <button type="button" onClick={() => setAdminPlanModalOpen(false)} style={{ flex: 1, padding: "11px", border: "1px solid var(--hairline)", borderRadius: "10px", background: "transparent", color: "var(--slate)", cursor: "pointer", fontWeight: 700 }}>Cancel</button>\r\n' +
    p + '          <button type="button" onClick={handleSavePlan} disabled={adminLoading} style={{ flex: 2, padding: "11px", border: "none", borderRadius: "10px", background: adminPlanColor, color: "#000", cursor: "pointer", fontWeight: 800, fontSize: "13px", opacity: adminLoading ? 0.6 : 1 }}>{adminLoading ? "Saving..." : (adminEditPlan ? "\ud83d\udcbe Save Changes" : "\u2705 Create Plan")}</button>\r\n' +
    p + '        </div>\r\n' +
    p + '      </div>\r\n' +
    p + '    )}\r\n' +
    p + '  </div>\r\n' +
    p + ')}\r\n\r\n' +
    p;
  c = c.substring(0, si) + newTab + c.substring(ei);
  console.log('v Plans Tab UI replaced');
  changed++;
}

// 5. Checkout modal plan selection
const oldCOSum = 'className="payment-summary-box">\r\n                <span>\ud83d\ude80 SheetCodeCrest Pro Lifetime</span>\r\n                <strong>?1,599</strong>\r\n              </div>';
if (c.includes(oldCOSum)) {
  const newCOSum = 
    '{checkoutPlans.length > 0 && (\r\n' +
    '                <div style={{ marginBottom: "1.25rem" }}>\r\n' +
    '                  <div style={{ fontSize: "11px", fontWeight: 700, textTransform: "uppercase", letterSpacing: "0.06em", color: "var(--slate)", marginBottom: "10px" }}>Choose Your Plan</div>\r\n' +
    '                  <div style={{ display: "grid", gridTemplateColumns: checkoutPlans.length > 2 ? "repeat(3,1fr)" : checkoutPlans.length === 2 ? "1fr 1fr" : "1fr", gap: "10px" }}>\r\n' +
    '                    {checkoutPlans.map(plan => { const isSel = selectedPlanId === plan.id; const ac = plan.color || "#f59e0b"; return (\r\n' +
    '                      <button key={plan.id || plan.name} type="button" onClick={() => setSelectedPlanId(plan.id || null)}\r\n' +
    '                        style={{ position: "relative", padding: "14px 10px", borderRadius: "12px", border: `2px solid ${isSel ? ac : "var(--hairline)"}`, background: isSel ? `${ac}14` : "rgba(255,255,255,0.02)", cursor: "pointer", textAlign: "left", transition: "all 0.15s" }}>\r\n' +
    '                        {plan.highlighted && <div style={{ position: "absolute", top: "-9px", left: "50%", transform: "translateX(-50%)", background: ac, color: "#000", fontSize: "7px", fontWeight: 800, padding: "2px 8px", borderRadius: "20px", whiteSpace: "nowrap" }}>\u2b50 BEST</div>}\r\n' +
    '                        <div style={{ fontWeight: 800, fontSize: "13px", color: isSel ? ac : "var(--text)" }}>{plan.name}</div>\r\n' +
    '                        {plan.description && <div style={{ fontSize: "9px", color: "var(--slate)", marginTop: "1px" }}>{plan.description}</div>}\r\n' +
    '                        <div style={{ fontWeight: 900, fontSize: "18px", color: ac, marginTop: "4px" }}>\u20b9{plan.price.toLocaleString()}<span style={{ fontSize: "9px", fontWeight: 400, color: "var(--slate)", marginLeft: "2px" }}>/{plan.billingPeriod}</span></div>\r\n' +
    '                        {isSel && <div style={{ position: "absolute", top: "8px", right: "8px", width: "14px", height: "14px", borderRadius: "50%", background: ac, display: "flex", alignItems: "center", justifyContent: "center", fontSize: "8px", color: "#000", fontWeight: 900 }}>\u2713</div>}\r\n' +
    '                      </button>\r\n' +
    '                    ); })}\r\n' +
    '                  </div>\r\n' +
    '                </div>\r\n' +
    '              )}\r\n' +
    '              <div className="payment-summary-box">\r\n' +
    '                {(() => { const sel = checkoutPlans.find(p => p.id === selectedPlanId) || checkoutPlans[0]; return sel ? <><span>\ud83d\ude80 {sel.name}</span><strong>\u20b9{sel.price.toLocaleString()}</strong></> : <><span>\ud83d\ude80 SheetCodeCrest Pro</span><strong>\u20b91,599</strong></>; })()}\r\n' +
    '              </div';
  c = c.replace(oldCOSum, newCOSum);
  console.log('v Checkout modal updated');
  changed++;
} else {
  const idx = c.indexOf('payment-summary-box');
  if (idx >= 0) console.log('payment-summary-box context:', JSON.stringify(c.substring(idx, idx+120)));
  else console.log('x payment-summary-box not found at all');
}

fs.writeFileSync(file, c, 'utf-8');
console.log('\nDone! Changes: ' + changed + ', size: ' + c.length);
