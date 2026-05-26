# upgrade-plans-feature.ps1
# Updates plan management functions + plans tab UI + checkout modal in App.tsx

$file = "src\App.tsx"
$content = [System.IO.File]::ReadAllText((Resolve-Path $file), [System.Text.Encoding]::UTF8)
$origLength = $content.Length
Write-Host "Original file: $origLength chars, $((($content -split "`n").Count)) lines"

# =============================================
# 1. Update openNewPlan to reset extended fields
# =============================================
$old1 = @'
  const openNewPlan = () => {
    setAdminEditPlan(null);
    setAdminPlanName("");
    setAdminPlanPrice(0);
    setAdminPlanPeriod("monthly");
    setAdminPlanFeatures([]);
    setAdminPlanFeatureInput("");
    setAdminPlanActive(true);
    setAdminPlanModalOpen(true);
  };
'@
$new1 = @'
  const openNewPlan = () => {
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
  };
'@
if ($content.Contains($old1.Trim())) {
    $content = $content.Replace($old1.Trim(), $new1.Trim())
    Write-Host "✓ Updated openNewPlan"
} else { Write-Host "✗ openNewPlan pattern not found" }

# =============================================
# 2. Update openEditPlan to populate extended fields
# =============================================
$old2 = @'
  const openEditPlan = (plan: Plan) => {
    setAdminEditPlan(plan);
    setAdminPlanName(plan.name);
    setAdminPlanPrice(plan.price);
    setAdminPlanPeriod(plan.billingPeriod);
    setAdminPlanFeatures([...plan.features]);
    setAdminPlanFeatureInput("");
    setAdminPlanActive(plan.isActive);
    setAdminPlanModalOpen(true);
  };
'@
$new2 = @'
  const openEditPlan = (plan: Plan) => {
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
  };
'@
if ($content.Contains($old2.Trim())) {
    $content = $content.Replace($old2.Trim(), $new2.Trim())
    Write-Host "✓ Updated openEditPlan"
} else { Write-Host "✗ openEditPlan pattern not found" }

# =============================================
# 3. Update handleSavePlan to persist new fields
# =============================================
$old3 = @'
      const plan: Plan = {
        id: adminEditPlan?.id,
        name: adminPlanName.trim(),
        price: adminPlanPrice,
        billingPeriod: adminPlanPeriod,
        features: adminPlanFeatures,
        isActive: adminPlanActive
      };
'@
$new3 = @'
      const plan: Plan = {
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
      };
'@
if ($content.Contains($old3.Trim())) {
    $content = $content.Replace($old3.Trim(), $new3.Trim())
    Write-Host "✓ Updated handleSavePlan plan object"
} else { Write-Host "✗ handleSavePlan pattern not found" }

# =============================================
# 4. Add loadCheckoutPlans after loadAdminData block
# =============================================
$insertAfterLoadAdmin = "  const handleDeletePlan = async (plan: Plan) => {"
$loadCheckoutFn = @'
  // Load plans for the checkout modal
  const loadCheckoutPlans = async () => {
    try {
      const plans = await dbGetPlans();
      const activePlans = plans.filter(p => p.isActive && p.price > 0);
      setCheckoutPlans(activePlans);
      if (activePlans.length > 0 && !selectedPlanId) {
        const recommended = activePlans.find(p => p.highlighted) || activePlans[0];
        setSelectedPlanId(recommended.id || null);
      }
    } catch (err) {
      console.error("Failed to load checkout plans", err);
    }
  };

'@
if ($content.Contains($insertAfterLoadAdmin) -and -not $content.Contains("loadCheckoutPlans")) {
    $content = $content.Replace($insertAfterLoadAdmin, $loadCheckoutFn + $insertAfterLoadAdmin)
    Write-Host "✓ Added loadCheckoutPlans function"
} else { Write-Host "✗ loadCheckoutPlans insertion skipped (already exists or anchor not found)" }

# =============================================
# 5. Call loadCheckoutPlans when checkoutOpen opens
#    Find the checkoutOpen useState and add a useEffect for it
# =============================================
$checkoutEffectAnchor = "  const recordSuccessfulReport = useCallback"
$checkoutEffect = @'
  // Load available plans when the checkout modal opens
  useEffect(() => {
    if (checkoutOpen) {
      loadCheckoutPlans();
    }
  }, [checkoutOpen]);

'@
if ($content.Contains($checkoutEffectAnchor) -and -not $content.Contains("if (checkoutOpen)")) {
    $content = $content.Replace($checkoutEffectAnchor, $checkoutEffect + $checkoutEffectAnchor)
    Write-Host "✓ Added checkout useEffect"
} else { Write-Host "✗ checkout useEffect skipped" }

# =============================================
# 6. Replace the Plans Tab UI with enhanced version
# =============================================
# Find start marker
$plansTabStart = "{/* ─────────── TAB: PLANS ─────────── */}"
$plansTabEnd = "{/* ─────────── TAB: ANALYTICS ─────────── */}"

$startIdx = $content.IndexOf($plansTabStart)
$endIdx = $content.IndexOf($plansTabEnd)

if ($startIdx -ge 0 -and $endIdx -gt $startIdx) {
    $newPlansTab = @'
                  {/* ─────────── TAB: PLANS ─────────── */}
                  {adminTab === "plans" && (
                    <div>
                      {/* Header row */}
                      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: "1rem", flexWrap: "wrap", gap: "8px" }}>
                        <div style={{ fontSize: "13px", color: "var(--slate)" }}>
                          Manage subscription packages. Changes appear instantly in the checkout modal.
                        </div>
                        <button type="button" onClick={openNewPlan}
                          style={{ padding: "8px 18px", borderRadius: "8px", background: "var(--amber)", color: "#000", border: "none", fontWeight: 700, fontSize: "12px", cursor: "pointer", display: "flex", alignItems: "center", gap: "6px" }}>
                          ➕ New Plan
                        </button>
                      </div>

                      {/* Plan Cards Grid */}
                      <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill, minmax(280px, 1fr))", gap: "16px" }}>
                        {adminPlans.map(plan => {
                          const accentColor = plan.color || "#f59e0b";
                          return (
                            <div key={plan.id || plan.name} style={{ padding: "1.5rem", borderRadius: "16px", border: `2px solid ${plan.isActive ? accentColor : "var(--hairline)"}`, background: `linear-gradient(135deg, ${plan.isActive ? accentColor + "0a" : "rgba(255,255,255,0.02)"}, transparent)`, position: "relative", transition: "box-shadow 0.2s" }}>
                              {/* Highlighted badge */}
                              {plan.highlighted && plan.isActive && (
                                <div style={{ position: "absolute", top: "-12px", left: "50%", transform: "translateX(-50%)", background: accentColor, color: "#000", fontSize: "9px", fontWeight: 800, padding: "3px 12px", borderRadius: "20px", letterSpacing: "0.08em", whiteSpace: "nowrap" }}>
                                  ⭐ RECOMMENDED
                                </div>
                              )}
                              {/* Status badge */}
                              <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: "10px" }}>
                                <span style={{ fontSize: "10px", padding: "2px 8px", borderRadius: "20px", fontWeight: 700, background: plan.isActive ? "rgba(16,185,129,0.15)" : "rgba(100,116,139,0.15)", color: plan.isActive ? "#10b981" : "#64748b" }}>
                                  {plan.isActive ? "● ACTIVE" : "○ INACTIVE"}
                                </span>
                                <span style={{ fontSize: "10px", color: "var(--slate)", fontFamily: "monospace" }}>#{plan.sortOrder ?? "—"}</span>
                              </div>

                              {/* Name & Description */}
                              <div style={{ fontWeight: 800, fontSize: "18px", color: accentColor, marginBottom: "2px" }}>{plan.name}</div>
                              {plan.description && <div style={{ fontSize: "12px", color: "var(--slate)", marginBottom: "10px", fontStyle: "italic" }}>{plan.description}</div>}

                              {/* Price */}
                              <div style={{ fontSize: "28px", fontWeight: 900, marginBottom: "4px", color: plan.price === 0 ? "#10b981" : "var(--text)" }}>
                                {plan.price === 0 ? "Free" : `₹${plan.price.toLocaleString()}`}
                                {plan.price > 0 && <span style={{ fontSize: "12px", fontWeight: 400, color: "var(--slate)", marginLeft: "4px" }}>/{plan.billingPeriod}</span>}
                              </div>

                              {/* Reports cap */}
                              <div style={{ fontSize: "11px", color: "var(--slate)", marginBottom: "10px" }}>
                                {plan.maxReports === 0 ? "♾️ Unlimited reports" : `📊 ${plan.maxReports} reports/${plan.billingPeriod}`}
                              </div>

                              {/* Features */}
                              <div style={{ display: "flex", flexWrap: "wrap", gap: "5px", marginBottom: "14px" }}>
                                {plan.features.map((f, i) => (
                                  <span key={i} style={{ fontSize: "10px", padding: "2px 8px", borderRadius: "20px", background: `${accentColor}18`, color: accentColor, fontWeight: 500 }}>✓ {f}</span>
                                ))}
                              </div>

                              {/* Action Buttons */}
                              <div style={{ display: "flex", gap: "6px" }}>
                                <button type="button" onClick={() => openEditPlan(plan)}
                                  style={{ flex: 1, padding: "7px 0", borderRadius: "8px", fontSize: "11px", fontWeight: 700, cursor: "pointer", border: `1px solid ${accentColor}`, background: `${accentColor}18`, color: accentColor }}>
                                  ✏️ Edit
                                </button>
                                <button type="button"
                                  onClick={() => {
                                    const dup: Plan = { ...plan, id: undefined, name: plan.name + " (Copy)", sortOrder: (plan.sortOrder ?? 99) + 1 };
                                    setAdminEditPlan(null);
                                    setAdminPlanName(dup.name);
                                    setAdminPlanPrice(dup.price);
                                    setAdminPlanPeriod(dup.billingPeriod);
                                    setAdminPlanFeatures([...dup.features]);
                                    setAdminPlanFeatureInput("");
                                    setAdminPlanActive(false);
                                    setAdminPlanDescription(dup.description || "");
                                    setAdminPlanHighlighted(false);
                                    setAdminPlanColor(dup.color || "#f59e0b");
                                    setAdminPlanMaxReports(dup.maxReports ?? 0);
                                    setAdminPlanSortOrder(dup.sortOrder ?? 99);
                                    setAdminPlanModalOpen(true);
                                  }}
                                  style={{ padding: "7px 10px", borderRadius: "8px", fontSize: "11px", fontWeight: 700, cursor: "pointer", border: "1px solid var(--hairline)", background: "rgba(255,255,255,0.04)", color: "var(--slate)" }}>
                                  📋
                                </button>
                                <button type="button" onClick={() => handleDeletePlan(plan)}
                                  style={{ padding: "7px 10px", borderRadius: "8px", fontSize: "11px", fontWeight: 700, cursor: "pointer", border: "1px solid #ef4444", background: "rgba(239,68,68,0.08)", color: "#ef4444" }}>
                                  🗑️
                                </button>
                              </div>
                            </div>
                          );
                        })}

                        {/* Empty state */}
                        {adminPlans.length === 0 && (
                          <div style={{ gridColumn: "1/-1", padding: "3rem", textAlign: "center", color: "var(--slate)", fontSize: "13px", border: "2px dashed var(--hairline)", borderRadius: "16px" }}>
                            <div style={{ fontSize: "32px", marginBottom: "8px" }}>📦</div>
                            No plans configured yet. Click <strong>➕ New Plan</strong> to get started.
                          </div>
                        )}
                      </div>

                      {/* Plan Edit/Create Side Panel */}
                      {adminPlanModalOpen && (
                        <div style={{ position: "fixed", top: 0, right: 0, width: "440px", height: "100vh", background: "var(--glass-bg)", backdropFilter: "blur(24px)", borderLeft: "2px solid var(--hairline)", zIndex: 9999, padding: "0", display: "flex", flexDirection: "column", overflowY: "auto" }}>
                          {/* Panel Header */}
                          <div style={{ padding: "1.5rem 1.5rem 1rem", borderBottom: "1px solid var(--hairline)", display: "flex", alignItems: "center", justifyContent: "space-between", background: `${adminPlanColor}12`, flexShrink: 0 }}>
                            <div>
                              <h4 style={{ margin: 0, fontSize: "16px", fontWeight: 800, color: adminPlanColor }}>
                                {adminEditPlan ? "✏️ Edit Plan" : "➕ Create New Plan"}
                              </h4>
                              <div style={{ fontSize: "11px", color: "var(--slate)", marginTop: "2px" }}>
                                {adminEditPlan ? `Editing: ${adminEditPlan.name}` : "New package configuration"}
                              </div>
                            </div>
                            <button type="button" onClick={() => setAdminPlanModalOpen(false)} style={{ background: "none", border: "none", color: "var(--slate)", fontSize: "20px", cursor: "pointer", padding: "4px" }}>✕</button>
                          </div>

                          {/* Panel Body */}
                          <div style={{ padding: "1.5rem", display: "flex", flexDirection: "column", gap: "1rem", flex: 1 }}>

                            {/* Name */}
                            <div className="form-group">
                              <label className="form-label">Plan Name *</label>
                              <input className="form-input" type="text" value={adminPlanName} onChange={e => setAdminPlanName(e.target.value)} placeholder="e.g. Starter, Pro, Enterprise..." />
                            </div>

                            {/* Description */}
                            <div className="form-group">
                              <label className="form-label">Tagline / Description</label>
                              <input className="form-input" type="text" value={adminPlanDescription} onChange={e => setAdminPlanDescription(e.target.value)} placeholder="e.g. Perfect for growing teams" />
                            </div>

                            {/* Price & Period row */}
                            <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: "10px" }}>
                              <div className="form-group">
                                <label className="form-label">Price (₹) — 0 = Free</label>
                                <input className="form-input" type="number" min={0} value={adminPlanPrice} onChange={e => setAdminPlanPrice(Number(e.target.value))} />
                              </div>
                              <div className="form-group">
                                <label className="form-label">Billing Period</label>
                                <select className="form-input" value={adminPlanPeriod} onChange={e => setAdminPlanPeriod(e.target.value as Plan["billingPeriod"])} style={{ appearance: "none", cursor: "pointer" }}>
                                  {(["free","monthly","yearly","lifetime"] as const).map(p => (
                                    <option key={p} value={p}>{p.charAt(0).toUpperCase() + p.slice(1)}</option>
                                  ))}
                                </select>
                              </div>
                            </div>

                            {/* Max Reports & Sort Order row */}
                            <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: "10px" }}>
                              <div className="form-group">
                                <label className="form-label">Report Limit (0 = ∞)</label>
                                <input className="form-input" type="number" min={0} value={adminPlanMaxReports} onChange={e => setAdminPlanMaxReports(Number(e.target.value))} />
                              </div>
                              <div className="form-group">
                                <label className="form-label">Sort Order</label>
                                <input className="form-input" type="number" min={0} value={adminPlanSortOrder} onChange={e => setAdminPlanSortOrder(Number(e.target.value))} />
                              </div>
                            </div>

                            {/* Accent Colour */}
                            <div className="form-group">
                              <label className="form-label">Accent Color (card highlight)</label>
                              <div style={{ display: "flex", gap: "10px", alignItems: "center" }}>
                                <input type="color" value={adminPlanColor} onChange={e => setAdminPlanColor(e.target.value)}
                                  style={{ width: "48px", height: "40px", padding: "2px", borderRadius: "8px", border: "1px solid var(--hairline)", cursor: "pointer", background: "none" }} />
                                <div style={{ flex: 1, height: "40px", borderRadius: "8px", border: `2px solid ${adminPlanColor}`, background: `${adminPlanColor}18`, display: "flex", alignItems: "center", justifyContent: "center", fontSize: "12px", fontWeight: 700, color: adminPlanColor }}>
                                  {adminPlanName || "Plan Preview"}
                                </div>
                              </div>
                            </div>

                            {/* Features */}
                            <div className="form-group">
                              <label className="form-label">Features (add one at a time)</label>
                              <div style={{ display: "flex", gap: "6px" }}>
                                <input className="form-input" type="text" value={adminPlanFeatureInput} onChange={e => setAdminPlanFeatureInput(e.target.value)}
                                  onKeyDown={e => { if (e.key === "Enter" && adminPlanFeatureInput.trim()) { setAdminPlanFeatures(prev => [...prev, adminPlanFeatureInput.trim()]); setAdminPlanFeatureInput(""); e.preventDefault(); }}}
                                  placeholder="Type feature, press Enter..." />
                                <button type="button" onClick={() => { if (adminPlanFeatureInput.trim()) { setAdminPlanFeatures(prev => [...prev, adminPlanFeatureInput.trim()]); setAdminPlanFeatureInput(""); }}}
                                  style={{ padding: "0 14px", borderRadius: "8px", border: "none", background: "var(--amber)", color: "#000", fontWeight: 700, cursor: "pointer", fontSize: "18px" }}>+</button>
                              </div>
                              <div style={{ display: "flex", flexWrap: "wrap", gap: "6px", marginTop: "8px" }}>
                                {adminPlanFeatures.map((f, i) => (
                                  <span key={i} style={{ fontSize: "11px", padding: "4px 10px", borderRadius: "20px", background: `${adminPlanColor}18`, color: adminPlanColor, display: "flex", alignItems: "center", gap: "5px", fontWeight: 600 }}>
                                    {f}
                                    <button type="button" onClick={() => setAdminPlanFeatures(prev => prev.filter((_, j) => j !== i))}
                                      style={{ background: "none", border: "none", color: "#ef4444", cursor: "pointer", padding: 0, fontSize: "12px", lineHeight: 1 }}>✕</button>
                                  </span>
                                ))}
                              </div>
                            </div>

                            {/* Toggle row: Active + Highlighted */}
                            <div style={{ display: "flex", gap: "10px" }}>
                              <label style={{ flex: 1, display: "flex", alignItems: "center", gap: "8px", padding: "10px 12px", background: "rgba(255,255,255,0.04)", borderRadius: "8px", cursor: "pointer", border: `1px solid ${adminPlanActive ? "#10b981" : "var(--hairline)"}` }}>
                                <input type="checkbox" checked={adminPlanActive} onChange={e => setAdminPlanActive(e.target.checked)} style={{ width: "16px", height: "16px", accentColor: "#10b981" }} />
                                <span style={{ fontSize: "12px", fontWeight: 700 }}>Active</span>
                              </label>
                              <label style={{ flex: 1, display: "flex", alignItems: "center", gap: "8px", padding: "10px 12px", background: "rgba(255,255,255,0.04)", borderRadius: "8px", cursor: "pointer", border: `1px solid ${adminPlanHighlighted ? adminPlanColor : "var(--hairline)"}` }}>
                                <input type="checkbox" checked={adminPlanHighlighted} onChange={e => setAdminPlanHighlighted(e.target.checked)} style={{ width: "16px", height: "16px", accentColor: adminPlanColor }} />
                                <span style={{ fontSize: "12px", fontWeight: 700 }}>⭐ Recommended</span>
                              </label>
                            </div>
                          </div>

                          {/* Panel Footer */}
                          <div style={{ padding: "1rem 1.5rem", borderTop: "1px solid var(--hairline)", display: "flex", gap: "8px", flexShrink: 0 }}>
                            <button type="button" onClick={() => setAdminPlanModalOpen(false)} style={{ flex: 1, padding: "11px", border: "1px solid var(--hairline)", borderRadius: "10px", background: "transparent", color: "var(--slate)", cursor: "pointer", fontWeight: 700 }}>Cancel</button>
                            <button type="button" onClick={handleSavePlan} disabled={adminLoading}
                              style={{ flex: 2, padding: "11px", border: "none", borderRadius: "10px", background: adminPlanColor, color: "#000", cursor: "pointer", fontWeight: 800, fontSize: "13px", opacity: adminLoading ? 0.6 : 1 }}>
                              {adminLoading ? "Saving..." : (adminEditPlan ? "💾 Save Changes" : "✅ Create Plan")}
                            </button>
                          </div>
                        </div>
                      )}
                    </div>
                  )}

'@
    $before = $content.Substring(0, $startIdx)
    $after = $content.Substring($endIdx)
    $content = $before + $newPlansTab + $after
    Write-Host "✓ Replaced Plans Tab UI"
} else {
    Write-Host "✗ Plans Tab markers not found (start=$startIdx, end=$endIdx)"
}

# =============================================
# 7. Update the Checkout Modal to show plan cards
# =============================================
$checkoutPaymentSummaryOld = @'
            <div className="modal-body">
              <div className="payment-summary-box">
                <span>🚀 SheetCodeCrest Pro Lifetime</span>
                <strong>?1,599</strong>
              </div>
'@
$checkoutPaymentSummaryNew = @'
            <div className="modal-body">
              {/* Plan Selection Cards */}
              {checkoutPlans.length > 0 && (
                <div style={{ marginBottom: "1.25rem" }}>
                  <div style={{ fontSize: "11px", fontWeight: 700, textTransform: "uppercase", letterSpacing: "0.06em", color: "var(--slate)", marginBottom: "10px" }}>
                    Choose Your Plan
                  </div>
                  <div style={{ display: "grid", gridTemplateColumns: checkoutPlans.length === 1 ? "1fr" : "repeat(auto-fit, minmax(140px, 1fr))", gap: "10px" }}>
                    {checkoutPlans.map(plan => {
                      const isSelected = selectedPlanId === plan.id;
                      const accentColor = plan.color || "#f59e0b";
                      return (
                        <button
                          key={plan.id || plan.name}
                          type="button"
                          onClick={() => setSelectedPlanId(plan.id || null)}
                          style={{ position: "relative", padding: "14px 12px", borderRadius: "12px", border: `2px solid ${isSelected ? accentColor : "var(--hairline)"}`, background: isSelected ? `${accentColor}14` : "rgba(255,255,255,0.03)", cursor: "pointer", textAlign: "left", transition: "all 0.15s" }}>
                          {plan.highlighted && (
                            <div style={{ position: "absolute", top: "-9px", left: "50%", transform: "translateX(-50%)", background: accentColor, color: "#000", fontSize: "8px", fontWeight: 800, padding: "2px 8px", borderRadius: "20px", whiteSpace: "nowrap" }}>⭐ BEST</div>
                          )}
                          <div style={{ fontWeight: 800, fontSize: "14px", color: isSelected ? accentColor : "var(--text)" }}>{plan.name}</div>
                          {plan.description && <div style={{ fontSize: "10px", color: "var(--slate)", marginTop: "2px" }}>{plan.description}</div>}
                          <div style={{ fontWeight: 900, fontSize: "20px", color: accentColor, marginTop: "6px" }}>
                            ₹{plan.price.toLocaleString()}
                            <span style={{ fontSize: "10px", fontWeight: 400, color: "var(--slate)", marginLeft: "3px" }}>/{plan.billingPeriod}</span>
                          </div>
                          <div style={{ marginTop: "6px", display: "flex", flexDirection: "column", gap: "2px" }}>
                            {plan.features.slice(0, 3).map((f, i) => (
                              <div key={i} style={{ fontSize: "9px", color: "var(--slate)", display: "flex", alignItems: "center", gap: "4px" }}>
                                <span style={{ color: accentColor }}>✓</span> {f}
                              </div>
                            ))}
                          </div>
                          {isSelected && (
                            <div style={{ position: "absolute", top: "8px", right: "8px", width: "16px", height: "16px", borderRadius: "50%", background: accentColor, display: "flex", alignItems: "center", justifyContent: "center", fontSize: "9px", color: "#000", fontWeight: 900 }}>✓</div>
                          )}
                        </button>
                      );
                    })}
                  </div>
                </div>
              )}

              {/* Selected Plan Summary */}
              {(() => {
                const selPlan = checkoutPlans.find(p => p.id === selectedPlanId) || checkoutPlans[0];
                return (
                  <div className="payment-summary-box">
                    <span>🚀 {selPlan ? selPlan.name : "SheetCodeCrest Pro"}</span>
                    <strong>₹{selPlan ? selPlan.price.toLocaleString() : "1,599"}</strong>
                  </div>
                );
              })()}
'@

if ($content.Contains($checkoutPaymentSummaryOld.Trim())) {
    $content = $content.Replace($checkoutPaymentSummaryOld.Trim(), $checkoutPaymentSummaryNew.Trim())
    Write-Host "✓ Updated checkout modal with plan selection"
} else { Write-Host "✗ Checkout modal summary pattern not found" }

# =============================================
# Write the file
# =============================================
[System.IO.File]::WriteAllText((Resolve-Path $file), $content, [System.Text.Encoding]::UTF8)
Write-Host ""
Write-Host "Done! New file: $($content.Length) chars"
