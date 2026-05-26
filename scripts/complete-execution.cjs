const fs = require('fs');
const path = require('path');

const file = path.resolve('src/App.tsx');
let c = fs.readFileSync(file, 'utf-8');
console.log('Original size:', c.length, 'chars');

const crlf = s => s.replace(/\r\n/g, '\n').replace(/\n/g, '\r\n');
let changed = 0;

// 1. Fix the syntax error: fontWeight: 700" -> fontWeight: 700
const oldSyntax = 'fontWeight: 700"';
const newSyntax = 'fontWeight: 700';

if (c.includes(oldSyntax)) {
  c = c.replace(oldSyntax, newSyntax);
  console.log('v Fixed syntax error in Recommended badge');
  changed++;
} else {
  console.log('x syntax error not found (might already be fixed)');
}

// 2. Replace hardcoded payment-summary-box
const oldBox = crlf(
  '              <div className="payment-summary-box">\r\n' +
  '                <span>🚀 SheetCodeCrest Pro Lifetime</span>\r\n' +
  '                <strong>₹1,599</strong>\r\n' +
  '              </div>'
);

const newBox = crlf(
  '              {checkoutPlans.length > 0 && (\r\n' +
  '                <div style={{ marginBottom: "1.25rem" }}>\r\n' +
  '                  <div style={{ fontSize: "11px", fontWeight: 700, textTransform: "uppercase", letterSpacing: "0.06em", color: "var(--slate)", marginBottom: "10px" }}>Choose Your Plan</div>\r\n' +
  '                  <div style={{ display: "grid", gridTemplateColumns: checkoutPlans.length > 2 ? "repeat(3, 1fr)" : checkoutPlans.length === 2 ? "1fr 1fr" : "1fr", gap: "10px" }}>\r\n' +
  '                    {checkoutPlans.map(plan => {\r\n' +
  '                      const isSel = selectedPlanId === plan.id;\r\n' +
  '                      const ac = plan.color || "#f59e0b";\r\n' +
  '                      return (\r\n' +
  '                        <button key={plan.id || plan.name} type="button" onClick={() => setSelectedPlanId(plan.id || null)}\r\n' +
  '                          style={{\r\n' +
  '                            position: "relative",\r\n' +
  '                            padding: "14px 10px",\r\n' +
  '                            borderRadius: "12px",\r\n' +
  '                            border: `2px solid ${isSel ? ac : "var(--hairline)"}`,\r\n' +
  '                            background: isSel ? `${ac}14` : "rgba(255,255,255,0.02)",\r\n' +
  '                            cursor: "pointer",\r\n' +
  '                            textAlign: "left",\r\n' +
  '                            transition: "all 0.15s"\r\n' +
  '                          }}>\r\n' +
  '                          {plan.highlighted && (\r\n' +
  '                            <div style={{ position: "absolute", top: "-9px", left: "50%", transform: "translateX(-50%)", background: ac, color: "#000", fontSize: "7px", fontWeight: 800, padding: "2px 8px", borderRadius: "20px", whiteSpace: "nowrap" }}>⭐ BEST</div>\r\n' +
  '                          )}\r\n' +
  '                          <div style={{ fontWeight: 800, fontSize: "13px", color: isSel ? ac : "var(--text)" }}>{plan.name}</div>\r\n' +
  '                          {plan.description && <div style={{ fontSize: "9px", color: "var(--slate)", marginTop: "1px" }}>{plan.description}</div>}\r\n' +
  '                          <div style={{ fontWeight: 900, fontSize: "18px", color: ac, marginTop: "4px" }}>\r\n' +
  '                            ₹{plan.price.toLocaleString()}\r\n' +
  '                            <span style={{ fontSize: "9px", fontWeight: 400, color: "var(--slate)", marginLeft: "2px" }}>/{plan.billingPeriod}</span>\r\n' +
  '                          </div>\r\n' +
  '                          {isSel && (\r\n' +
  '                            <div style={{ position: "absolute", top: "8px", right: "8px", width: "14px", height: "14px", borderRadius: "50%", background: ac, display: "flex", alignItems: "center", justifyContent: "center", fontSize: "8px", color: "#000", fontWeight: 900 }}>✓</div>\r\n' +
  '                          )}\r\n' +
  '                        </button>\r\n' +
  '                      );\r\n' +
  '                    })}\r\n' +
  '                  </div>\r\n' +
  '                </div>\r\n' +
  '              )}\r\n' +
  '              <div className="payment-summary-box">\r\n' +
  '                {(() => {\r\n' +
  '                  const sel = checkoutPlans.find(p => p.id === selectedPlanId) || checkoutPlans[0];\r\n' +
  '                  return sel ? (\r\n' +
  '                    <>\r\n' +
  '                      <span>🚀 {sel.name}</span>\r\n' +
  '                      <strong>₹{sel.price.toLocaleString()}</strong>\r\n' +
  '                    </>\r\n' +
  '                  ) : (\r\n' +
  '                    <>\r\n' +
  '                      <span>🚀 SheetCodeCrest Pro</span>\r\n' +
  '                      <strong>₹1,599</strong>\r\n' +
  '                    </>\r\n' +
  '                  );\r\n' +
  '                })()}\r\n' +
  '              </div>'
);

if (c.includes(oldBox)) {
  c = c.replace(oldBox, newBox);
  console.log('v Replaced payment summary box with dynamic selector');
  changed++;
} else if (c.includes('className="payment-summary-box"')) {
  // Let's do a more robust substring matching if there's a minor difference
  console.log('payment-summary-box present, attempting fallback match');
  const targetText = 'payment-summary-box';
  const firstIndex = c.indexOf('className="payment-summary-box"');
  // Make sure it's the one in JSX, which is at the end of the file
  const lastIndex = c.lastIndexOf('className="payment-summary-box"');
  console.log('Indices of payment-summary-box:', firstIndex, lastIndex);
  
  // Let's find the enclosing div: from <div className="payment-summary-box"> to </div>
  const startIdx = c.indexOf('<div className="payment-summary-box">', lastIndex - 50);
  const endIdx = c.indexOf('</div>', startIdx) + 6;
  if (startIdx >= 0 && endIdx > startIdx) {
    c = c.substring(0, startIdx) + newBox + c.substring(endIdx);
    console.log('v Replaced payment summary box via fallback indices');
    changed++;
  } else {
    console.log('x Failed to find start/end of payment summary box');
  }
} else {
  console.log('x payment-summary-box JSX not found');
}

fs.writeFileSync(file, c, 'utf-8');
console.log('Done! Changed elements:', changed, 'new size:', c.length);
