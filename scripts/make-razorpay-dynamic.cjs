const fs = require('fs');
const path = require('path');

const file = path.resolve('src/App.tsx');
let c = fs.readFileSync(file, 'utf-8');
console.log('Original size:', c.length, 'chars');

const crlf = s => s.replace(/\r\n/g, '\n').replace(/\n/g, '\r\n');
let changed = 0;

// 1. startPaymentSimulation
const oldSimStart = crlf(
  '    // Live Razorpay Mode\r\n' +
  '    const razorpayKey = import.meta.env.VITE_RAZORPAY_KEY_ID || "rzp_test_StUNrV1X2WAvV4";\r\n' +
  '    if (razorpayKey && typeof (window as any).Razorpay !== "undefined") {\r\n' +
  '      const options = {\r\n' +
  '        key: razorpayKey,\r\n' +
  '        amount: 159900, // ₹1,599 in paisa\r\n' +
  '        currency: "INR",\r\n' +
  '        name: "SheetCodeCrest Pro",\r\n' +
  '        description: "Premium Spreadsheet Analytics Subscription",\r\n' +
  '        image: "https://sheetcodecrest.vercel.app/logo.png",\r\n' +
  '        handler: async function (response: any) {\r\n' +
  '          try {\r\n' +
  '            addLog(`💳 Razorpay transaction completed! Payment ID: ${response.razorpay_payment_id}`, "success");\r\n' +
  '            \r\n' +
  '            if (currentUser) {\r\n' +
  '              const updatedUser = { ...currentUser, isPro: true };\r\n' +
  '              await dbSaveUser(updatedUser);\r\n' +
  '              setCurrentUser(updatedUser);\r\n' +
  '              \r\n' +
  '              // Log payment securely in Supabase\r\n' +
  '              const { dbLogPayment } = await import("./db");\r\n' +
  '              await dbLogPayment({\r\n' +
  '                username: currentUser.username,\r\n' +
  '                gateway: "razorpay",\r\n' +
  '                paymentId: response.razorpay_payment_id,\r\n' +
  '                orderId: response.razorpay_order_id || "",\r\n' +
  '                signature: response.razorpay_signature || "",\r\n' +
  '                amount: 1599,\r\n' +
  '                status: "success"\r\n' +
  '              });'
);

const newSimStart = crlf(
  '    // Live Razorpay Mode\r\n' +
  '    const razorpayKey = import.meta.env.VITE_RAZORPAY_KEY_ID || "rzp_test_StUNrV1X2WAvV4";\r\n' +
  '    if (razorpayKey && typeof (window as any).Razorpay !== "undefined") {\r\n' +
  '      const planToCharge = checkoutPlans.find(p => p.id === selectedPlanId) || checkoutPlans[0];\r\n' +
  '      const chargePrice = planToCharge ? planToCharge.price : 1599;\r\n' +
  '      const chargeName = planToCharge ? planToCharge.name : "SheetCodeCrest Pro";\r\n\r\n' +
  '      const options = {\r\n' +
  '        key: razorpayKey,\r\n' +
  '        amount: chargePrice * 100, // in paisa\r\n' +
  '        currency: "INR",\r\n' +
  '        name: chargeName,\r\n' +
  '        description: `${chargeName} Subscription`,\r\n' +
  '        image: "https://sheetcodecrest.vercel.app/logo.png",\r\n' +
  '        handler: async function (response: any) {\r\n' +
  '          try {\r\n' +
  '            addLog(`💳 Razorpay transaction completed! Payment ID: ${response.razorpay_payment_id}`, "success");\r\n' +
  '            \r\n' +
  '            if (currentUser) {\r\n' +
  '              const updatedUser = { ...currentUser, isPro: true };\r\n' +
  '              await dbSaveUser(updatedUser);\r\n' +
  '              setCurrentUser(updatedUser);\r\n' +
  '              \r\n' +
  '              // Log payment securely in Supabase\r\n' +
  '              const { dbLogPayment } = await import("./db");\r\n' +
  '              await dbLogPayment({\r\n' +
  '                username: currentUser.username,\r\n' +
  '                gateway: "razorpay",\r\n' +
  '                paymentId: response.razorpay_payment_id,\r\n' +
  '                orderId: response.razorpay_order_id || "",\r\n' +
  '                signature: response.razorpay_signature || "",\r\n' +
  '                amount: chargePrice,\r\n' +
  '                status: "success"\r\n' +
  '              });'
);

if (c.includes(oldSimStart)) {
  c = c.replace(oldSimStart, newSimStart);
  console.log('v Made startPaymentSimulation Razorpay params dynamic');
  changed++;
} else {
  console.log('x startPaymentSimulation old string not matched exactly');
}

// 2. handleManualUpiVerification
const oldManualUPI = crlf(
  '          // Log manual transaction request to Supabase payments as pending\r\n' +
  '          const { dbLogPayment } = await import("./db");\r\n' +
  '          await dbLogPayment({\r\n' +
  '            username: currentUser.username,\r\n' +
  '            gateway: "razorpay",\r\n' +
  '            paymentId: `upi_utr_pending_${upiUTR.trim()}`,\r\n' +
  '            amount: 1599,\r\n' +
  '            status: "pending_verification"\r\n' +
  '          });'
);

const newManualUPI = crlf(
  '          // Log manual transaction request to Supabase payments as pending\r\n' +
  '          const planToCharge = checkoutPlans.find(p => p.id === selectedPlanId) || checkoutPlans[0];\r\n' +
  '          const chargePrice = planToCharge ? planToCharge.price : 1599;\r\n\r\n' +
  '          const { dbLogPayment } = await import("./db");\r\n' +
  '          await dbLogPayment({\r\n' +
  '            username: currentUser.username,\r\n' +
  '            gateway: "razorpay",\r\n' +
  '            paymentId: `upi_utr_pending_${upiUTR.trim()}`,\r\n' +
  '            amount: chargePrice,\r\n' +
  '            status: "pending_verification"\r\n' +
  '          });'
);

if (c.includes(oldManualUPI)) {
  c = c.replace(oldManualUPI, newManualUPI);
  console.log('v Made handleManualUpiVerification Supabase amount dynamic');
  changed++;
} else {
  console.log('x handleManualUpiVerification old string not matched exactly');
}

// 3. Razorpay button label
const oldButton = crlf(
  '                      <button type="submit" className="auth-submit-btn">\r\n' +
  '                        🔒 Pay ₹1,599 Securely via Razorpay\r\n' +
  '                      </button>'
);

const newButton = crlf(
  '                      <button type="submit" className="auth-submit-btn">\r\n' +
  '                        🔒 Pay ₹{((checkoutPlans.find(p => p.id === selectedPlanId) || checkoutPlans[0])?.price || 1599).toLocaleString()} Securely via Razorpay\r\n' +
  '                      </button>'
);

if (c.includes(oldButton)) {
  c = c.replace(oldButton, newButton);
  console.log('v Made Razorpay button text dynamic');
  changed++;
} else {
  console.log('x Razorpay button text not matched exactly');
}

// 4. QR Code simulator elements
const oldQR = crlf(
  '                        <img \r\n' +
  '                          src={`https://api.qrserver.com/v1/create-qr-code/?size=180x180&data=${encodeURIComponent(`upi://pay?pa=${PERSONAL_UPI_ID}&pn=SheetCodeCrest&am=1599.00&cu=INR&tn=SheetCodeCrest%20Pro`)}`} \r\n' +
  '                          alt="UPI QR Code" \r\n' +
  '                          style={{ display: "block" }} \r\n' +
  '                        />\r\n' +
  '                      </div>\r\n' +
  '                      <div style={{ textAlign: "center" }}>\r\n' +
  '                        <div style={{ fontSize: "11px", color: "var(--slate)", textTransform: "uppercase", letterSpacing: "0.05em", fontWeight: 700 }}>Scan QR to Pay with GPay / Paytm / PhonePe</div>\r\n' +
  '                        <div style={{ fontSize: "13px", fontWeight: 600, color: "var(--coral)", marginTop: "4px", fontFamily: "var(--font-technical)" }}>₹1,599 (SheetCodeCrest Pro Lifetime)</div>'
);

const newQR = crlf(
  '                        <img \r\n' +
  '                          src={`https://api.qrserver.com/v1/create-qr-code/?size=180x180&data=${encodeURIComponent(`upi://pay?pa=${PERSONAL_UPI_ID}&pn=SheetCodeCrest&am=${((checkoutPlans.find(p => p.id === selectedPlanId) || checkoutPlans[0])?.price || 1599).toFixed(2)}&cu=INR&tn=${((checkoutPlans.find(p => p.id === selectedPlanId) || checkoutPlans[0])?.name || "SheetCodeCrest Pro")}`)}`} \r\n' +
  '                          alt="UPI QR Code" \r\n' +
  '                          style={{ display: "block" }} \r\n' +
  '                        />\r\n' +
  '                      </div>\r\n' +
  '                      <div style={{ textAlign: "center" }}>\r\n' +
  '                        <div style={{ fontSize: "11px", color: "var(--slate)", textTransform: "uppercase", letterSpacing: "0.05em", fontWeight: 700 }}>Scan QR to Pay with GPay / Paytm / PhonePe</div>\r\n' +
  '                        <div style={{ fontSize: "13px", fontWeight: 600, color: "var(--coral)", marginTop: "4px", fontFamily: "var(--font-technical)" }}>₹{((checkoutPlans.find(p => p.id === selectedPlanId) || checkoutPlans[0])?.price || 1599).toLocaleString()} ({((checkoutPlans.find(p => p.id === selectedPlanId) || checkoutPlans[0])?.name || "SheetCodeCrest Pro")})</div>'
);

if (c.includes(oldQR)) {
  c = c.replace(oldQR, newQR);
  console.log('v Made UPI QR simulator src and description dynamic');
  changed++;
} else {
  console.log('x UPI QR simulator elements not matched exactly');
}

fs.writeFileSync(file, c, 'utf-8');
console.log('Done! Changed elements:', changed, 'new size:', c.length);
