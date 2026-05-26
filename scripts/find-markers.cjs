const fs = require('fs');
const c = fs.readFileSync('src/App.tsx', 'utf8');

// Find the Plans tab marker
const plansStart = 'TAB: PLANS';
const idx1 = c.indexOf(plansStart);
console.log('TAB: PLANS at:', idx1);
if (idx1 >= 0) {
  console.log('Context:', JSON.stringify(c.substring(idx1 - 10, idx1 + 60)));
}

// Find the payment summary box in JSX (not CSS)
let searchFrom = 0;
while (true) {
  const i = c.indexOf('payment-summary-box', searchFrom);
  if (i < 0) break;
  console.log('payment-summary-box at', i, ':', JSON.stringify(c.substring(i - 5, i + 80)));
  searchFrom = i + 1;
}
