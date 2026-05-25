const fs = require('fs');
const path = require('path');
const XLSX = require('xlsx');

const root = path.resolve(__dirname, '..');
const buildDir = path.join(root, '.tmp');
const analyticsPath = path.join(buildDir, 'excelAnalytics.js');

if (!fs.existsSync(analyticsPath)) {
  console.error('Compiled excelAnalytics.js not found. Run `npm run validate:samples` again after compilation.');
  process.exit(1);
}

const analytics = require(analyticsPath);
const shiprocketDir = path.join(root, 'shiprocket');
let sampleFiles = [];
if (fs.existsSync(shiprocketDir)) {
  sampleFiles = fs.readdirSync(shiprocketDir)
    .filter(file => file.endsWith('.xlsx'))
    .map(file => path.join('shiprocket', file));
}

if (sampleFiles.length === 0) {
  sampleFiles = [
    'Order_Product_Analytics.xlsx',
    'Customer_Analytics_Retargeting.xlsx',
    'Product_Customer_Orders.xlsx',
  ];
}

if (!fs.existsSync(path.join(root, 'dist'))) {
  fs.mkdirSync(path.join(root, 'dist'), { recursive: true });
}

for (const relativePath of sampleFiles) {
  const filePath = path.join(root, relativePath);
  if (!fs.existsSync(filePath)) {
    console.warn(`Sample file missing: ${relativePath}`);
    continue;
  }

  const buffer = fs.readFileSync(filePath);
  const parsed = analytics.parseExcelBuffer(buffer);
  const profile = analytics.analyzeData(parsed.data, parsed.headers);
  const base = path.basename(relativePath, path.extname(relativePath));
  const wb = analytics.buildAnalyticsWorkbook(base, parsed.data, profile);
  const outPath = path.join(root, 'dist', `${base}_validation_report.xlsx`);
  XLSX.writeFile(wb, outPath);
  console.log(`✓ ${relativePath}: sheet="${parsed.sheetName}" rows=${parsed.data.length} cols=${parsed.headers.length} output=${outPath}`);
}
