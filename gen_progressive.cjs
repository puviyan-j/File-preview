const XLSX = require('xlsx');
const wb = XLSX.utils.book_new();

// Big sheet (100,000 rows)
const data = [['Row Index', 'Random Value', 'Timestamp', 'Status']];
for (let i = 0; i < 100000; i++) {
    data.push([
        i + 1,
        Math.random().toFixed(4),
        Date.now(),
        ['OK', 'WARN', 'ERROR'][i % 3]
    ]);
}
const ws = XLSX.utils.aoa_to_sheet(data);
XLSX.utils.book_append_sheet(wb, ws, 'Big Data');

XLSX.writeFile(wb, 'progressive_test.xlsx');
console.log('Created progressive_test.xlsx with 100,000 rows');
