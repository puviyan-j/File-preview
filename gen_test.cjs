const XLSX = require('xlsx');
const wb = XLSX.utils.book_new();

const data = [['Name', 'Age', 'City', 'Score', 'Department']];
const cities = ['NYC', 'LA', 'SF', 'Austin', 'Portland', 'Chicago', 'Denver', 'Seattle'];
const depts = ['Engineering', 'Marketing', 'Sales', 'HR', 'Finance'];
for (let i = 0; i < 1000; i++) {
    data.push([
        'Employee_' + (i + 1),
        20 + Math.floor(Math.random() * 40),
        cities[i % cities.length],
        Math.floor(Math.random() * 100),
        depts[i % depts.length]
    ]);
}
const ws = XLSX.utils.aoa_to_sheet(data);
XLSX.utils.book_append_sheet(wb, ws, 'Employees');

const data2 = [
    ['Product', 'Price', 'Quantity', 'Category'],
    ['Widget A', 9.99, 100, 'Hardware'],
    ['Gadget B', 19.99, 50, 'Electronics'],
    ['Doohickey C', 4.99, 500, 'Accessories'],
    ['Thingamajig D', 29.99, 25, 'Premium'],
    ['Whatchamacallit E', 14.99, 200, 'General'],
];
const ws2 = XLSX.utils.aoa_to_sheet(data2);
XLSX.utils.book_append_sheet(wb, ws2, 'Products');

const data3 = [['X', 'Y', 'Z']];
for (let i = 0; i < 500; i++) {
    data3.push([Math.random() * 1000, Math.random() * 1000, Math.random() * 1000]);
}
const ws3 = XLSX.utils.aoa_to_sheet(data3);
XLSX.utils.book_append_sheet(wb, ws3, 'Numeric Data');

XLSX.writeFile(wb, 'test_data.xlsx');
console.log('Created test_data.xlsx');
