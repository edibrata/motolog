import * as XLSX from 'xlsx';
import { ServiceRecord, Vehicle } from './types';

export const CLOUD_DATABASE_URL = 'https://script.google.com/macros/s/AKfycbzSSxKoPOSQjvKRtOwIMzgmb4sjRWEBYo92kebLaLobLtvyZCx0kdRhwaZfZtl2fczU/exec';

export const INTEL_CONFIG = [
    { id: 'engine-oil', name: 'Ganti Oli Mesin', interval: 2500, kws: ['oli mesin', 'mesin'], color: 'blue' },
    { id: 'gear-oil', name: 'Ganti Oli Gardan', interval: 5000, kws: ['oli gardan', 'gardan'], color: 'indigo' },
    { id: 'cvt', name: 'Servis CVT', interval: 10000, kws: ['servis cvt', 'cvt'], color: 'blue' },
    { id: 'v-belt', name: 'Ganti V-Belt', interval: 25000, kws: ['v-belt', 'vbelt', 'van belt'], color: 'indigo' },
    { id: 'brake-rear', name: 'Kampas Rem Belakang', interval: 15000, kws: ['kampas rem belakang', 'rem belakang'], color: 'blue' },
    { id: 'brake-front', name: 'Kampas Rem Depan', interval: 17000, kws: ['kampas rem depan', 'rem depan'], color: 'indigo' },
    { id: 'ban-front', name: 'Ganti Ban Depan', interval: 15000, kws: ['ban depan', 'ban luar depan'], color: 'blue' },
    { id: 'ban-rear', name: 'Ganti Ban Belakang', interval: 12000, kws: ['ban belakang', 'ban luar belakang'], color: 'indigo' },
    { id: 'spark-plug', name: 'Ganti Busi', interval: 15000, kws: ['busi', 'spark plug'], color: 'blue' },
    { id: 'air-filter', name: 'Ganti Filter Udara', interval: 15000, kws: ['filter udara', 'air filter'], color: 'indigo' }
];

export function formatCurrency(n: number) {
  return new Intl.NumberFormat('id-ID').format(n);
}

export function formatDate(d: string | Date | number) {
  return new Date(d).toLocaleDateString('id-ID', { year: 'numeric', month: 'short', day: 'numeric' });
}

export function exportToExcel(records: ServiceRecord[], vehicles: Vehicle[], showToast: (msg: string, type?: 'success' | 'error') => void) {
    if (!records.length) return showToast("Tidak ada data", "error");
    const workbook = XLSX.utils.book_new();

    vehicles.forEach(vehicle => {
        const curRecords = records.filter(r => r.vehicleId === vehicle.id);
        if (curRecords.length > 0) {
            const flatData = curRecords.sort((a, b) => new Date(b.date).getTime() - new Date(a.date).getTime()).flatMap(record => 
                record.items.map(item => ({ 
                    "Tanggal": record.date, 
                    "Odometer": parseInt(record.odometer.toString()), 
                    "Bengkel": record.workshop || "-", 
                    "Item": item.name, 
                    "Harga": item.price, 
                    "Total": record.totalCost 
                }))
            );
            XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet(flatData), vehicle.name.substring(0, 31));
        }
    });

    const now = new Date();
    const ts = `${now.getFullYear()}${String(now.getMonth()+1).padStart(2,'0')}${String(now.getDate()).padStart(2,'0')} ${String(now.getHours()).padStart(2,'0')}.${String(now.getSeconds()).padStart(2,'0')}`;
    XLSX.writeFile(workbook, `Edi Brata Moto Journal ${ts}.xlsx`);
    showToast("Excel diunduh");
}

export function backupJSON(records: ServiceRecord[], vehicles: Vehicle[], showToast: (msg: string) => void) {
    const data = { vehicles, records, timestamp: new Date().toISOString() };
    const blob = new Blob([JSON.stringify(data, null, 2)], { type: 'application/json' });
    const now = new Date();
    const ts = `${now.getFullYear()}${String(now.getMonth()+1).padStart(2,'0')}${String(now.getDate()).padStart(2,'0')} ${String(now.getHours()).padStart(2,'0')}.${String(now.getSeconds()).padStart(2,'0')}`;
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = `Edi Brata Moto Journal ${ts}.json`;
    a.click();
    showToast("Backup JSON berhasil");
}
