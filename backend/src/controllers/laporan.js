import prisma from '../utils/prisma.js';
import { createRequire } from 'module';
const require = createRequire(import.meta.url);
const ExcelJS = require('exceljs');
import puppeteer from 'puppeteer';

async function getInitialBalance(periode) {
  const MONTH_NAMES = ['Januari', 'Februari', 'Maret', 'April', 'Mei', 'Juni', 'Juli', 'Agustus', 'September', 'Oktober', 'November', 'Desember'];
  function parsePeriode(periodeString) {
    const parts = periodeString.split(' ');
    if (parts.length === 2) {
      const monthIndex = MONTH_NAMES.findIndex(m => m.toLowerCase() === parts[0].toLowerCase());
      const year = parseInt(parts[1]);
      if (monthIndex !== -1 && !isNaN(year)) {
        return `${year}-${String(monthIndex + 1).padStart(2, '0')}`;
      }
    }
    return periodeString;
  }

  const allPemasukan = await prisma.pemasukan.findMany();
  const allPengeluaran = await prisma.pengeluaran.findMany();
  const allSaldoAwal = await prisma.saldoawal.findMany();

  const periodMap = new Map();

  allSaldoAwal.forEach(s => {
    const ym = parsePeriode(s.periode);
    periodMap.set(ym, { pemasukan: 0, pengeluaran: 0, saldoAwal: Number(s.saldo_awal) });
  });
  allPemasukan.forEach(p => {
    const ym = `${p.tanggal.getFullYear()}-${String(p.tanggal.getMonth() + 1).padStart(2, '0')}`;
    if (!periodMap.has(ym)) periodMap.set(ym, { pemasukan: 0, pengeluaran: 0, saldoAwal: null });
    periodMap.get(ym).pemasukan += Number(p.nominal);
  });
  allPengeluaran.forEach(p => {
    const ym = `${p.tanggal.getFullYear()}-${String(p.tanggal.getMonth() + 1).padStart(2, '0')}`;
    if (!periodMap.has(ym)) periodMap.set(ym, { pemasukan: 0, pengeluaran: 0, saldoAwal: null });
    periodMap.get(ym).pengeluaran += Number(p.nominal);
  });

  const sortedKeys = Array.from(periodMap.keys()).sort();

  if (!periode) {
    if (sortedKeys.length > 0 && periodMap.get(sortedKeys[0]).saldoAwal !== null) {
      return periodMap.get(sortedKeys[0]).saldoAwal;
    }
    return 0;
  }

  let currentSaldo = 0;
  let initialBalance = 0;

  for (const key of sortedKeys) {
    if (key === periode) {
      if (periodMap.get(key).saldoAwal !== null) {
        return periodMap.get(key).saldoAwal;
      } else {
        return currentSaldo;
      }
    }

    const data = periodMap.get(key);
    if (data.saldoAwal !== null) {
      currentSaldo = data.saldoAwal;
    }
    currentSaldo = currentSaldo + data.pemasukan - data.pengeluaran;
    initialBalance = currentSaldo;
  }

  return initialBalance;
}

export const exportLaporanKasExcel = async (req, res) => {
  try {
    const { periode } = req.query;

    let sheetName = 'Buku Kas RT';
    const localMonthNames = ['Januari', 'Februari', 'Maret', 'April', 'Mei', 'Juni', 'Juli', 'Agustus', 'September', 'Oktober', 'November', 'Desember'];
    let targetYear = -1, targetMonthIndex = -1;
    if (periode) {
      const [y, m] = periode.split('-');
      targetYear = parseInt(y);
      targetMonthIndex = parseInt(m) - 1;
      sheetName += ` - ${localMonthNames[targetMonthIndex]} ${targetYear}`;
    }

    const pemasukan = await prisma.pemasukan.findMany({
      include: { kategoripemasukan: true, warga: true },
      orderBy: { tanggal: 'asc' }
    });
    const pengeluaran = await prisma.pengeluaran.findMany({
      include: { kategoripengeluaran: true },
      orderBy: { tanggal: 'asc' }
    });

    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet(sheetName.substring(0, 31));

    // Kolom Header
    sheet.columns = [
      { header: 'No', key: 'no', width: 6 },
      { header: 'Tanggal', key: 'tanggal', width: 15 },
      { header: 'Kategori', key: 'kategori', width: 20 },
      { header: 'Uraian', key: 'uraian', width: 30 },
      { header: 'Penerima', key: 'penerima', width: 20 },
      { header: 'Pemasukan', key: 'pemasukan', width: 20 },
      { header: 'Pengeluaran', key: 'pengeluaran', width: 20 },
      { header: 'Saldo', key: 'saldo', width: 20 }
    ];

    // Style Header
    sheet.getRow(1).font = { bold: true };
    sheet.getRow(1).alignment = { horizontal: 'center', vertical: 'middle' };

    // Format Currency
    const currencyFormat = '"Rp" #,##0';
    sheet.getColumn('pemasukan').numFmt = currencyFormat;
    sheet.getColumn('pengeluaran').numFmt = currencyFormat;
    sheet.getColumn('saldo').numFmt = currencyFormat;

    // Data Gabungan
    const allTransactions = [
      ...pemasukan.map(p => ({
        type: 'in',
        tanggal: p.tanggal,
        kategori: p.kategoripemasukan?.nama || 'Pemasukan',
        uraian: p.keterangan || '-',
        penerima: p.warga?.nama_lengkap || '-',
        nominal: Number(p.nominal)
      })),
      ...pengeluaran.map(p => ({
        type: 'out',
        tanggal: p.tanggal,
        kategori: p.kategoripengeluaran?.nama || 'Pengeluaran',
        uraian: p.deskripsi || '-',
        penerima: p.penerima || '-',
        nominal: Number(p.nominal)
      }))
    ].sort((a, b) => a.tanggal.getTime() - b.tanggal.getTime());

    let initialBalance = await getInitialBalance(periode);
    let currentSaldo = initialBalance;
    let totalPemasukan = 0;
    let totalPengeluaran = 0;

    // Baris Saldo Awal
    const saldoAwalRow = sheet.addRow({
      no: '',
      tanggal: '',
      kategori: '',
      uraian: 'Saldo Awal',
      penerima: '',
      pemasukan: '',
      pengeluaran: '',
      saldo: currentSaldo
    });
    saldoAwalRow.getCell('uraian').font = { bold: true };

    // Baris Data Transaksi
    let displayedRowNumber = 1;
    allTransactions.forEach((t) => {
      const tYear = t.tanggal.getFullYear();
      const tMonth = t.tanggal.getMonth();
      const isInsidePeriod = !periode || (tYear === targetYear && tMonth === targetMonthIndex);

      if (isInsidePeriod) {
        const isMasuk = t.type === 'in';

        if (isMasuk) {
          currentSaldo += t.nominal;
          totalPemasukan += t.nominal;
        } else {
          currentSaldo -= t.nominal;
          totalPengeluaran += t.nominal;
        }

        sheet.addRow({
          no: displayedRowNumber++,
          tanggal: t.tanggal.toLocaleDateString('id-ID', { day: '2-digit', month: '2-digit', year: 'numeric' }),
          kategori: t.kategori,
          uraian: t.uraian,
          penerima: t.penerima,
          pemasukan: isMasuk ? t.nominal : '',
          pengeluaran: !isMasuk ? t.nominal : '',
          saldo: currentSaldo
        });
      }
    });

    // Baris Total
    const totalRow = sheet.addRow({
      no: '',
      tanggal: '',
      kategori: '',
      uraian: 'TOTAL',
      penerima: '',
      pemasukan: totalPemasukan,
      pengeluaran: totalPengeluaran,
      saldo: currentSaldo
    });
    totalRow.font = { bold: true };

    // Beri border untuk semua sel
    sheet.eachRow((row) => {
      row.eachCell((cell) => {
        cell.border = {
          top: { style: 'thin' },
          left: { style: 'thin' },
          bottom: { style: 'thin' },
          right: { style: 'thin' }
        };
      });
    });

    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', 'attachment; filename=' + 'laporan_kas.xlsx');

    await workbook.xlsx.write(res);
    res.end();
  } catch (error) {
    console.error(error);
    res.status(500).json({ success: false, message: 'Server error' });
  }
};

export const exportLaporanKasPDF = async (req, res) => {
  try {
    const { periode } = req.query;

    let headerText = 'Laporan Buku Kas Umum RT';
    const localMonthNames = ['Januari', 'Februari', 'Maret', 'April', 'Mei', 'Juni', 'Juli', 'Agustus', 'September', 'Oktober', 'November', 'Desember'];
    let targetYear = -1, targetMonthIndex = -1;
    if (periode) {
      const [y, m] = periode.split('-');
      targetYear = parseInt(y);
      targetMonthIndex = parseInt(m) - 1;
      headerText += ` - Bulan ${localMonthNames[targetMonthIndex]} ${targetYear}`;
    }
    // Ganti pemanggilan puppeteer.launch Anda menjadi seperti ini:
    const browser = await puppeteer.launch({
      headless: true, // atau "new"
      args: [
        '--no-sandbox',
        '--disable-setuid-sandbox',
        '--disable-dev-shm-usage'
      ]
    });
    // --- FETCH DATA FIRST ---
    const pemasukan = await prisma.pemasukan.findMany({
      include: { kategoripemasukan: true, warga: true },
      orderBy: { tanggal: 'asc' }
    });
    const pengeluaran = await prisma.pengeluaran.findMany({
      include: { kategoripengeluaran: true },
      orderBy: { tanggal: 'asc' }
    });

    const allTransactions = [
      ...pemasukan.map(p => ({
        type: 'in',
        tanggal: p.tanggal,
        kategori: p.kategoripemasukan?.nama || 'Pemasukan',
        uraian: p.keterangan || '-',
        penerima: p.warga?.nama_lengkap || '-',
        nominal: Number(p.nominal)
      })),
      ...pengeluaran.map(p => ({
        type: 'out',
        tanggal: p.tanggal,
        kategori: p.kategoripengeluaran?.nama || 'Pengeluaran',
        uraian: p.deskripsi || '-',
        penerima: p.penerima || '-',
        nominal: Number(p.nominal)
      }))
    ].sort((a, b) => a.tanggal.getTime() - b.tanggal.getTime());

    const formatRupiah = (angka) => {
      return new Intl.NumberFormat('id-ID', { style: 'currency', currency: 'IDR', maximumFractionDigits: 0 }).format(angka);
    };

    let initialBalance = await getInitialBalance(periode);
    let currentSaldo = initialBalance;
    let totalPemasukan = 0;
    let totalPengeluaran = 0;

    const tableData = [];

    // Saldo Awal Row
    tableData.push({
      no: '-',
      tanggal: '-',
      kategori: '-',
      uraian: 'Saldo Awal',
      penerima: '-',
      pemasukan: '-',
      pengeluaran: '-',
      saldo: formatRupiah(currentSaldo)
    });

    let displayedRowNumber = 1;
    allTransactions.forEach((t) => {
      const tYear = t.tanggal.getFullYear();
      const tMonth = t.tanggal.getMonth();
      const isInsidePeriod = !periode || (tYear === targetYear && tMonth === targetMonthIndex);

      if (isInsidePeriod) {
        const isMasuk = t.type === 'in';

        if (isMasuk) {
          currentSaldo += t.nominal;
          totalPemasukan += t.nominal;
        } else {
          currentSaldo -= t.nominal;
          totalPengeluaran += t.nominal;
        }

        tableData.push({
          no: (displayedRowNumber++).toString(),
          tanggal: t.tanggal.toLocaleDateString('id-ID', { day: '2-digit', month: '2-digit', year: 'numeric' }),
          kategori: String(t.kategori || '-'),
          uraian: String(t.uraian || '-'),
          penerima: String(t.penerima || '-'),
          pemasukan: isMasuk ? formatRupiah(t.nominal) : '-',
          pengeluaran: !isMasuk ? formatRupiah(t.nominal) : '-',
          saldo: formatRupiah(currentSaldo)
        });
      }
    });

    // Total Row
    tableData.push({
      no: '-',
      tanggal: '-',
      kategori: '-',
      uraian: 'TOTAL',
      penerima: '-',
      pemasukan: formatRupiah(totalPemasukan),
      pengeluaran: formatRupiah(totalPengeluaran),
      saldo: formatRupiah(currentSaldo)
    });

    // --- GENERATE HTML FOR PUPPETEER ---
    let tableHtml = `
      <table style="width: 100%; border-collapse: collapse; font-family: Helvetica, Arial, sans-serif; font-size: 10px;">
        <thead>
          <tr style="background-color: #f2f2f2; font-weight: bold; text-align: left;">
            <th style="border: 1px solid #ddd; padding: 6px;">No</th>
            <th style="border: 1px solid #ddd; padding: 6px;">Tanggal</th>
            <th style="border: 1px solid #ddd; padding: 6px;">Kategori</th>
            <th style="border: 1px solid #ddd; padding: 6px;">Uraian</th>
            <th style="border: 1px solid #ddd; padding: 6px;">Penerima</th>
            <th style="border: 1px solid #ddd; padding: 6px;">Pemasukan</th>
            <th style="border: 1px solid #ddd; padding: 6px;">Pengeluaran</th>
            <th style="border: 1px solid #ddd; padding: 6px;">Saldo</th>
          </tr>
        </thead>
        <tbody>
    `;

    tableData.forEach(row => {
      tableHtml += `
        <tr>
          <td style="border: 1px solid #ddd; padding: 6px;">${row.no}</td>
          <td style="border: 1px solid #ddd; padding: 6px;">${row.tanggal}</td>
          <td style="border: 1px solid #ddd; padding: 6px;">${row.kategori}</td>
          <td style="border: 1px solid #ddd; padding: 6px;">${row.uraian}</td>
          <td style="border: 1px solid #ddd; padding: 6px;">${row.penerima}</td>
          <td style="border: 1px solid #ddd; padding: 6px;">${row.pemasukan}</td>
          <td style="border: 1px solid #ddd; padding: 6px;">${row.pengeluaran}</td>
          <td style="border: 1px solid #ddd; padding: 6px;">${row.saldo}</td>
        </tr>
      `;
    });

    tableHtml += `
        </tbody>
      </table>
    `;

    const htmlContent = `
      <html>
        <head>
          <style>
            body { font-family: Helvetica, Arial, sans-serif; margin: 30px; }
            h2 { text-align: center; font-size: 16px; margin-bottom: 20px; }
          </style>
        </head>
        <body>
          <h2>${headerText}</h2>
          ${tableHtml}
        </body>
      </html>
    `;

    const page = await browser.newPage();
    await page.setContent(htmlContent, { waitUntil: 'networkidle0' });

    const pdfBuffer = await page.pdf({ format: 'A4', printBackground: true });

    await browser.close();

    // === INI ADALAH KUNCI AGAR LAYAR TIDAK PUTIH ===
    // Memberitahu browser bahwa ini adalah dokumen PDF
    res.setHeader('Content-Type', 'application/pdf');

    // 'inline' berarti langsung dibuka di tab browser. 
    // Jika ingin otomatis terdownload, ganti 'inline' menjadi 'attachment'
    res.setHeader('Content-Disposition', 'inline; filename="Laporan-Keuangan.pdf"');

    // Mengirimkan file ke browser
    res.send(pdfBuffer);
  } catch (error) {
    console.error("PDF Generation Error:", error);
    if (!res.headersSent) {
      res.status(500).json({ success: false, message: 'Server error' });
    } else {
      res.end();
    }
  }
};

const MONTH_NAMES = ['Januari', 'Februari', 'Maret', 'April', 'Mei', 'Juni', 'Juli', 'Agustus', 'September', 'Oktober', 'November', 'Desember'];

function parsePeriode(periodeString) {
  const parts = periodeString.split(' ');
  if (parts.length === 2) {
    const monthIndex = MONTH_NAMES.findIndex(m => m.toLowerCase() === parts[0].toLowerCase());
    const year = parseInt(parts[1]);
    if (monthIndex !== -1 && !isNaN(year)) {
      return `${year}-${String(monthIndex + 1).padStart(2, '0')}`;
    }
  }
  return periodeString;
}

function formatPeriode(yearMonth) {
  const [year, month] = yearMonth.split('-');
  return `${MONTH_NAMES[parseInt(month) - 1]} ${year}`;
}

export const getLaporanKumulatif = async (req, res) => {
  try {
    const pemasukan = await prisma.pemasukan.findMany();
    const pengeluaran = await prisma.pengeluaran.findMany();
    const saldoAwal = await prisma.saldoawal.findMany();

    const periodMap = new Map();

    saldoAwal.forEach(s => {
      const ym = parsePeriode(s.periode);
      periodMap.set(ym, { pemasukan: 0, pengeluaran: 0, saldoAwal: Number(s.saldo_awal), label: s.periode });
    });

    pemasukan.forEach(p => {
      const ym = `${p.tanggal.getFullYear()}-${String(p.tanggal.getMonth() + 1).padStart(2, '0')}`;
      if (!periodMap.has(ym)) periodMap.set(ym, { pemasukan: 0, pengeluaran: 0, saldoAwal: null, label: formatPeriode(ym) });
      periodMap.get(ym).pemasukan += Number(p.nominal);
    });

    pengeluaran.forEach(p => {
      const ym = `${p.tanggal.getFullYear()}-${String(p.tanggal.getMonth() + 1).padStart(2, '0')}`;
      if (!periodMap.has(ym)) periodMap.set(ym, { pemasukan: 0, pengeluaran: 0, saldoAwal: null, label: formatPeriode(ym) });
      periodMap.get(ym).pengeluaran += Number(p.nominal);
    });

    const sortedKeys = Array.from(periodMap.keys()).sort();
    const kumulatifData = [];
    let currentSaldo = 0;

    for (const key of sortedKeys) {
      const data = periodMap.get(key);
      if (data.saldoAwal !== null) {
        currentSaldo = data.saldoAwal;
      }

      const akhir = currentSaldo + data.pemasukan - data.pengeluaran;
      kumulatifData.push({
        key,
        periode: data.label,
        pemasukan: data.pemasukan,
        pengeluaran: data.pengeluaran,
        saldo_akhir: akhir
      });
      currentSaldo = akhir;
    }

    res.json({ success: true, data: kumulatifData });
  } catch (error) {
    console.error(error);
    res.status(500).json({ success: false, message: 'Server error' });
  }
};

export const getLaporanBulanan = async (req, res) => {
  try {
    const { periode } = req.query; // e.g. "2026-01"
    if (!periode || typeof periode !== 'string') return res.status(400).json({ success: false, message: 'Periode is required' });

    const [year, month] = periode.split('-');
    const startDate = new Date(parseInt(year), parseInt(month) - 1, 1);
    const endDate = new Date(parseInt(year), parseInt(month), 0, 23, 59, 59);

    const pemasukan = await prisma.pemasukan.findMany({
      where: { tanggal: { gte: startDate, lte: endDate } },
      include: { kategoripemasukan: true },
      orderBy: { tanggal: 'asc' }
    });

    const pengeluaran = await prisma.pengeluaran.findMany({
      where: { tanggal: { gte: startDate, lte: endDate } },
      include: { kategoripengeluaran: true },
      orderBy: { tanggal: 'asc' }
    });

    const allPemasukan = await prisma.pemasukan.findMany();
    const allPengeluaran = await prisma.pengeluaran.findMany();
    const allSaldoAwal = await prisma.saldoawal.findMany();

    const periodMap = new Map();
    allSaldoAwal.forEach(s => {
      const ym = parsePeriode(s.periode);
      periodMap.set(ym, { pemasukan: 0, pengeluaran: 0, saldoAwal: Number(s.saldo_awal) });
    });
    allPemasukan.forEach(p => {
      const ym = `${p.tanggal.getFullYear()}-${String(p.tanggal.getMonth() + 1).padStart(2, '0')}`;
      if (!periodMap.has(ym)) periodMap.set(ym, { pemasukan: 0, pengeluaran: 0, saldoAwal: null });
      periodMap.get(ym).pemasukan += Number(p.nominal);
    });
    allPengeluaran.forEach(p => {
      const ym = `${p.tanggal.getFullYear()}-${String(p.tanggal.getMonth() + 1).padStart(2, '0')}`;
      if (!periodMap.has(ym)) periodMap.set(ym, { pemasukan: 0, pengeluaran: 0, saldoAwal: null });
      periodMap.get(ym).pengeluaran += Number(p.nominal);
    });

    const sortedKeys = Array.from(periodMap.keys()).sort();
    let currentSaldo = 0;
    let initialBalance = 0;

    for (const key of sortedKeys) {
      if (key === periode) {
        if (periodMap.get(key).saldoAwal !== null) {
          initialBalance = periodMap.get(key).saldoAwal;
        } else {
          initialBalance = currentSaldo;
        }
        break;
      }

      const data = periodMap.get(key);
      if (data.saldoAwal !== null) {
        currentSaldo = data.saldoAwal;
      }
      currentSaldo = currentSaldo + data.pemasukan - data.pengeluaran;
      initialBalance = currentSaldo;
    }

    const allTransactions = [
      ...pemasukan.map(p => ({
        id: p.id,
        type: 'in',
        tanggal: p.tanggal,
        kategori: p.kategoripemasukan?.nama || 'Pemasukan',
        uraian: p.keterangan || '-',
        nominal: Number(p.nominal)
      })),
      ...pengeluaran.map(p => ({
        id: p.id,
        type: 'out',
        tanggal: p.tanggal,
        kategori: p.kategoripengeluaran?.nama || 'Pengeluaran',
        uraian: p.deskripsi || '-',
        nominal: Number(p.nominal)
      }))
    ].sort((a, b) => a.tanggal.getTime() - b.tanggal.getTime());

    const recapPemasukan = {};
    pemasukan.forEach(p => {
      const cat = p.kategoripemasukan?.nama || 'Pemasukan';
      recapPemasukan[cat] = (recapPemasukan[cat] || 0) + Number(p.nominal);
    });
    const recapPengeluaran = {};
    pengeluaran.forEach(p => {
      const cat = p.kategoripengeluaran?.nama || 'Pengeluaran';
      recapPengeluaran[cat] = (recapPengeluaran[cat] || 0) + Number(p.nominal);
    });

    res.json({
      success: true,
      data: {
        saldoAwal: initialBalance,
        transactions: allTransactions,
        recap: {
          pemasukan: Object.keys(recapPemasukan).map(k => ({ kategori: k, total: recapPemasukan[k] })),
          pengeluaran: Object.keys(recapPengeluaran).map(k => ({ kategori: k, total: recapPengeluaran[k] }))
        }
      }
    });
  } catch (error) {
    console.error(error);
    res.status(500).json({ success: false, message: 'Server error' });
  }
};
