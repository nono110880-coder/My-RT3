<script setup>
import { ref, onMounted, watch } from 'vue';
import { Download, Calendar, BarChart3, TrendingUp, TrendingDown } from '@lucide/vue';
import api, { getBaseUrl } from '../services/api.js';
import { useAuthStore } from '../store/authStore.js';

const activeTab = ref('kumulatif');
const loading = ref(false);

const kumulatifData = ref([]);
const bulananData = ref(null);

const today = new Date();
const defaultPeriod = `${today.getFullYear()}-${String(today.getMonth() + 1).padStart(2, '0')}`;
const selectedPeriod = ref(defaultPeriod);

const fetchKumulatif = async () => {
  loading.value = true;
  try {
    const response = await api.get('/laporan/kumulatif');
    if (response.data.success) {
      kumulatifData.value = response.data.data;
    }
  } catch (error) {
    console.error("Gagal mengambil data kumulatif", error);
  } finally {
    loading.value = false;
  }
};

const fetchBulanan = async () => {
  loading.value = true;
  try {
    const response = await api.get(`/laporan/bulanan?periode=${selectedPeriod.value}`);
    if (response.data.success) {
      bulananData.value = response.data.data;
    }
  } catch (error) {
    console.error("Gagal mengambil data bulanan", error);
  } finally {
    loading.value = false;
  }
};

watch([activeTab, selectedPeriod], () => {
  if (activeTab.value === 'kumulatif') {
    fetchKumulatif();
  } else {
    fetchBulanan();
  }
});

onMounted(() => {
  if (activeTab.value === 'kumulatif') {
    fetchKumulatif();
  } else {
    fetchBulanan();
  }
});

const handleExport = (format) => {
  try {
    const params = new URLSearchParams();
    if (activeTab.value === 'bulanan') {
      params.append('periode', selectedPeriod.value);
    }
    
    // Cache buster untuk menghindari cache Service Worker / Browser
    params.append('_cb', Date.now());
    
    const url = `${getBaseUrl()}/laporan/${format}?${params.toString()}`;

    if (format === 'pdf') {
      window.open(url, '_blank');
    } else {
      const link = document.createElement('a');
      link.href = url;
      link.setAttribute('download', 'Laporan_Keuangan.xlsx');
      link.setAttribute('target', '_blank');
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
    }
  } catch (error) {
    console.error(`Gagal mengunduh laporan ${format}`, error);
    alert("Terjadi kesalahan saat mengunduh laporan.");
  }
};

const formatRupiah = (angka) => {
  return new Intl.NumberFormat('id-ID', { style: 'currency', currency: 'IDR', maximumFractionDigits: 0 }).format(angka);
};

const formatDate = (dateString) => {
  return new Date(dateString).toLocaleDateString('id-ID', { day: 'numeric', month: 'long', year: 'numeric' });
};

const generateMonthOptions = () => {
  const options = [];
  const date = new Date();
  for (let i = 0; i < 24; i++) {
    const year = date.getFullYear();
    const month = String(date.getMonth() + 1).padStart(2, '0');
    const label = date.toLocaleDateString('id-ID', { month: 'long', year: 'numeric' });
    options.push({ value: `${year}-${month}`, label });
    date.setMonth(date.getMonth() - 1);
  }
  return options;
};

const monthOptions = generateMonthOptions();
</script>

<template>
  <div class="space-y-6 animate-in fade-in duration-500">
    <div class="flex flex-col sm:flex-row justify-between items-start sm:items-end gap-4">
      <div>
        <h1 class="text-2xl font-bold text-slate-800">Laporan Keuangan</h1>
        <p class="text-slate-500 text-sm mt-1">Pantau arus kas dan rekapitulasi keuangan RT</p>
      </div>
      <div class="flex space-x-2">
        <button @click="handleExport('excel')" class="flex items-center px-4 py-2 bg-green-600 text-white rounded-lg hover:bg-green-700 transition-colors text-sm font-medium shadow-sm">
          <Download :size="16" class="mr-2" />
          Excel
        </button>
        <button @click="handleExport('pdf')" class="flex items-center px-4 py-2 bg-red-500 text-white rounded-lg hover:bg-red-600 transition-colors text-sm font-medium shadow-sm">
          <Download :size="16" class="mr-2" />
          PDF
        </button>
      </div>
    </div>

    <!-- Tabs -->
    <div class="flex space-x-1 bg-slate-100 p-1 rounded-xl max-w-md">
      <button
        @click="activeTab = 'kumulatif'"
        :class="['flex-1 py-2 px-4 text-sm font-medium rounded-lg transition-colors', activeTab === 'kumulatif' ? 'bg-white text-primary-700 shadow-sm' : 'text-slate-500 hover:text-slate-700 hover:bg-slate-200']"
      >
        <BarChart3 :size="16" class="inline mr-2" />
        Laporan Keseluruhan
      </button>
      <button
        @click="activeTab = 'bulanan'"
        :class="['flex-1 py-2 px-4 text-sm font-medium rounded-lg transition-colors', activeTab === 'bulanan' ? 'bg-white text-primary-700 shadow-sm' : 'text-slate-500 hover:text-slate-700 hover:bg-slate-200']"
      >
        <Calendar :size="16" class="inline mr-2" />
        Laporan per Bulan
      </button>
    </div>

    <!-- Tab Content: Kumulatif -->
    <div v-if="activeTab === 'kumulatif'" class="bg-white rounded-xl shadow-sm border border-slate-200 overflow-hidden">
      <div class="p-5 border-b border-slate-200 bg-slate-50">
        <h2 class="text-lg font-semibold text-slate-800">Rekapitulasi Arus Kas Kumulatif</h2>
        <p class="text-sm text-slate-500">Menampilkan mutasi kas berdasarkan bulan</p>
      </div>
      <div class="overflow-x-auto">
        <table class="w-full text-left border-collapse">
          <thead>
            <tr class="bg-slate-50 border-b border-slate-200 text-slate-600 text-xs uppercase tracking-wider">
              <th class="px-6 py-4 font-medium">Periode</th>
              <th class="px-6 py-4 font-medium text-right text-green-600">Pemasukan</th>
              <th class="px-6 py-4 font-medium text-right text-red-500">Pengeluaran</th>
              <th class="px-6 py-4 font-medium text-right">Saldo Akhir</th>
            </tr>
          </thead>
          <tbody class="divide-y divide-slate-200 text-sm">
            <tr v-if="loading">
              <td colspan="4" class="px-6 py-10 text-center text-slate-500">Memuat data...</td>
            </tr>
            <tr v-else-if="kumulatifData.length === 0">
              <td colspan="4" class="px-6 py-10 text-center text-slate-500">Belum ada data transaksi.</td>
            </tr>
            <tr v-else v-for="item in kumulatifData" :key="item.key" class="hover:bg-slate-50 transition-colors">
              <td class="px-6 py-4 font-medium text-slate-700">{{ item.periode }}</td>
              <td class="px-6 py-4 text-right text-green-600">{{ formatRupiah(item.pemasukan) }}</td>
              <td class="px-6 py-4 text-right text-red-500">{{ formatRupiah(item.pengeluaran) }}</td>
              <td class="px-6 py-4 text-right font-bold text-slate-800">{{ formatRupiah(item.saldo_akhir) }}</td>
            </tr>
          </tbody>
        </table>
      </div>
    </div>

    <!-- Tab Content: Bulanan -->
    <div v-if="activeTab === 'bulanan'" class="space-y-6">
      <div class="flex items-center space-x-3 bg-white p-4 rounded-xl shadow-sm border border-slate-200">
        <label class="text-sm font-medium text-slate-700">Pilih Periode Laporan:</label>
        <select
          v-model="selectedPeriod"
          class="px-4 py-2 border border-slate-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-primary-500 text-sm min-w-[200px]"
        >
          <option v-for="opt in monthOptions" :key="opt.value" :value="opt.value">{{ opt.label }}</option>
        </select>
      </div>

      <div v-if="loading" class="bg-white rounded-xl shadow-sm border border-slate-200 p-10 text-center text-slate-500">
        Memuat data laporan bulan ini...
      </div>
      
      <template v-else-if="bulananData">
        <!-- Summary Cards -->
        <div class="grid grid-cols-1 md:grid-cols-3 gap-4">
          <div class="bg-white rounded-xl shadow-sm border border-slate-200 p-5">
            <div class="flex items-center text-slate-500 mb-2">
              <BarChart3 :size="18" class="mr-2" />
              <h3 class="text-sm font-medium">Saldo Awal Bulan</h3>
            </div>
            <p class="text-2xl font-bold text-slate-800">{{ formatRupiah(bulananData.saldoAwal) }}</p>
          </div>
          <div class="bg-white rounded-xl shadow-sm border border-slate-200 p-5">
            <div class="flex items-center text-green-600 mb-2">
              <TrendingUp :size="18" class="mr-2" />
              <h3 class="text-sm font-medium text-slate-600">Total Pemasukan</h3>
            </div>
            <p class="text-2xl font-bold text-green-600">
              {{ formatRupiah(bulananData.recap.pemasukan.reduce((acc, cur) => acc + cur.total, 0)) }}
            </p>
          </div>
          <div class="bg-white rounded-xl shadow-sm border border-slate-200 p-5">
            <div class="flex items-center text-red-500 mb-2">
              <TrendingDown :size="18" class="mr-2" />
              <h3 class="text-sm font-medium text-slate-600">Total Pengeluaran</h3>
            </div>
            <p class="text-2xl font-bold text-red-500">
              {{ formatRupiah(bulananData.recap.pengeluaran.reduce((acc, cur) => acc + cur.total, 0)) }}
            </p>
          </div>
        </div>

        <!-- Transactions Table -->
        <div class="bg-white rounded-xl shadow-sm border border-slate-200 overflow-hidden">
          <div class="p-4 border-b border-slate-200 bg-slate-50">
            <h3 class="font-semibold text-slate-800">Detail Transaksi Bulan Ini</h3>
          </div>
          <div class="overflow-x-auto">
            <table class="w-full text-left border-collapse">
              <thead>
                <tr class="bg-slate-50 border-b border-slate-200 text-slate-600 text-xs uppercase tracking-wider">
                  <th class="px-6 py-4 font-medium">Tanggal</th>
                  <th class="px-6 py-4 font-medium">Kategori</th>
                  <th class="px-6 py-4 font-medium">Uraian</th>
                  <th class="px-6 py-4 font-medium text-right">Pemasukan</th>
                  <th class="px-6 py-4 font-medium text-right">Pengeluaran</th>
                </tr>
              </thead>
              <tbody class="divide-y divide-slate-200 text-sm">
                <tr v-if="bulananData.transactions.length === 0">
                  <td colspan="5" class="px-6 py-8 text-center text-slate-500">Tidak ada transaksi pada bulan ini.</td>
                </tr>
                <tr v-else v-for="t in bulananData.transactions" :key="t.id" class="hover:bg-slate-50 transition-colors">
                  <td class="px-6 py-4 text-slate-600">{{ formatDate(t.tanggal) }}</td>
                  <td class="px-6 py-4">
                    <span :class="['inline-flex items-center px-2.5 py-1 rounded-full text-xs font-medium', t.type === 'in' ? 'bg-green-50 text-green-700' : 'bg-red-50 text-red-700']">
                      {{ t.kategori }}
                    </span>
                  </td>
                  <td class="px-6 py-4 text-slate-700">{{ t.uraian }}</td>
                  <td class="px-6 py-4 text-right text-green-600">
                    {{ t.type === 'in' ? formatRupiah(t.nominal) : '-' }}
                  </td>
                  <td class="px-6 py-4 text-right text-red-500">
                    {{ t.type === 'out' ? formatRupiah(t.nominal) : '-' }}
                  </td>
                </tr>
                <tr class="bg-slate-50 border-t-2 border-slate-200 font-bold">
                  <td colspan="3" class="px-6 py-4 text-right text-slate-800">SALDO AKHIR BULAN</td>
                  <td colspan="2" class="px-6 py-4 text-right text-primary-700 text-lg">
                    {{ formatRupiah(
                      bulananData.saldoAwal + 
                      bulananData.recap.pemasukan.reduce((acc, cur) => acc + cur.total, 0) - 
                      bulananData.recap.pengeluaran.reduce((acc, cur) => acc + cur.total, 0)
                    ) }}
                  </td>
                </tr>
              </tbody>
            </table>
          </div>
        </div>

        <!-- Recap Grid -->
        <div class="grid grid-cols-1 md:grid-cols-2 gap-6">
          <div class="bg-white rounded-xl shadow-sm border border-slate-200 overflow-hidden">
            <div class="p-4 border-b border-slate-200 bg-slate-50">
              <h3 class="font-semibold text-slate-800">Rekap Pemasukan per Kategori</h3>
            </div>
            <table class="w-full text-left text-sm">
              <tbody class="divide-y divide-slate-200">
                <tr v-if="bulananData.recap.pemasukan.length === 0"><td class="px-6 py-4 text-slate-500 text-center">Belum ada data</td></tr>
                <tr v-else v-for="r in bulananData.recap.pemasukan" :key="r.kategori">
                  <td class="px-6 py-3 text-slate-600">{{ r.kategori }}</td>
                  <td class="px-6 py-3 text-right font-medium text-green-600">{{ formatRupiah(r.total) }}</td>
                </tr>
              </tbody>
            </table>
          </div>
          
          <div class="bg-white rounded-xl shadow-sm border border-slate-200 overflow-hidden">
            <div class="p-4 border-b border-slate-200 bg-slate-50">
              <h3 class="font-semibold text-slate-800">Rekap Pengeluaran per Kategori</h3>
            </div>
            <table class="w-full text-left text-sm">
              <tbody class="divide-y divide-slate-200">
                <tr v-if="bulananData.recap.pengeluaran.length === 0"><td class="px-6 py-4 text-slate-500 text-center">Belum ada data</td></tr>
                <tr v-else v-for="r in bulananData.recap.pengeluaran" :key="r.kategori">
                  <td class="px-6 py-3 text-slate-600">{{ r.kategori }}</td>
                  <td class="px-6 py-3 text-right font-medium text-red-500">{{ formatRupiah(r.total) }}</td>
                </tr>
              </tbody>
            </table>
          </div>
        </div>
      </template>
    </div>
  </div>
</template>
