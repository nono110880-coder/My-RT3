import { Router } from 'express';
import { exportLaporanKasExcel, exportLaporanKasPDF, getLaporanKumulatif, getLaporanBulanan } from '../controllers/laporan.js';

const router = Router();

router.get('/excel', exportLaporanKasExcel);
router.get('/pdf', exportLaporanKasPDF);
router.get('/kumulatif', getLaporanKumulatif);
router.get('/bulanan', getLaporanBulanan);

export default router;
