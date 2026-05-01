// Realistic D&T 419A inventory data based on the user's actual Item IDs
// CHM = Chemicals, TOL = Tools, MAT = Materials, MAC = Machines, ITM = Items, MAG = Magnets/misc

window.DT_DATA = {
  rooms: [
    { code: '419A', name: 'D&T Workshop 419A', rollout: true, locationCount: 14, itemCount: 87, attention: 6 },
    { code: 'V++', name: 'V++ Maker Studio', rollout: false, locationCount: 9, itemCount: 52, attention: 2 },
    { code: '420', name: 'Materials Store 420', rollout: false, locationCount: 6, itemCount: 31, attention: 0 },
    { code: '418', name: 'Electronics Lab 418', rollout: false, locationCount: 5, itemCount: 24, attention: 1 },
  ],

  locations: [
    // 419A — populated rollout room
    { id: 'CHM-CAB-01', room: '419A', label: 'Chemical Cabinet A', specific: 'North wall, locked', code: 'L-CHM-A', items: 6, attention: 2, hazard: true, qrReady: true },
    { id: 'CHM-CAB-02', room: '419A', label: 'Chemical Cabinet B', specific: 'North wall, locked', code: 'L-CHM-B', items: 4, attention: 0, hazard: true, qrReady: true },
    { id: 'TOL-WALL-01', room: '419A', label: 'Hand Tool Wall', specific: 'East wall pegboard', code: 'L-TOL-1', items: 18, attention: 1, hazard: false, qrReady: true },
    { id: 'TOL-DRW-02', room: '419A', label: 'Tool Drawer 2', specific: 'Bench 2, top drawer', code: 'L-TOL-2', items: 9, attention: 0, hazard: false, qrReady: true },
    { id: 'TOL-DRW-03', room: '419A', label: 'Tool Drawer 3', specific: 'Bench 2, mid drawer', code: 'L-TOL-3', items: 7, attention: 1, hazard: false, qrReady: true },
    { id: 'MAT-TRL-01', room: '419A', label: 'Material Trolley 1', specific: 'Mobile, near door', code: 'L-MAT-1', items: 12, attention: 0, hazard: false, qrReady: true },
    { id: 'MAT-SHF-04', room: '419A', label: 'Sheet Material Shelf', specific: 'West wall, low', code: 'L-MAT-4', items: 5, attention: 0, hazard: false, qrReady: false },
    { id: 'MAC-ZONE-01', room: '419A', label: 'Laser Cutter Zone', specific: 'Glowforge bay', code: 'L-MAC-1', items: 3, attention: 1, hazard: false, qrReady: true },
    { id: 'MAC-ZONE-02', room: '419A', label: '3D Printer Zone', specific: 'Bambu rack', code: 'L-MAC-2', items: 4, attention: 0, hazard: false, qrReady: true },
    { id: 'MAG-TRY-01', room: '419A', label: 'Fastener Tray 1', specific: 'Bench 1, tray rack', code: 'L-MAG-1', items: 11, attention: 0, hazard: false, qrReady: true },
    { id: 'ITM-CON-07', room: '419A', label: 'Consumables Bin 7', specific: 'Storage shelf C', code: 'L-ITM-7', items: 8, attention: 1, hazard: false, qrReady: true },
    { id: 'ITM-CON-08', room: '419A', label: 'Consumables Bin 8', specific: 'Storage shelf C', code: 'L-ITM-8', items: 0, attention: 0, hazard: false, qrReady: false },
    { id: 'TOL-DRW-04', room: '419A', label: 'Tool Drawer 4', specific: 'Bench 3, top drawer', code: 'L-TOL-4', items: 0, attention: 0, hazard: false, qrReady: true },
    { id: 'CHM-CAB-03', room: '419A', label: 'Chemical Cabinet C', specific: 'Storeroom annex', code: 'L-CHM-C', items: 0, attention: 0, hazard: true, qrReady: false },
    // V++
    { id: 'VPP-BENCH-01', room: 'V++', label: 'Maker Bench 1', specific: 'Centre island', code: 'L-VPP-1', items: 14, attention: 1, hazard: false, qrReady: true },
    { id: 'VPP-ELEC-01', room: 'V++', label: 'Electronics Drawer', specific: 'East cabinet', code: 'L-VPP-2', items: 22, attention: 0, hazard: false, qrReady: true },
    { id: 'VPP-3DP-01', room: 'V++', label: '3D Print Filament Rack', specific: 'Print zone', code: 'L-VPP-3', items: 8, attention: 1, hazard: false, qrReady: true },
    // 420
    { id: 'MAT-STORE-01', room: '420', label: 'Plywood Stack', specific: 'Materials store', code: 'L-420-1', items: 6, attention: 0, hazard: false, qrReady: true },
    { id: 'MAT-STORE-02', room: '420', label: 'Acrylic Stack', specific: 'Materials store', code: 'L-420-2', items: 4, attention: 0, hazard: false, qrReady: true },
    // 418
    { id: 'ELEC-DRW-01', room: '418', label: 'Microcontroller Drawer', specific: 'Electronics bench', code: 'L-418-1', items: 12, attention: 1, hazard: false, qrReady: true },
  ],

  // Items keyed by storage id
  items: {
    'CHM-CAB-01': [
      { id: 'CHM-002', name: 'Acetone (500ml)', category: 'Chemicals', qty: 2, unit: 'bottle', status: 'Good', remarks: 'Use in fume area', hazard: true },
      { id: 'CHM-014', name: 'Isopropyl Alcohol 99%', category: 'Chemicals', qty: 1, unit: 'bottle', status: 'Low Stock', remarks: 'Reorder this week', hazard: true },
      { id: 'CHM-017', name: 'Wood Glue (PVA)', category: 'Chemicals', qty: 6, unit: 'bottle', status: 'Good', remarks: '', hazard: true },
      { id: 'CHM-021', name: 'Contact Cement', category: 'Chemicals', qty: 0, unit: 'tin', status: 'Missing', remarks: 'Last seen Mar 12', hazard: true },
      { id: 'CHM-022', name: 'Spray Mount', category: 'Chemicals', qty: 3, unit: 'can', status: 'Good', remarks: '', hazard: true },
      { id: 'CHM-023', name: 'Methylated Spirit', category: 'Chemicals', qty: 2, unit: 'bottle', status: 'Good', remarks: '', hazard: true },
    ],
    'TOL-WALL-01': [
      { id: 'TOL-001', name: 'Tenon Saw 12"', category: 'Tools', qty: 4, unit: 'pcs', status: 'Good', remarks: '', hazard: false },
      { id: 'TOL-005', name: 'Coping Saw', category: 'Tools', qty: 6, unit: 'pcs', status: 'Good', remarks: '', hazard: false },
      { id: 'TOL-007', name: 'Claw Hammer 16oz', category: 'Tools', qty: 8, unit: 'pcs', status: 'Good', remarks: '', hazard: false },
      { id: 'TOL-011', name: 'Mallet (rubber)', category: 'Tools', qty: 5, unit: 'pcs', status: 'Good', remarks: '', hazard: false },
      { id: 'TOL-013', name: 'Try Square 6"', category: 'Tools', qty: 7, unit: 'pcs', status: 'Needs Maintenance', remarks: 'Two have bent blades', hazard: false },
      { id: 'TOL-014', name: 'Steel Rule 300mm', category: 'Tools', qty: 12, unit: 'pcs', status: 'Good', remarks: '', hazard: false },
      { id: 'TOL-022', name: 'Marking Gauge', category: 'Tools', qty: 4, unit: 'pcs', status: 'Good', remarks: '', hazard: false },
      { id: 'TOL-023', name: 'Engineer Square', category: 'Tools', qty: 3, unit: 'pcs', status: 'Good', remarks: '', hazard: false },
    ],
    'MAC-ZONE-01': [
      { id: 'MAC-001', name: 'Glowforge Pro Laser Cutter', category: 'Machines', qty: 1, unit: 'unit', status: 'Needs Maintenance', remarks: 'Lens cleaning due', hazard: false },
      { id: 'MAC-005', name: 'Fume Extractor', category: 'Machines', qty: 1, unit: 'unit', status: 'Good', remarks: '', hazard: false },
      { id: 'MAC-006', name: 'Air Compressor', category: 'Machines', qty: 1, unit: 'unit', status: 'Good', remarks: '', hazard: false },
    ],
    'ITM-CON-08': [], // empty case
  },

  warnings: {
    webAppBaseUrl: 'WEB_APP_BASE_URL is not configured. In-app browsing still works, but QR generation and external links are disabled.',
  },

  readiness: {
    score: 78,
    errors: 2,
    warnings: 9,
    missingQr: 3,
    missingStorageId: 1,
    invalidQty: 0,
    duplicates: 1,
    rollout419A: { ready: 11, total: 14, pct: 79 },
  },
};
