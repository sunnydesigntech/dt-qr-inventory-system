// VSA D&T QR Inventory — realistic operational data
// Storage IDs follow the pattern <ROOM>-<TYPE>-<NN>, e.g. 419A-CHEM-001

window.DT_DATA = (function () {

  const rooms = [
    { code: '419A', name: 'D&T Workshop 419A', rollout: true,  storageCount: 14, itemCount: 87, attention: 6 },
    { code: 'V++',  name: 'V++ Maker Studio',  rollout: false, storageCount: 9,  itemCount: 52, attention: 2 },
    { code: '415B', name: 'Materials Store 415B', rollout: false, storageCount: 6, itemCount: 31, attention: 1 },
    { code: '415C', name: 'Electronics Lab 415C', rollout: false, storageCount: 5, itemCount: 24, attention: 1 },
  ];

  // Storage points
  const locations = [
    // 419A — rollout room
    { id: '419A-CHEM-001', room: '419A', label: 'Chemical Cabinet A', specific: 'North wall, locked', code: 'L-CHM-A', storageType: 'Cabinet · Locked', items: 6, attention: 2, hazard: true,  qrReady: true,  scanTested: true,  lastChecked: '2026-04-22', notes: 'PPE station alongside' },
    { id: '419A-CHEM-002', room: '419A', label: 'Chemical Cabinet B', specific: 'North wall, locked', code: 'L-CHM-B', storageType: 'Cabinet · Locked', items: 4, attention: 0, hazard: true,  qrReady: true,  scanTested: true,  lastChecked: '2026-04-22', notes: '' },
    { id: '419A-TOOL-001', room: '419A', label: 'Hand Tool Wall',     specific: 'East wall pegboard', code: 'L-TOL-1', storageType: 'Pegboard',         items: 18, attention: 1, hazard: false, qrReady: true,  scanTested: true,  lastChecked: '2026-04-25', notes: '' },
    { id: '419A-TOOL-002', room: '419A', label: 'Tool Drawer 2',      specific: 'Bench 2, top drawer', code: 'L-TOL-2', storageType: 'Drawer',          items: 9,  attention: 0, hazard: false, qrReady: true,  scanTested: true,  lastChecked: '2026-04-25', notes: '' },
    { id: '419A-TOOL-003', room: '419A', label: 'Tool Drawer 3',      specific: 'Bench 2, mid drawer', code: 'L-TOL-3', storageType: 'Drawer',          items: 7,  attention: 1, hazard: false, qrReady: true,  scanTested: false, lastChecked: '2026-04-25', notes: 'Re-test after relabel' },
    { id: '419A-MAT-001',  room: '419A', label: 'Material Trolley 1', specific: 'Mobile, near door',   code: 'L-MAT-1', storageType: 'Trolley · Mobile', items: 12, attention: 0, hazard: false, qrReady: true,  scanTested: true,  lastChecked: '2026-04-25', notes: '' },
    { id: '419A-MAT-002',  room: '419A', label: 'Sheet Material Shelf', specific: 'West wall, low',    code: 'L-MAT-4', storageType: 'Shelf',           items: 5,  attention: 0, hazard: false, qrReady: false, scanTested: false, lastChecked: '',          notes: 'QR not yet generated' },
    { id: '419A-MACH-001', room: '419A', label: 'Laser Cutter Zone',  specific: 'Glowforge bay',       code: 'L-MAC-1', storageType: 'Machine zone',    items: 3,  attention: 1, hazard: false, qrReady: true,  scanTested: true,  lastChecked: '2026-04-20', notes: 'Lens cleaning due' },
    { id: '419A-MACH-002', room: '419A', label: '3D Printer Zone',    specific: 'Bambu rack',          code: 'L-MAC-2', storageType: 'Machine zone',    items: 4,  attention: 0, hazard: false, qrReady: true,  scanTested: true,  lastChecked: '2026-04-20', notes: '' },
    { id: '419A-CONS-001', room: '419A', label: 'Fastener Tray 1',    specific: 'Bench 1, tray rack',  code: 'L-MAG-1', storageType: 'Tray',            items: 11, attention: 0, hazard: false, qrReady: true,  scanTested: true,  lastChecked: '2026-04-25', notes: '' },
    { id: '419A-CONS-002', room: '419A', label: 'Consumables Bin 7',  specific: 'Storage shelf C',     code: 'L-ITM-7', storageType: 'Bin',             items: 8,  attention: 1, hazard: false, qrReady: true,  scanTested: true,  lastChecked: '2026-04-25', notes: '' },
    { id: '419A-CONS-003', room: '419A', label: 'Consumables Bin 8',  specific: 'Storage shelf C',     code: 'L-ITM-8', storageType: 'Bin',             items: 0,  attention: 0, hazard: false, qrReady: false, scanTested: false, lastChecked: '',          notes: 'Empty placeholder' },
    { id: '419A-TOOL-004', room: '419A', label: 'Tool Drawer 4',      specific: 'Bench 3, top drawer', code: 'L-TOL-4', storageType: 'Drawer',          items: 0,  attention: 0, hazard: false, qrReady: true,  scanTested: false, lastChecked: '',          notes: 'Empty placeholder' },
    { id: '419A-CHEM-003', room: '419A', label: 'Chemical Cabinet C', specific: 'Storeroom annex',     code: 'L-CHM-C', storageType: 'Cabinet · Locked', items: 0,  attention: 0, hazard: true,  qrReady: false, scanTested: false, lastChecked: '',          notes: 'Pending hazard approval' },
    // V++
    { id: 'VPP-BENCH-001', room: 'V++',  label: 'Maker Bench 1',      specific: 'Centre island',       code: 'L-VPP-1', storageType: 'Bench',           items: 14, attention: 1, hazard: false, qrReady: true, scanTested: true, lastChecked: '2026-04-18', notes: '' },
    { id: 'VPP-ELEC-001',  room: 'V++',  label: 'Electronics Drawer', specific: 'East cabinet',        code: 'L-VPP-2', storageType: 'Drawer',          items: 22, attention: 0, hazard: false, qrReady: true, scanTested: true, lastChecked: '2026-04-18', notes: '' },
    { id: 'VPP-3DP-001',   room: 'V++',  label: '3D Print Filament Rack', specific: 'Print zone',      code: 'L-VPP-3', storageType: 'Rack',            items: 8,  attention: 1, hazard: false, qrReady: true, scanTested: true, lastChecked: '2026-04-18', notes: '' },
    // 415B
    { id: '415B-MAT-001',  room: '415B', label: 'Plywood Stack',      specific: 'Materials store',     code: 'L-415B-1', storageType: 'Rack',           items: 6,  attention: 0, hazard: false, qrReady: true, scanTested: true, lastChecked: '2026-04-15', notes: '' },
    { id: '415B-MAT-002',  room: '415B', label: 'Acrylic Stack',      specific: 'Materials store',     code: 'L-415B-2', storageType: 'Rack',           items: 4,  attention: 1, hazard: false, qrReady: true, scanTested: true, lastChecked: '2026-04-15', notes: '' },
    // 415C
    { id: '415C-ELEC-001', room: '415C', label: 'Microcontroller Drawer', specific: 'Electronics bench', code: 'L-415C-1', storageType: 'Drawer',      items: 12, attention: 1, hazard: false, qrReady: true, scanTested: true, lastChecked: '2026-04-12', notes: '' },
    { id: '415C-ELEC-002', room: '415C', label: 'Solder Station',     specific: 'East bench',          code: 'L-415C-2', storageType: 'Station',        items: 5,  attention: 0, hazard: true,  qrReady: true, scanTested: true, lastChecked: '2026-04-12', notes: 'Flux fume warning' },
  ];

  const items = {
    '419A-CHEM-001': [
      { id: 'CHM-002', name: 'Acetone (500ml)',         category: 'Chemicals', qty: 2, unit: 'bottle', status: 'Good',      remarks: 'Use in fume area', hazard: true, safetyNote: 'Flammable. Use only in fume cupboard. Goggles + nitrile gloves.', sds: 'sds/CHM-002.pdf', reorderLevel: 2 },
      { id: 'CHM-014', name: 'Isopropyl Alcohol 99%',   category: 'Chemicals', qty: 1, unit: 'bottle', status: 'Low Stock', remarks: 'Reorder this week', hazard: true, safetyNote: 'Flammable. Avoid open flame.', sds: 'sds/CHM-014.pdf', reorderLevel: 3, supplier: 'School Lab Supplies', purchaseLink: '#', cost: 'HK$48' },
      { id: 'CHM-017', name: 'Wood Glue (PVA)',         category: 'Chemicals', qty: 6, unit: 'bottle', status: 'Good',      remarks: '', hazard: true, safetyNote: 'Skin contact: wash with water.', sds: 'sds/CHM-017.pdf' },
      { id: 'CHM-021', name: 'Contact Cement',          category: 'Chemicals', qty: 0, unit: 'tin',    status: 'Missing',   remarks: 'Last seen Mar 12', hazard: true, safetyNote: 'Highly flammable. Strictly fume cupboard only.', sds: 'sds/CHM-021.pdf' },
      { id: 'CHM-022', name: 'Spray Mount',             category: 'Chemicals', qty: 3, unit: 'can',    status: 'Good',      remarks: '', hazard: true, safetyNote: 'Aerosol. Use outdoors or in fume cupboard.', sds: 'sds/CHM-022.pdf' },
      { id: 'CHM-023', name: 'Methylated Spirit',       category: 'Chemicals', qty: 2, unit: 'bottle', status: 'Good',      remarks: '', hazard: true, safetyNote: 'Flammable. Keep away from heat.', sds: 'sds/CHM-023.pdf' },
    ],
    '419A-TOOL-001': [
      { id: 'TOL-001', name: 'Tenon Saw 12"',           category: 'Tools', qty: 4,  unit: 'pcs', status: 'Good', remarks: '' },
      { id: 'TOL-005', name: 'Coping Saw',              category: 'Tools', qty: 6,  unit: 'pcs', status: 'Good', remarks: '' },
      { id: 'TOL-007', name: 'Claw Hammer 16oz',        category: 'Tools', qty: 8,  unit: 'pcs', status: 'Good', remarks: '' },
      { id: 'TOL-011', name: 'Mallet (rubber)',         category: 'Tools', qty: 5,  unit: 'pcs', status: 'Good', remarks: '' },
      { id: 'TOL-013', name: 'Try Square 6"',           category: 'Tools', qty: 7,  unit: 'pcs', status: 'Needs Maintenance', remarks: 'Two have bent blades', maintenanceDue: '2026-05-08' },
      { id: 'TOL-014', name: 'Steel Rule 300mm',        category: 'Tools', qty: 12, unit: 'pcs', status: 'Good', remarks: '' },
      { id: 'TOL-022', name: 'Marking Gauge',           category: 'Tools', qty: 4,  unit: 'pcs', status: 'Good', remarks: '' },
      { id: 'TOL-023', name: 'Engineer Square',         category: 'Tools', qty: 3,  unit: 'pcs', status: 'Low Stock', remarks: 'Order 3 more', reorderLevel: 4, supplier: 'D&T Supplies HK', cost: 'HK$120' },
    ],
    '419A-MACH-001': [
      { id: 'MAC-001', name: 'Glowforge Pro Laser Cutter', category: 'Machines', qty: 1, unit: 'unit', status: 'Needs Maintenance', remarks: 'Lens cleaning due', maintenanceDue: '2026-05-02', assetValue: 'HK$58,000' },
      { id: 'MAC-005', name: 'Fume Extractor',             category: 'Machines', qty: 1, unit: 'unit', status: 'Good', remarks: '', maintenanceDue: '2026-08-15' },
      { id: 'MAC-006', name: 'Air Compressor',             category: 'Machines', qty: 1, unit: 'unit', status: 'Good', remarks: '', maintenanceDue: '2026-07-01' },
    ],
    '419A-CONS-003': [],
    'VPP-ELEC-001': [
      { id: 'ELE-001', name: 'Arduino Uno R4',          category: 'Electronics', qty: 8, unit: 'pcs', status: 'Good', remarks: '' },
      { id: 'ELE-014', name: 'Jumper Wire Set (M-M)',   category: 'Electronics', qty: 12, unit: 'pack', status: 'Good', remarks: '' },
      { id: 'ELE-022', name: 'Breadboard 830-pt',       category: 'Electronics', qty: 14, unit: 'pcs', status: 'Good', remarks: '' },
    ],
    'VPP-3DP-001': [
      { id: 'CON-031', name: 'PLA Filament 1.75mm Black', category: 'Consumables', qty: 2, unit: 'roll', status: 'Low Stock', remarks: 'Reorder before next print week', reorderLevel: 4, supplier: 'Bambu HK', cost: 'HK$165/roll' },
      { id: 'CON-035', name: 'PETG Filament 1.75mm Clear', category: 'Consumables', qty: 5, unit: 'roll', status: 'Good', remarks: '' },
    ],
    '415B-MAT-002': [
      { id: 'MAT-014', name: 'Acrylic Sheet 3mm 600×400 Clear', category: 'Materials', qty: 3, unit: 'sheet', status: 'Low Stock', remarks: 'Y9 project consumed 6 sheets', reorderLevel: 6, supplier: 'PSP Plastics', cost: 'HK$85/sheet', purchaseLink: '#' },
    ],
    '415C-ELEC-001': [
      { id: 'ELE-041', name: 'micro:bit V2',             category: 'Electronics', qty: 16, unit: 'pcs', status: 'Good', remarks: '' },
      { id: 'ELE-052', name: 'Servo SG90',               category: 'Electronics', qty: 4,  unit: 'pcs', status: 'Low Stock', remarks: '', reorderLevel: 8, supplier: 'Pololu HK' },
    ],
    '415C-ELEC-002': [
      { id: 'MAC-012', name: 'Solder Station (60W)',     category: 'Machines', qty: 2, unit: 'unit', status: 'Good', remarks: '', maintenanceDue: '2026-09-30', assetValue: 'HK$1,200/ea' },
      { id: 'CHM-031', name: 'Lead-free Solder 0.8mm',   category: 'Chemicals', qty: 2, unit: 'roll', status: 'Good', remarks: '', hazard: true, safetyNote: 'Wash hands after use. Use fume extractor.', sds: 'sds/CHM-031.pdf' },
    ],
  };

  const audit = [
    { ts: '2026-04-29 14:22', user: 'm.lai@vsa.edu', action: 'Update', room: '419A', storage: '419A-CHEM-001', item: 'CHM-014', oldQty: 2, newQty: 1, oldStatus: 'Good', newStatus: 'Low Stock', notes: 'Used in Y10 project' },
    { ts: '2026-04-28 10:05', user: 'k.wong@vsa.edu', action: 'Update', room: '419A', storage: '419A-TOOL-001', item: 'TOL-013', oldQty: 7, newQty: 7, oldStatus: 'Good', newStatus: 'Needs Maintenance', notes: 'Bent blades flagged' },
    { ts: '2026-04-28 09:41', user: 'k.wong@vsa.edu', action: 'Update', room: '419A', storage: '419A-MACH-001', item: 'MAC-001', oldQty: 1, newQty: 1, oldStatus: 'Good', newStatus: 'Needs Maintenance', notes: 'Lens cleaning' },
    { ts: '2026-04-26 16:08', user: 'm.lai@vsa.edu', action: 'Update', room: '419A', storage: '419A-CHEM-001', item: 'CHM-021', oldQty: 1, newQty: 0, oldStatus: 'Good', newStatus: 'Missing', notes: 'Last seen Mar 12' },
    { ts: '2026-04-26 11:30', user: 'admin@vsa.edu', action: 'Refresh QR', room: '419A', storage: '419A-MAT-001', item: '—', oldQty: '—', newQty: '—', oldStatus: '—', newStatus: '—', notes: 'Bulk QR rebuild' },
    { ts: '2026-04-25 13:55', user: 'k.wong@vsa.edu', action: 'Update', room: 'V++', storage: 'VPP-3DP-001', item: 'CON-031', oldQty: 4, newQty: 2, oldStatus: 'Good', newStatus: 'Low Stock', notes: '' },
  ];

  return {
    rooms,
    locations,
    items,
    audit,
    warnings: { webAppBaseUrl: 'WEB_APP_BASE_URL is not configured. Browsing works, but QR generation needs WEB_APP_BASE_URL.' },
    readiness: {
      score: 78, errors: 2, warnings: 9, missingQr: 3, missingStorageId: 1, invalidQty: 0, invalidStatus: 0,
      duplicates: 1, chemMissingSafety: 0, chemMissingSds: 1, lowStockMissingReorder: 1, maintMissingDue: 0, vppUrlIssues: 0,
      rollout419A: { ready: 11, total: 14, pct: 79 },
    },
  };
})();
