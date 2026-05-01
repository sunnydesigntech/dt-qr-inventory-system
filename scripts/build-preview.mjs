import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');

const sampleLocationKey = '419a||storageid||chm-cab-01';
const emptyLocationKey = '419a||storageid||tol-drw-04';

const bootstrap = {
  appTitle: 'D&T QR Inventory System',
  webAppBaseUrl: '',
  route: { name: 'landing' },
  error: '',
  warnings: [
    'WEB_APP_BASE_URL is not configured. In-app browsing still works; QR generation and external links are disabled until it is set.'
  ],
  diagnostics: {
    spreadsheetIdSet: true,
    webAppBaseUrlSet: false,
    inventorySheetNameSet: true,
    sheetFound: true,
    sheetInUse: 'Inventory · local preview',
    requiredColumnsOk: true,
    missingRequiredColumns: [],
    qrImageColumn: true
  },
  appData: {
    rooms: [
      { code: '419A', name: 'D&T Workshop 419A', rollout: true, locationCount: 2, itemCount: 2, attention: 1 },
      { code: 'V++', name: 'V++ Maker Studio', rollout: false, locationCount: 1, itemCount: 1, attention: 0 }
    ],
    locations: [
      {
        key: sampleLocationKey,
        id: 'CHM-CAB-01',
        room: '419A',
        label: 'Chemical Cabinet A',
        specific: 'North wall, locked',
        code: 'L-CHM-A',
        storageId: 'CHM-CAB-01',
        storageLabel: 'Chemical Cabinet A',
        routeLoc: 'CHM-CAB-01',
        searchText: '419A Chemical Cabinet A CHM-CAB-01 L-CHM-A North wall locked',
        sortKey: '419A | CHM-CAB-01',
        items: 2,
        attention: 1,
        hazard: true,
        qrReady: true,
        viewUrl: '?room=419A&loc=CHM-CAB-01',
        techUrl: '?room=419A&loc=CHM-CAB-01&mode=tech',
        qrImageUrl: ''
      },
      {
        key: emptyLocationKey,
        id: 'TOL-DRW-04',
        room: '419A',
        label: 'Tool Drawer 4',
        specific: 'Bench 3, top drawer',
        code: 'L-TOL-4',
        storageId: 'TOL-DRW-04',
        storageLabel: 'Tool Drawer 4',
        routeLoc: 'TOL-DRW-04',
        searchText: '419A Tool Drawer 4 TOL-DRW-04 L-TOL-4 Bench 3 top drawer',
        sortKey: '419A | TOL-DRW-04',
        items: 0,
        attention: 0,
        hazard: false,
        qrReady: true,
        viewUrl: '?room=419A&loc=TOL-DRW-04',
        techUrl: '?room=419A&loc=TOL-DRW-04&mode=tech',
        qrImageUrl: ''
      },
      {
        key: 'v++||location||maker bench 1',
        id: 'Maker Bench 1',
        room: 'V++',
        label: 'Maker Bench 1',
        specific: 'Centre island',
        code: 'L-VPP-1',
        storageId: '',
        storageLabel: '',
        routeLoc: 'Maker Bench 1',
        searchText: 'V++ Maker Bench 1 Centre island L-VPP-1',
        sortKey: 'V++ | Maker Bench 1',
        items: 1,
        attention: 0,
        hazard: false,
        qrReady: true,
        viewUrl: '?room=V%2B%2B&loc=Maker%20Bench%201',
        techUrl: '?room=V%2B%2B&loc=Maker%20Bench%201&mode=tech',
        qrImageUrl: ''
      }
    ],
    items: {
      [sampleLocationKey]: [
        {
          sheetRow: 2,
          id: 'CHM-002',
          name: 'Acetone (500ml)',
          category: 'Chemicals',
          qty: 2,
          unit: 'bottle',
          status: 'Good',
          remarks: 'Use in fume area',
          hazard: true
        },
        {
          sheetRow: 3,
          id: 'CHM-014',
          name: 'Isopropyl Alcohol 99%',
          category: 'Chemicals',
          qty: 1.5,
          unit: 'bottle',
          status: 'Low Stock',
          remarks: 'Reorder this week',
          hazard: true
        }
      ],
      [emptyLocationKey]: [],
      'v++||location||maker bench 1': [
        {
          sheetRow: 4,
          id: 'VPP-001',
          name: 'Microcontroller kit',
          category: 'Electronics',
          qty: 8,
          unit: 'set',
          status: 'Good',
          remarks: '',
          hazard: false
        }
      ]
    },
    readiness: {
      score: 87,
      errors: 0,
      warnings: 4,
      missingQr: 0,
      missingStorageId: 0,
      invalidQty: 0,
      duplicates: 0,
      rollout419A: { ready: 2, total: 2, pct: 100, itemCount: 2, chemicalCount: 2 }
    }
  }
};

function read(name) {
  return fs.readFileSync(path.join(root, name), 'utf8');
}

const html = read('index.html')
  .replace("<?!= include('app_styles'); ?>", read('app_styles.html'))
  .replace('<?!= bootstrapJson ?>', JSON.stringify(bootstrap).replace(/</g, '\\u003c'))
  .replace("<?!= include('app_script'); ?>", read('app_script.html'));

fs.writeFileSync(path.join(root, 'preview.html'), html);
console.log('Wrote preview.html');
