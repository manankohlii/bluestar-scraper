import { test, expect } from '@playwright/test';
import { promises as fs } from 'fs';
import * as path from 'path';
import * as dotenv from 'dotenv';
import ExcelJS from 'exceljs';

dotenv.config();

const LOGIN_URL = 'https://www.medstatsupplies.com/scs/checkout.ssp?is=login&login=T&fragment=login-register#login-register';
const SEARCH_URL_PAGE1 = 'https://www.medstatsupplies.com/search?order=relevance:desc&keywords=BLUESTAR';
const OUTPUT_DIR = path.resolve('data');

async function ensureOutputDir(): Promise<void> {
  await fs.mkdir(OUTPUT_DIR, { recursive: true });
}

async function saveJson(filename: string, data: unknown): Promise<void> {
  await ensureOutputDir();
  const filePath = path.join(OUTPUT_DIR, filename);
  await fs.writeFile(filePath, JSON.stringify(data, null, 2), 'utf8');
  // eslint-disable-next-line no-console
  console.log(`Saved: ${filePath}`);
}

async function loginIfNeeded(page): Promise<void> {
  const email = process.env.MEDSTAT_EMAIL || '';
  const password = process.env.MEDSTAT_PASSWORD || '';
  if (!email || !password) {
    throw new Error('Please set MEDSTAT_EMAIL and MEDSTAT_PASSWORD in your .env file.');
  }

  await page.goto(LOGIN_URL, { waitUntil: 'domcontentloaded' });
  // Fill email & password (selectors based on codegen ARIA roles)
  await page.getByRole('textbox', { name: /Email Address/i }).click();
  await page.getByRole('textbox', { name: /Email Address/i }).fill(email);
  await page.getByRole('textbox', { name: /Password/i }).click();
  await page.getByRole('textbox', { name: /Password/i }).fill(password);
  await page.getByRole('button', { name: /Log In/i }).click();

  // Wait until we leave the login URL or the login form disappears; avoid strict mode issues.
  await Promise.race([
    page.waitForURL((url) => !url.toString().includes('login'), { timeout: 15000 }),
    page.getByRole('textbox', { name: /Email Address/i }).waitFor({ state: 'detached', timeout: 15000 })
  ]).catch(() => undefined);
  await page.waitForLoadState('networkidle');
}

async function collectFirstPageProductUrls(page): Promise<string[]> {
  await page.goto(SEARCH_URL_PAGE1, { waitUntil: 'domcontentloaded' });
  await page.waitForLoadState('networkidle');

  // Wait for product grid to render; try common SCA class, fallback to at least some links
  const gridCell = page.locator('div.facets-item-cell, li.facets-item-cell');
  await gridCell.first().waitFor({ timeout: 15000 }).catch(() => undefined);

  // Prefer explicit product tile anchors on SuiteCommerce search grid
  const byImage = await page.$$eval('a.facets-item-cell-grid-link-image[href]', (as) =>
    Array.from(new Set(as.map((a) => (a as HTMLAnchorElement).href)))
  ).catch(() => [] as string[]);
  const byTitle = await page.$$eval('a.facets-item-cell-grid-link-title[href]', (as) =>
    Array.from(new Set(as.map((a) => (a as HTMLAnchorElement).href)))
  ).catch(() => [] as string[]);

  // Fallback: all anchors filtered by pathname heuristics
  const allHrefs = await page.$$eval('a[href]', (as) => Array.from(new Set(as.map((a) => (a as HTMLAnchorElement).href))));
  const byHeuristic = allHrefs.filter((u) => {
    try {
      const url = new URL(u);
      if (!/https?:/i.test(url.protocol)) return false;
      if (url.pathname.includes('/search')) return false;
      if (url.pathname.includes('/checkout')) return false;
      if (url.hash) return false;
      return /\/product\//i.test(url.pathname) || /\/p\//i.test(url.pathname) || /\/sku\//i.test(url.pathname) || /\/prod\//i.test(url.pathname);
    } catch {
      return false;
    }
  });

  const unique = Array.from(new Set([...byTitle, ...byImage, ...byHeuristic]));
  // eslint-disable-next-line no-console
  console.log(`Found ${unique.length} product URLs on page 1`);
  return unique;
}

type ProductRecord = {
  productUrl: string;
  productName: string | null;
  sku: string | null;
  mpn: string | null;
  manufacturer: string | null;
  price: string | null;
  stock: string | null;
  description: string | null;
};

async function extractProductDetails(page): Promise<ProductRecord> {
  const productUrl = page.url();
  const productName = (await page.locator('h1').first().textContent().catch(() => null))?.trim() || null;

  // Scope to details container if present
  const detailsLocator = page.locator('#product-details-full-form');
  const detailsText = (await detailsLocator.first().innerText().catch(() => '')) || (await page.innerText('body').catch(() => ''));

  const matchAfter = (labelRegex: RegExp): string | null => {
    const m = detailsText.match(labelRegex);
    if (!m) return null;
    return m[1]?.trim() || null;
  };

  const sku = matchAfter(/\bSKU\s*[:#-]?\s*([^\n]+)/i) || matchAfter(/\bItem\s*[:#-]?\s*([^\n]+)/i) || null;
  const mpn = matchAfter(/\bMPN\s*[:#-]?\s*([^\n]+)/i) || null;
  const manufacturer = matchAfter(/\bMANUFACTURER\s*[:#-]?\s*([^\n]+)/i) || null;
  const description = matchAfter(/\bDescription\s*[:#-]?\s*([^\n][\s\S]*?)$/i) || null;

  // Price: first currency found near details
  const priceMatch = detailsText.match(/\$[\d,.]+/);
  const price = priceMatch ? priceMatch[0] : null;

  const stock = matchAfter(/\bCurrent Stock\s*[:#-]?\s*([^\n]+)/i) || null;

  return { productUrl, productName, sku, mpn, manufacturer, price, stock, description };
}

test('Scrape BLUESTAR page 1 products and save data', async ({ page }) => {
  test.slow();
  test.setTimeout(10 * 60 * 1000);
  await ensureOutputDir();

  await loginIfNeeded(page);

  const productUrls = await collectFirstPageProductUrls(page);
  await saveJson('product_urls_page1.json', productUrls);

  const results: ProductRecord[] = [];
  let processed = 0;
  for (const url of productUrls) {
    // eslint-disable-next-line no-console
    console.log(`Processing (${++processed}/${productUrls.length}): ${url}`);
    await page.goto(url, { waitUntil: 'domcontentloaded' });
    await page.waitForLoadState('networkidle');

    // Wait for any product-specific marker
    await page.locator('h1, h2.product-title, .product-details-info').first().waitFor({ timeout: 15000 }).catch(() => undefined);
    const record = await extractProductDetails(page);
    results.push(record);
  }

  await saveJson('products_page1.json', results);

  // Also append to the consolidated Excel workbook used across runs
  await (async function writeExcel(records: ProductRecord[], excelFilename: string): Promise<string> {
    const filePath = path.join(OUTPUT_DIR, excelFilename);
    await fs.mkdir(OUTPUT_DIR, { recursive: true });

    const workbook = new ExcelJS.Workbook();
    let sheet = undefined as ExcelJS.Worksheet | undefined;
    try {
      await fs.stat(filePath);
      await workbook.xlsx.readFile(filePath);
      sheet = workbook.getWorksheet('Products') || workbook.worksheets[0];
    } catch {
      // brand new file
    }
    if (!sheet) {
      sheet = workbook.addWorksheet('Products');
    }

    const defaultHeaders = [
      'DATE',
      'TIME',
      'Item Name',
      'SKU',
      'DESCRIPTION',
      'MPN',
      'MANUFACTURER',
      'PRICE',
      'STOCK',
      'PRODUCT URL',
    ];

    // Canonical columns and header enforcement to avoid index drift
    const columnDefs = [
      { header: 'DATE', key: 'runDate', width: 12 },
      { header: 'TIME', key: 'runTime', width: 18 },
      { header: 'Item Name', key: 'productName', width: 50 },
      { header: 'SKU', key: 'sku', width: 30 },
      { header: 'DESCRIPTION', key: 'description', width: 80 },
      { header: 'MPN', key: 'mpn', width: 30 },
      { header: 'MANUFACTURER', key: 'manufacturer', width: 30 },
      { header: 'PRICE', key: 'price', width: 15 },
      { header: 'STOCK', key: 'stock', width: 15 },
      { header: 'PRODUCT URL', key: 'productUrl', width: 80 },
    ];
    if (sheet.rowCount === 0) {
      sheet.columns = columnDefs;
    } else {
      const headerRow = sheet.getRow(1);
      headerRow.values = [
        undefined,
        'DATE',
        'TIME',
        'Item Name',
        'SKU',
        'DESCRIPTION',
        'MPN',
        'MANUFACTURER',
        'PRICE',
        'STOCK',
        'PRODUCT URL',
      ];
      headerRow.commit();
      sheet.columns = columnDefs;
    }

    try { sheet.getColumn('price').numFmt = '$#,##0.00'; } catch {}

    const parsePrice = (p: string | null): number | null => {
      if (!p) return null;
      const n = Number(p.replace(/[^0-9.\-]/g, ''));
      return Number.isFinite(n) ? n : null;
    };
    const parseStock = (s: string | null): number | null => {
      if (!s) return null;
      const cleaned = s.replace(/[^0-9\-]/g, '');
      if (!cleaned) return null;
      const n = Number.parseInt(cleaned, 10);
      return Number.isNaN(n) ? null : n;
    };

    const now = new Date();
    const runDate = new Intl.DateTimeFormat('en-US', { timeZone: 'America/Los_Angeles', month: '2-digit', day: '2-digit', year: 'numeric' }).format(now);
    const runTime = new Intl.DateTimeFormat('en-US', { timeZone: 'America/Los_Angeles', hour: '2-digit', minute: '2-digit', hour12: true, timeZoneName: 'short' }).format(now);

    for (const rec of records) {
      const priceNum = parsePrice(rec.price);
      const stockNum = parseStock(rec.stock);
      sheet.addRow({
        runDate,
        runTime,
        productName: rec.productName ?? null,
        sku: rec.sku ?? null,
        description: rec.description ?? null,
        mpn: rec.mpn ?? null,
        manufacturer: rec.manufacturer ?? null,
        price: priceNum ?? (rec.price ?? null),
        stock: stockNum ?? (rec.stock ?? null),
        productUrl: rec.productUrl ?? null,
      });
    }
    sheet.views = [{ state: 'frozen', ySplit: 1 }];

    await workbook.xlsx.writeFile(filePath);
    // eslint-disable-next-line no-console
    console.log(`Excel written: ${filePath}`);
    return filePath;
  })(results, 'products_all.xlsx');

  // Quick sanity checks
  expect(results.length).toBeGreaterThan(0);
  expect(results.every((r) => !!r.productUrl)).toBeTruthy();
});

