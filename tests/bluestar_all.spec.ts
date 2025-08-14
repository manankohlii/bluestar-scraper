import { test, expect } from '@playwright/test';
import { promises as fs } from 'fs';
import * as path from 'path';
import * as dotenv from 'dotenv';
import ExcelJS from 'exceljs';

dotenv.config();

const LOGIN_URL = 'https://www.medstatsupplies.com/scs/checkout.ssp?is=login&login=T&fragment=login-register#login-register';
const SEARCH_URL_BASE = 'https://www.medstatsupplies.com/search?order=relevance:desc&keywords=BLUESTAR';
const OUTPUT_DIR = path.resolve('data');
const TOTAL_PAGES = 7;

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
  await page.getByRole('textbox', { name: /Email Address/i }).click();
  await page.getByRole('textbox', { name: /Email Address/i }).fill(email);
  await page.getByRole('textbox', { name: /Password/i }).click();
  await page.getByRole('textbox', { name: /Password/i }).fill(password);
  await page.getByRole('button', { name: /Log In/i }).click();
  await Promise.race([
    page.waitForURL((url) => !url.toString().includes('login'), { timeout: 20000 }),
    page.getByRole('textbox', { name: /Email Address/i }).waitFor({ state: 'detached', timeout: 20000 })
  ]).catch(() => undefined);
  await page.waitForLoadState('networkidle');
}

async function collectProductUrlsForPage(page, pageIndex: number): Promise<string[]> {
  const url = pageIndex === 1 ? SEARCH_URL_BASE : `${SEARCH_URL_BASE}&page=${pageIndex}`;
  await page.goto(url, { waitUntil: 'domcontentloaded' });
  // Some pages keep network requests open; wait for product grid/tile instead of networkidle.
  await page
    .locator('a.facets-item-cell-grid-link-title[href], a.facets-item-cell-grid-link-image[href], .facets-item-cell')
    .first()
    .waitFor({ timeout: 30000 })
    .catch(() => undefined);

  const byImage = await page.$$eval('a.facets-item-cell-grid-link-image[href]', (as) =>
    Array.from(new Set(as.map((a) => (a as HTMLAnchorElement).href)))
  ).catch(() => [] as string[]);
  const byTitle = await page.$$eval('a.facets-item-cell-grid-link-title[href]', (as) =>
    Array.from(new Set(as.map((a) => (a as HTMLAnchorElement).href)))
  ).catch(() => [] as string[]);
  const allHrefs = await page.$$eval('a[href]', (as) => Array.from(new Set(as.map((a) => (a as HTMLAnchorElement).href))));
  const byHeuristic = allHrefs.filter((u) => {
    try {
      const l = new URL(u);
      if (!/https?:/i.test(l.protocol)) return false;
      if (l.pathname.includes('/search')) return false;
      if (l.pathname.includes('/checkout')) return false;
      if (l.hash) return false;
      return /\/product\//i.test(l.pathname) || /\/p\//i.test(l.pathname) || /\/sku\//i.test(l.pathname) || /\/prod\//i.test(l.pathname);
    } catch {
      return false;
    }
  });
  const unique = Array.from(new Set([...byTitle, ...byImage, ...byHeuristic]));
  // eslint-disable-next-line no-console
  console.log(`Page ${pageIndex}: found ${unique.length} product URLs`);
  return unique;
}

async function extractProductDetails(page): Promise<ProductRecord> {
  const productUrl = page.url();
  const productName = (await page.locator('h1').first().textContent().catch(() => null))?.trim() || null;
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
  const priceMatch = detailsText.match(/\$[\d,.]+/);
  const price = priceMatch ? priceMatch[0] : null;
  const stock = matchAfter(/\bCurrent Stock\s*[:#-]?\s*([^\n]+)/i) || null;
  return { productUrl, productName, sku, mpn, manufacturer, price, stock, description };
}

async function writeExcel(records: ProductRecord[], excelFilename: string): Promise<string> {
  await ensureOutputDir();
  const filePath = path.join(OUTPUT_DIR, excelFilename);

  const workbook = new ExcelJS.Workbook();
  let sheet = undefined as ExcelJS.Worksheet | undefined;

  // Try to read existing workbook to append; otherwise create new
  try {
    await fs.stat(filePath);
    await workbook.xlsx.readFile(filePath);
    sheet = workbook.getWorksheet('Products') || workbook.worksheets[0];
  } catch {
    // no existing file; will create new
  }

  if (!sheet) {
    sheet = workbook.addWorksheet('Products');
  }

  // Build or extend headers
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

  let headers: string[] = [];
  if (sheet.rowCount === 0) {
    headers = defaultHeaders;
    sheet.columns = [
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
  } else {
    const headerRow = sheet.getRow(1);
    headers = headerRow.values
      .slice(1)
      .map((v) => (typeof v === 'string' ? v : (v as any)?.richText?.map((rt: any) => rt.text).join('') || '')) as string[];
    if (!headers.includes('MANUFACTURER')) {
      headers.push('MANUFACTURER');
      const newHeaderValues = [undefined, ...headers];
      headerRow.values = newHeaderValues;
      headerRow.commit();
    }
  }

  // Helper: header -> column index
  const headerIndex: Record<string, number> = {};
  headers.forEach((h, idx) => {
    headerIndex[h] = idx + 1; // 1-based
  });

  // Apply number formats if possible
  const priceColIdx = headerIndex['PRICE'];
  if (priceColIdx) {
    sheet.getColumn(priceColIdx).numFmt = '$#,##0.00';
  }

  const parsePriceToNumber = (p: string | null): number | null => {
    if (!p) return null;
    const n = Number(p.replace(/[^0-9.\-]/g, ''));
    return Number.isFinite(n) ? n : null;
  };
  const parseStockToNumber = (s: string | null): number | null => {
    if (!s) return null;
    const cleaned = s.replace(/[^0-9\-]/g, '');
    if (!cleaned) return null;
    const n = Number.parseInt(cleaned, 10);
    return Number.isNaN(n) ? null : n;
  };

  const now = new Date();
  const runDate = new Intl.DateTimeFormat('en-US', {
    timeZone: 'America/Los_Angeles',
    month: '2-digit',
    day: '2-digit',
    year: 'numeric',
  }).format(now);
  const runTime = new Intl.DateTimeFormat('en-US', {
    timeZone: 'America/Los_Angeles',
    hour: '2-digit',
    minute: '2-digit',
    hour12: true,
    timeZoneName: 'short',
  }).format(now);

  // Append rows after last row
  for (const rec of records) {
    const rowValues: any[] = [];
    rowValues[headerIndex['DATE']] = runDate;
    rowValues[headerIndex['TIME']] = runTime;
    rowValues[headerIndex['Item Name']] = rec.productName ?? null;
    rowValues[headerIndex['SKU']] = rec.sku ?? null;
    rowValues[headerIndex['DESCRIPTION']] = rec.description ?? null;
    rowValues[headerIndex['MPN']] = rec.mpn ?? null;
    if (headerIndex['MANUFACTURER']) {
      rowValues[headerIndex['MANUFACTURER']] = rec.manufacturer ?? null;
    }
    const priceNumber = parsePriceToNumber(rec.price);
    rowValues[headerIndex['PRICE']] = priceNumber ?? rec.price ?? null;
    const stockNumber = parseStockToNumber(rec.stock);
    rowValues[headerIndex['STOCK']] = stockNumber ?? rec.stock ?? null;
    rowValues[headerIndex['PRODUCT URL']] = rec.productUrl ?? null;

    sheet.addRow(rowValues);
  }
  sheet.views = [{ state: 'frozen', ySplit: 1 }];

  await workbook.xlsx.writeFile(filePath);
  // eslint-disable-next-line no-console
  console.log(`Excel written: ${filePath}`);
  return filePath;
}

test('Scrape BLUESTAR pages 1-7 and export to Excel', async ({ page }) => {
  test.slow();
  test.setTimeout(20 * 60 * 1000);
  await ensureOutputDir();

  await loginIfNeeded(page);

  const allUrls = new Set<string>();
  for (let i = 1; i <= TOTAL_PAGES; i += 1) {
    const urls = await collectProductUrlsForPage(page, i);
    urls.forEach((u) => allUrls.add(u));
  }
  const productUrls = Array.from(allUrls);
  await saveJson('product_urls_all.json', productUrls);

  const results: ProductRecord[] = [];
  let processed = 0;
  for (const url of productUrls) {
    // eslint-disable-next-line no-console
    console.log(`Processing (${++processed}/${productUrls.length}): ${url}`);
    await page.goto(url, { waitUntil: 'domcontentloaded' });
    await page.waitForLoadState('networkidle');
    await page.locator('h1, h2.product-title, .product-details-info').first().waitFor({ timeout: 15000 }).catch(() => undefined);
    const record = await extractProductDetails(page);
    results.push(record);
  }

  await saveJson('products_all.json', results);
  await writeExcel(results, 'products_all.xlsx');

  expect(results.length).toBeGreaterThan(0);
});

