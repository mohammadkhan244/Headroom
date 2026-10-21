#!/usr/bin/env node
// Seed Headroom Balances + Config via the proxy.
// Usage: HEADROOM_PASSCODE=xxx node scripts/seed.mjs [base-url]
// base-url defaults to http://localhost:3000

import { readFileSync } from 'fs';
import { fileURLToPath } from 'url';
import { dirname, join } from 'path';

const __dir  = dirname(fileURLToPath(import.meta.url));
const seed   = JSON.parse(readFileSync(join(__dir, '../seed/seed.json'), 'utf8'));
const base   = (process.argv[2] || 'http://localhost:3000').replace(/\/$/, '');
const pc     = process.env.HEADROOM_PASSCODE;

if (!pc) { console.error('Set HEADROOM_PASSCODE env var'); process.exit(1); }

const headers = { 'Content-Type': 'application/json', 'x-passcode': pc };

async function post(body) {
  const r = await fetch(`${base}/api/sheet`, { method: 'POST', headers, body: JSON.stringify(body) });
  const j = await r.json();
  if (!j.ok) throw new Error(JSON.stringify(j));
  return j;
}

async function main() {
  console.log(`Seeding ${base}…`);

  // Upsert balances
  for (const row of seed.balances) {
    const b = { ...row, updatedAt: Date.now() };
    await post({ op: 'upsert', sheet: 'Balances', row: b });
    console.log(`  balance  ${b.id}`);
  }

  // Upsert config keys
  for (const [key, value] of Object.entries(seed.config)) {
    await post({ op: 'config_set', sheet: 'Config', key, value });
    console.log(`  config   ${key}`);
  }

  console.log('Done.');
}

main().catch(e => { console.error(e); process.exit(1); });
