#!/usr/bin/env node
const { chromium } = require('playwright');
const path = require('path');
const fs = require('fs');

const HTML = path.resolve('/home/ouyan/llm/dataagent/docs/prototype.html');
const OUT   = path.resolve('/home/ouyan/llm/dataagent/docs/screenshots');
const pages = ['sql', 'analysis', 'catalog', 'governance', 'pipeline'];
const names = {
  sql:        '01-智能SQL工作台',
  analysis:   '02-数据分析中心',
  catalog:    '03-数据目录',
  governance: '04-数据治理-血缘图谱',
  pipeline:   '05-管道编排',
};

(async () => {
  if (!fs.existsSync(OUT)) fs.mkdirSync(OUT, { recursive: true });

  const browser = await chromium.launch();
  const ctx = await browser.newContext({ viewport: { width: 1440, height: 900 } });
  const page = await ctx.newPage();

  await page.goto('file://' + HTML);

  for (const p of pages) {
    // click nav item
    await page.click(`.nav-item[onclick="switchPage('${p}')"]`);
    await page.waitForTimeout(400);
    const file = path.join(OUT, `${names[p]}.png`);
    await page.screenshot({ path: file, fullPage: false });
    console.log('Saved:', file);
  }

  await browser.close();
  console.log('Done. All screenshots saved to:', OUT);
})();
