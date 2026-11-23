const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const { test } = require('node:test');

function createApp() {
    const elements = new Map();
    const context = vm.createContext({
        URLSearchParams,
        console: { log() {}, warn() {}, error() {} },
        window: { location: { search: '?disableNameMapAutoload' } },
        document: {
            readyState: 'complete',
            getElementById(id) {
                if (!elements.has(id)) elements.set(id, { addEventListener() {} });
                return elements.get(id);
            }
        }
    });
    const html = fs.readFileSync(path.join(__dirname, '..', 'phonesplit.html'), 'utf8');
    vm.runInContext(html.match(/<script>\s*([\s\S]*?)<\/script>/)[1], context);
    return context;
}

function inventory(overrides = {}) {
    return {
        SKU: 'AP-IPH16P-128-JPN-W', Manufacturer: 'Apple',
        Model: 'iPhone 16 Pro', Capacity: '128GB', Grade: 'W',
        location_code: 'HK_TOTEM_B2B', Category: 'Smartphones',
        qty_available: 10, floor_price: 321, currency: 'USD', ...overrides
    };
}

function pricing(overrides = {}) {
    return {
        Family: 'iPhone 16 Pro', Model: 'iPhone 16 Pro 128GB', Key: 'M1',
        'US (USD)': 525, 'ROW (USD)': 630,
        'EU MVA (EUR)': 660, 'UK MVA (GBP)': 545, ...overrides
    };
}

function calculate(app, stock, prices) {
    return app.calculateUsdPrice(app.processInventoryData(stock), app.processFloorUsdData(prices));
}

test('uses US (USD) for W stock in every output region', () => {
    const app = createApp();
    const locations = [
        'AU_B2B', 'HK_TOTEM_B2B', 'HK_IE_FSET_B2B', 'SCHJP', 'US_B2B',
        'UK_MRC_MVA_B2B', 'UK_MRC_VAT_B2B', 'IE_IE_MVA_B2B',
        'IE_IE_VAT_B2B', 'DUB_B2B', 'CA_B2B', 'HK_HTR_B2B_VZW'
    ];
    const result = calculate(app, locations.map(location_code => inventory({ location_code })), [pricing()]);
    assert.equal(result.length, 10);
    assert.equal(result.reduce((sum, row) => sum + row.Qty, 0), locations.length * 10);
    for (const row of result) {
        assert.equal(row.USD, 525, row.Country);
        assert.equal(row.USDSource, 'floor-file', row.Country);
    }
});

test('new workbook prices never overwrite another grade', () => {
    const app = createApp();
    for (const Grade of ['A', 'B', 'C', 'Z', 'AS', 'CP', 'WP', 'NB']) {
        const [row] = calculate(app, [inventory({ Grade })], [pricing()]);
        assert.equal(row.USD, 321, Grade);
        assert.equal(row.USDSource, 'inventory', Grade);
    }
    for (const Grade of ['D', 'DX', 'F']) {
        const [row] = calculate(app, [inventory({ Grade, location_code: 'HK_IE_FSET_B2B' })], [pricing()]);
        assert.equal(row.USD, 321, Grade);
    }
});

test('unmatched models retain their inventory floor price', () => {
    const app = createApp();
    const [row] = calculate(app, [inventory({ Model: 'iPhone 16 Pro Max' })], [pricing()]);
    assert.equal(row.USD, 321);
    assert.equal(row.USDSource, 'inventory');
});

test('normalises case, spaces, capacity units and explicit SE year aliases', () => {
    const app = createApp();
    const [row] = calculate(app, [inventory({ Model: 'IPHONE  16 PRO', Capacity: '128 GB' })], [pricing()]);
    assert.equal(row.USD, 525);
    for (const [year, generation] of [[2020, '2nd'], [2022, '3rd']]) {
        const [se] = calculate(app, [inventory({ Model: `iPhone SE (${year})` })], [pricing({ Model: `iPhone SE ${generation} Gen 128GB` })]);
        assert.equal(se.USD, 525);
    }
});

test('does not confuse model variants or capacities', () => {
    const app = createApp();
    for (const change of [{ Model: 'iPhone 16' }, { Model: 'iPhone 16 Pro Max' }, { Capacity: '256GB' }]) {
        assert.equal(calculate(app, [inventory(change)], [pricing()])[0].USD, 321);
    }
});

test('keeps old floor formats with grade-specific pricing and accepts Model grade', () => {
    const app = createApp();
    for (const header of ['Model Grade', 'Model grade', 'model grade', 'ModelGrade']) {
        for (const priceColumn of ['>500', '500', 'price']) {
            const [row] = calculate(app, [inventory({ Grade: 'B' })], [{
                [header]: 'iPhone 16 Pro 128GB-B', Grade: 'B', [priceColumn]: 234
            }]);
            assert.equal(row.USD, 234, `${header} / ${priceColumn}`);
        }
    }
});

test('accepts currency and thousands formatting and normalised headers', () => {
    const app = createApp();
    const prices = [{ ' Model ': 'iPhone 16 Pro 128GB', ' US (USD) ': '$1,300.50' }];
    assert.equal(calculate(app, [inventory()], prices)[0].USD, 1300.5);
});

test('keeps the legacy >500 price priority when multiple price columns exist', () => {
    const app = createApp();
    const [row] = calculate(app, [inventory()], [{
        'Model Grade': 'iPhone 16 Pro 128GB', Grade: 'W', price: 100, '500': 200, '>500': 300
    }]);
    assert.equal(row.USD, 300);
});

test('fails clearly on unsupported, empty or invalid pricing instead of silently ignoring it', () => {
    const app = createApp();
    assert.throws(() => app.processFloorUsdData([]), /no data rows/);
    assert.throws(() => app.processFloorUsdData([{ Model: 'iPhone 16 Pro', Price: 500 }]), /Unrecognised/);
    for (const price of ['', 'unknown', '-5', '52,5', 'Infinity']) {
        assert.throws(() => app.processFloorUsdData([pricing({ 'US (USD)': price })]), /invalid USD price/);
    }
});

test('rejects conflicting prices but consolidates identical duplicate models', () => {
    const app = createApp();
    assert.throws(() => app.processFloorUsdData([pricing(), pricing({ 'US (USD)': 526 })]), /Conflicting/);
    assert.equal(app.processFloorUsdData([pricing(), pricing()]).length, 1);
});

test('an explicit zero price is not replaced by the inventory price', () => {
    const app = createApp();
    const [row] = calculate(app, [inventory()], [pricing({ 'US (USD)': 0 })]);
    assert.equal(row.USD, 0);
    assert.equal(row.USDSource, 'floor-file');
});

test('FSET naming and filtered output preserve the selected floor price', () => {
    const app = createApp();
    app.applyLoadedNameMap([{ sku: 'AP-IPH14PM-128-JPN', en_name: 'iPhone 14 Pro Max 128GB' }], 'test-map');
    const stock = [inventory({ Model: 'iPhone 14 Pro Max', SKU: 'AP-IPH14PM-128-JPN-W', location_code: 'HK_IE_FSET_B2B' })];
    const result = calculate(app, stock, [pricing({ Model: 'iPhone 14 Pro Max 128GB', 'US (USD)': 375 })]);
    assert.equal(result[0]['Model grade'], 'iPhone 14 Pro Max 128GB-W');
    assert.equal(result[0].USD, 375);
    assert.equal(app.createSecondFilteredOutputData(result)[0].USD, 375);
});
