const CONFIG = {
  SHEETS: {
    TRIPS: 'Trips',
    SETTINGS: 'Settings',
    TRIPS_ALIASES: ['Trips', 'TRIPS', 'Поездки'],
    SETTINGS_ALIASES: ['Settings', 'SETTINGS', 'Настройки']
  },
  DEFAULTS: {
    fuelPrice: 73,
    fuelConsumption100: 11,
    driverTaxRate: 0.19,
    leasingShare: 0.2,
    repairShare: 0.4,
    driverShare: 0.4
  }
};

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Логистика')
    .addItem('Заявки', 'openRequestsPanel')
    .addItem('Внести поездку', 'openAddTripDialog')
    .addToUi();

  ensureStructure();
}

function doGet(e) {
  ensureStructure();
  const view = (e && e.parameter && e.parameter.view) || 'dashboard';
  const file = view === 'addTrip' ? 'addTrip' : 'dashboard';
  return HtmlService.createTemplateFromFile(file)
    .evaluate()
    .setTitle(view === 'addTrip' ? 'Внести поездку' : 'Панель заявок');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function openRequestsPanel() {
  const html = HtmlService.createTemplateFromFile('dashboard').evaluate().setWidth(1280).setHeight(760);
  SpreadsheetApp.getUi().showModelessDialog(html, 'Панель "Заявки"');
}

function openAddTripDialog() {
  const html = HtmlService.createTemplateFromFile('addTrip').evaluate().setWidth(760).setHeight(760);
  SpreadsheetApp.getUi().showModalDialog(html, 'Внести поездку');
}

function getAppBaseUrl() {
  return ScriptApp.getService().getUrl() || '';
}

function getTripsHeader() {
  return [
    'Дата',
    'ID поездки',
    'Основной маршрут',
    'Маршрут груза',
    'Сумма',
    'Тип оплаты',
    'Общий км',
    'Холостой км',
    'Себестоимость топлива',
    'Остаток после топлива',
    'Лизинг',
    'Ремонт',
    'Водитель начислено',
    'Налог водителя',
    'Водитель к выплате',
    'Корректировка (холостой км)',
    'Компании после корректировки'
  ];
}

function ensureStructure() {
  const ss = SpreadsheetApp.getActive();
  ensureSheetExists(ss, CONFIG.SHEETS.TRIPS, CONFIG.SHEETS.TRIPS_ALIASES);
  ensureSheetExists(ss, CONFIG.SHEETS.SETTINGS, CONFIG.SHEETS.SETTINGS_ALIASES);

  const tripsSheet = getSheetByAliases(ss, CONFIG.SHEETS.TRIPS, CONFIG.SHEETS.TRIPS_ALIASES);
  const tripHeaders = getTripsHeader();

  if (tripsSheet.getLastRow() === 0) {
    tripsSheet.appendRow(tripHeaders);
  } else {
    const currentHeader = tripsSheet.getRange(1, 1, 1, Math.max(tripsSheet.getLastColumn(), tripHeaders.length)).getValues()[0];
    const normalized = tripHeaders.every((expected, i) => String(currentHeader[i] || '').trim() === expected);
    if (!normalized) {
      tripsSheet.getRange(1, 1, 1, tripHeaders.length).setValues([tripHeaders]);
    }
  }

  const settingsSheet = getSheetByAliases(ss, CONFIG.SHEETS.SETTINGS, CONFIG.SHEETS.SETTINGS_ALIASES);
  if (settingsSheet.getLastRow() === 0) {
    settingsSheet.getRange(1, 1, 7, 2).setValues([
      ['Параметр', 'Значение'],
      ['Цена дизеля, руб/л', CONFIG.DEFAULTS.fuelPrice],
      ['Расход, л/100км', CONFIG.DEFAULTS.fuelConsumption100],
      ['Налог водителя (только С НДС)', CONFIG.DEFAULTS.driverTaxRate],
      ['Лизинг, доля', CONFIG.DEFAULTS.leasingShare],
      ['Ремонт, доля', CONFIG.DEFAULTS.repairShare],
      ['Водитель, доля', CONFIG.DEFAULTS.driverShare]
    ]);
  }
}

function getSettings() {
  ensureStructure();
  const sh = getSheetByAliases(
    SpreadsheetApp.getActive(),
    CONFIG.SHEETS.SETTINGS,
    CONFIG.SHEETS.SETTINGS_ALIASES
  );
  const rows = sh.getDataRange().getValues();
  const map = {};
  rows.slice(1).forEach((r) => {
    map[String(r[0]).trim()] = Number(r[1]);
  });

  return {
    fuelPrice: map['Цена дизеля, руб/л'] || CONFIG.DEFAULTS.fuelPrice,
    fuelConsumption100: map['Расход, л/100км'] || CONFIG.DEFAULTS.fuelConsumption100,
    driverTaxRate: map['Налог водителя (только С НДС)'] || CONFIG.DEFAULTS.driverTaxRate,
    leasingShare: map['Лизинг, доля'] || CONFIG.DEFAULTS.leasingShare,
    repairShare: map['Ремонт, доля'] || CONFIG.DEFAULTS.repairShare,
    driverShare: map['Водитель, доля'] || CONFIG.DEFAULTS.driverShare
  };
}

function saveTrip(payload) {
  ensureStructure();
  const trip = calculateTrip(payload);
  const sh = getSheetByAliases(
    SpreadsheetApp.getActive(),
    CONFIG.SHEETS.TRIPS,
    CONFIG.SHEETS.TRIPS_ALIASES
  );

  trip.cargos.forEach((cargo) => {
    sh.appendRow([
      trip.date,
      trip.tripId,
      trip.mainRoute,
      cargo.route,
      cargo.amount,
      cargo.paymentLabel,
      trip.totalKm,
      trip.emptyKm,
      cargo.fuelCost,
      cargo.netAfterFuel,
      cargo.leasing,
      cargo.repair,
      cargo.driverGross,
      cargo.driverTax,
      cargo.driverNet,
      cargo.emptyCorrection,
      cargo.companyNet
    ]);
  });

  return { success: true, message: 'Поездка внесена', tripId: trip.tripId };
}

function calculateTrip(payload) {
  const settings = getSettings();
  const mainRoute = String(payload.mainRoute || '').trim();
  const totalKm = Number(payload.totalKm);
  const emptyKm = Number(payload.emptyKm || 0);
  const cargos = Array.isArray(payload.cargos) ? payload.cargos : [];

  if (!mainRoute) throw new Error('Укажите основной маршрут поездки');
  if (!(totalKm > 0)) throw new Error('Укажите корректный общий километраж');
  if (emptyKm < 0 || emptyKm > totalKm) throw new Error('Холостой км должен быть от 0 до общего км');
  if (!cargos.length) throw new Error('Добавьте хотя бы один маршрут груза');

  const normCargos = cargos.map((c, idx) => {
    const route = String(c.route || '').trim();
    const amount = Number(c.amount);
    const paymentType = String(c.paymentType || 'cash').toLowerCase();
    if (!route) throw new Error(`Маршрут груза ${idx + 1} не заполнен`);
    if (!(amount > 0)) throw new Error(`Сумма у груза ${idx + 1} должна быть больше 0`);
    if (!['cash', 'no_vat', 'vat'].includes(paymentType)) throw new Error('Некорректный тип оплаты');
    return {
      route,
      amount,
      paymentType,
      paymentLabel: paymentType === 'cash' ? 'Наличные' : paymentType === 'no_vat' ? 'Без НДС' : 'С НДС'
    };
  });

  const totalAmount = normCargos.reduce((s, c) => s + c.amount, 0);
  const fuelPerKm = (settings.fuelPrice * settings.fuelConsumption100) / 100;
  const fuelTotal = totalKm * fuelPerKm;
  const emptyFuelTotal = emptyKm * fuelPerKm;

  const calculated = normCargos.map((cargo) => {
    const share = cargo.amount / totalAmount;
    const fuelCost = fuelTotal * share;
    const netAfterFuel = cargo.amount - fuelCost;
    const leasing = netAfterFuel * settings.leasingShare;
    const repair = netAfterFuel * settings.repairShare;
    const driverGross = netAfterFuel * settings.driverShare;
    const driverTax = cargo.paymentType === 'vat' ? driverGross * settings.driverTaxRate : 0;
    const driverNetBeforeEmpty = driverGross - driverTax;

    const emptyCorrection = emptyFuelTotal * share;
    const companyHalfCorrection = emptyCorrection / 2;
    const driverHalfCorrection = emptyCorrection / 2;

    return {
      ...cargo,
      fuelCost,
      netAfterFuel,
      leasing,
      repair,
      driverGross,
      driverTax,
      driverNet: Math.max(driverNetBeforeEmpty - driverHalfCorrection, 0),
      emptyCorrection,
      companyNet: leasing + repair - companyHalfCorrection
    };
  });

  return {
    date: new Date(),
    tripId: Utilities.getUuid(),
    mainRoute,
    totalKm,
    emptyKm,
    fuelPerKm,
    totalAmount,
    cargos: calculated
  };
}

function getRequestsPanelData() {
  ensureStructure();
  const sh = getSheetByAliases(
    SpreadsheetApp.getActive(),
    CONFIG.SHEETS.TRIPS,
    CONFIG.SHEETS.TRIPS_ALIASES
  );
  const rows = sh.getDataRange().getValues();
  if (rows.length <= 1) {
    return {
      trips: [],
      moneyFlow: { revenue: 0, fuel: 0, leasing: 0, repair: 0, driverNet: 0, driverTax: 0, companyNet: 0 }
    };
  }

  const header = rows[0];
  const idx = createIndex(header);
  const grouped = {};

  rows.slice(1).forEach((r, i) => {
    const tripId = resolveTripKey(r, idx, i + 2);
    if (!grouped[tripId]) {
      grouped[tripId] = {
        tripId,
        date: r[idx.date],
        mainRoute: r[idx.mainRoute] || '—',
        cargoRoutes: [],
        totalKm: Number(r[idx.totalKm]) || 0,
        emptyKm: Number(r[idx.emptyKm]) || 0,
        totalRevenue: 0,
        fuel: 0,
        leasing: 0,
        repair: 0,
        driverNet: 0,
        driverTax: 0,
        companyNet: 0,
        cargoCount: 0
      };
    }

    grouped[tripId].totalRevenue += Number(r[idx.amount]) || 0;
    grouped[tripId].fuel += Number(r[idx.fuel]) || 0;
    grouped[tripId].leasing += Number(r[idx.leasing]) || 0;
    grouped[tripId].repair += Number(r[idx.repair]) || 0;
    grouped[tripId].driverNet += Number(r[idx.driverNet]) || 0;
    grouped[tripId].driverTax += Number(r[idx.driverTax]) || 0;
    grouped[tripId].companyNet += Number(r[idx.companyNet]) || 0;
    grouped[tripId].cargoCount += 1;
    if (r[idx.cargoRoute]) grouped[tripId].cargoRoutes.push(String(r[idx.cargoRoute]));
  });

  const trips = Object.values(grouped).sort((a, b) => new Date(b.date) - new Date(a.date));

  const moneyFlow = trips.reduce((acc, t) => {
    acc.revenue += t.totalRevenue;
    acc.fuel += t.fuel;
    acc.leasing += t.leasing;
    acc.repair += t.repair;
    acc.driverNet += t.driverNet;
    acc.driverTax += t.driverTax;
    acc.companyNet += t.companyNet;
    return acc;
  }, { revenue: 0, fuel: 0, leasing: 0, repair: 0, driverNet: 0, driverTax: 0, companyNet: 0 });

  return { trips, moneyFlow };
}

function getTripDetails(tripId) {
  ensureStructure();
  if (!tripId) throw new Error('Не указан ID поездки');

  const sh = getSheetByAliases(
    SpreadsheetApp.getActive(),
    CONFIG.SHEETS.TRIPS,
    CONFIG.SHEETS.TRIPS_ALIASES
  );
  const rows = sh.getDataRange().getValues();
  const header = rows[0];
  const idx = createIndex(header);
  const items = rows.slice(1).filter((r, i) => resolveTripKey(r, idx, i + 2) === String(tripId));
  if (!items.length) throw new Error('Поездка не найдена');

  const first = items[0];
  const detail = {
    tripId,
    date: first[idx.date],
    mainRoute: first[idx.mainRoute],
    totalKm: Number(first[idx.totalKm]) || 0,
    emptyKm: Number(first[idx.emptyKm]) || 0,
    cargos: items.map((r) => ({
      route: r[idx.cargoRoute],
      amount: Number(r[idx.amount]) || 0,
      paymentType: r[idx.paymentType],
      fuel: Number(r[idx.fuel]) || 0,
      netAfterFuel: Number(r[idx.netAfterFuel]) || 0,
      leasing: Number(r[idx.leasing]) || 0,
      repair: Number(r[idx.repair]) || 0,
      driverGross: Number(r[idx.driverGross]) || 0,
      driverTax: Number(r[idx.driverTax]) || 0,
      driverNet: Number(r[idx.driverNet]) || 0,
      emptyCorrection: Number(r[idx.emptyCorrection]) || 0,
      companyNet: Number(r[idx.companyNet]) || 0
    }))
  };

  detail.totals = detail.cargos.reduce((acc, c) => {
    acc.amount += c.amount;
    acc.fuel += c.fuel;
    acc.netAfterFuel += c.netAfterFuel;
    acc.leasing += c.leasing;
    acc.repair += c.repair;
    acc.driverGross += c.driverGross;
    acc.driverTax += c.driverTax;
    acc.driverNet += c.driverNet;
    acc.emptyCorrection += c.emptyCorrection;
    acc.companyNet += c.companyNet;
    return acc;
  }, { amount: 0, fuel: 0, netAfterFuel: 0, leasing: 0, repair: 0, driverGross: 0, driverTax: 0, driverNet: 0, emptyCorrection: 0, companyNet: 0 });

  return detail;
}


function resolveTripKey(row, idx, rowNumber) {
  const rawTripId = idx.tripId !== undefined ? row[idx.tripId] : '';
  if (rawTripId) return String(rawTripId);

  const date = idx.date !== undefined ? row[idx.date] : '';
  const mainRoute = idx.mainRoute !== undefined ? row[idx.mainRoute] : '';
  const totalKm = idx.totalKm !== undefined ? row[idx.totalKm] : '';
  const emptyKm = idx.emptyKm !== undefined ? row[idx.emptyKm] : '';
  const composite = [date, mainRoute, totalKm, emptyKm].map((v) => String(v || '').trim()).join('|');

  return composite.replace(/^\|+|\|+$/g, '') || `row-${rowNumber}`;
}

function createIndex(header) {
  const by = {};
  header.forEach((h, i) => by[String(h).trim()] = i);

  const pick = (...names) => {
    for (let i = 0; i < names.length; i++) {
      const idx = by[names[i]];
      if (typeof idx === 'number') return idx;
    }
    return undefined;
  };

  return {
    date: pick('Дата', 'Date'),
    tripId: pick('ID поездки', 'ID', 'Trip ID'),
    mainRoute: pick('Основной маршрут', 'Маршрут', 'Main route'),
    cargoRoute: pick('Маршрут груза', 'Груз', 'Cargo route'),
    amount: pick('Сумма', 'Стоимость', 'Amount'),
    paymentType: pick('Тип оплаты', 'Оплата', 'Payment type'),
    totalKm: pick('Общий км', 'Км', 'Total km'),
    emptyKm: pick('Холостой км', 'Пустой км', 'Empty km'),
    fuel: pick('Себестоимость топлива', 'Топливо', 'Fuel cost'),
    netAfterFuel: pick('Остаток после топлива', 'После топлива', 'Net after fuel'),
    leasing: pick('Лизинг', 'Leasing'),
    repair: pick('Ремонт', 'Repair'),
    driverGross: pick('Водитель начислено', 'Водитель (грязными)', 'Driver gross'),
    driverTax: pick('Налог водителя', 'Налог', 'Driver tax'),
    driverNet: pick('Водитель к выплате', 'Водитель (чистыми)', 'Driver net'),
    emptyCorrection: pick('Корректировка (холостой км)', 'Корректировка', 'Empty correction'),
    companyNet: pick('Компании после корректировки', 'Компании', 'Company net')
  };
}

function ensureSheetExists(ss, primaryName, aliases) {
  const found = getSheetByAliases(ss, primaryName, aliases || []);
  if (found) return found;
  return ss.insertSheet(primaryName);
}

function getSheetByAliases(ss, primaryName, aliases) {
  const names = [primaryName].concat(Array.isArray(aliases) ? aliases : []);
  for (let i = 0; i < names.length; i++) {
    const direct = ss.getSheetByName(names[i]);
    if (direct) return direct;
  }

  const normalizedAliases = names.map(normalizeSheetName).filter(Boolean);
  const allSheets = ss.getSheets();
  for (let i = 0; i < allSheets.length; i++) {
    const name = allSheets[i].getName();
    if (normalizedAliases.includes(normalizeSheetName(name))) return allSheets[i];
  }

  return null;
}

function normalizeSheetName(name) {
  return String(name || '').trim().toLowerCase();
}
