const CONFIG = {
  SHEETS: {
    DASHBOARD: 'Dashboard',
    TRIPS: 'Trips',
    SETTINGS: 'Settings',
    CALCULATIONS: 'Calculations'
  },
  FUEL_PRICE: 73,
  CONSUMPTION_PER_100KM: 11,
  DRIVER_TAX_RATE: 0.19
};

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('CRM Логистика')
    .addItem('Открыть Dashboard', 'openDashboard')
    .addItem('Добавить поездку', 'showAddTripModal')
    .addToUi();

  ensureStructure();
}

function doGet(e) {
  ensureStructure();
  const view = (e && e.parameter && e.parameter.view) || 'dashboard';
  const templateName = view === 'addTrip' ? 'addTrip' : 'dashboard';
  return HtmlService.createTemplateFromFile(templateName)
    .evaluate()
    .setTitle(view === 'addTrip' ? 'Добавить поездку' : 'CRM Dashboard');
}

function openDashboard() {
  ensureStructure();
  const html = HtmlService.createTemplateFromFile('dashboard')
    .evaluate()
    .setWidth(1180)
    .setHeight(760);

  SpreadsheetApp.getUi().showModalDialog(html, '🚛 CRM Dashboard');
}

function showAddTripModal() {
  const html = HtmlService.createTemplateFromFile('addTrip')
    .evaluate()
    .setWidth(680)
    .setHeight(760);

  SpreadsheetApp.getUi().showModalDialog(html, 'Добавить поездку');
}

function getAppBaseUrl() {
  const serviceUrl = ScriptApp.getService().getUrl();
  return serviceUrl || '';
}

function saveTrip(data) {
  ensureStructure();
  const calc = calculateTrip(data);
  const tripsSheet = SpreadsheetApp.getActive().getSheetByName(CONFIG.SHEETS.TRIPS);

  calc.cargos.forEach(cargo => {
    tripsSheet.appendRow([
      calc.date,
      calc.tripId,
      calc.tripName,
      cargo.cargoName,
      cargo.amount,
      calc.totalKm,
      calc.emptyKm,
      calc.loadedKm,
      cargo.paymentTypeLabel,
      cargo.fuelTotal,
      cargo.leasing,
      cargo.repair,
      cargo.driverNet,
      cargo.companyProfit,
      calc.hasBackhaul ? 'Да' : 'Нет'
    ]);
  });

  return {
    success: true,
    message: `Поездка сохранена. Добавлено грузов: ${calc.cargos.length}`,
    trip: calc,
    dashboard: getDashboardData()
  };
}

function calculateTrip(data) {
  const tripName = String(data.tripName || '').trim();
  const totalKm = Number(data.totalKm) || 0;
  const emptyKm = Number(data.emptyKm) || 0;
  const hasBackhaul = emptyKm <= 0;
  const cargos = Array.isArray(data.cargos) ? data.cargos : [];

  if (!tripName) throw new Error('Укажите название поездки');
  if (totalKm <= 0) throw new Error('Введите корректный общий километраж');
  if (emptyKm < 0 || emptyKm > totalKm) throw new Error('Пустой пробег должен быть от 0 до общего километража');
  if (!cargos.length) throw new Error('Добавьте минимум один груз');

  const normalizedCargos = cargos.map((cargo, idx) => {
    const amount = Number(cargo.amount) || 0;
    const paymentType = String(cargo.paymentType || 'cash').toLowerCase();
    const cargoName = String(cargo.cargoName || `Груз ${idx + 1}`).trim();

    if (amount <= 0) throw new Error(`Проверьте сумму у груза ${idx + 1}`);

    return {
      cargoName,
      amount,
      paymentType,
      paymentTypeLabel: paymentType === 'vat' ? 'С НДС' : 'Нал'
    };
  });

  const totalAmount = normalizedCargos.reduce((sum, cargo) => sum + cargo.amount, 0);
  const fuel100 = CONFIG.CONSUMPTION_PER_100KM * CONFIG.FUEL_PRICE;
  const fuel1km = fuel100 / 100;
  const totalFuel = totalKm * fuel1km;
  const loadedKm = Math.max(totalKm - emptyKm, 0);
  const tripId = Utilities.getUuid();

  const cargoCalculations = normalizedCargos.map(cargo => {
    const revenueShare = totalAmount > 0 ? cargo.amount / totalAmount : 0;
    const fuelTotal = totalFuel * revenueShare;
    const profitBase = cargo.amount - fuelTotal;

    const leasing = profitBase * 0.2;
    const repair = profitBase * 0.4;
    const driverGross = profitBase * 0.4;
    const driverTax = cargo.paymentType === 'vat' ? driverGross * CONFIG.DRIVER_TAX_RATE : 0;
    const driverNet = driverGross - driverTax;
    const companyProfit = leasing + repair;

    return {
      ...cargo,
      fuelTotal,
      profitBase,
      leasing,
      repair,
      driverGross,
      driverTax,
      driverNet,
      companyProfit
    };
  });

  return {
    date: new Date(),
    tripId,
    tripName,
    totalKm,
    emptyKm,
    loadedKm,
    hasBackhaul,
    fuel100,
    fuel1km,
    totalFuel,
    amount: totalAmount,
    cargos: cargoCalculations
  };
}

function getDashboardData(startDate, endDate) {
  ensureStructure();
  const tripsSheet = SpreadsheetApp.getActive().getSheetByName(CONFIG.SHEETS.TRIPS);
  const values = tripsSheet.getDataRange().getValues();

  if (values.length < 2) {
    return {
      totals: {
        totalProfit: 0,
        tripCount: 0,
        totalFuel: 0,
        totalDriverIncome: 0,
        totalCompanyProfit: 0,
        totalRevenue: 0
      },
      chart: [],
      trips: []
    };
  }

  const header = values[0];
  const rows = values.slice(1).filter(r => r[0]);

  const col = getColumnIndexMap(header);

  const start = startDate ? new Date(startDate) : null;
  const end = endDate ? new Date(endDate) : null;

  const filtered = rows.filter(row => {
    const d = new Date(row[0]);
    if (start && d < start) return false;
    if (end) {
      const endInclusive = new Date(end);
      endInclusive.setHours(23, 59, 59, 999);
      if (d > endInclusive) return false;
    }
    return true;
  });

  const tripIds = new Set();

  const totals = filtered.reduce((acc, row, idx) => {
    const fallbackTripKey = `${row[0]}-${idx}`;
    const tripId = row[col.tripId] || fallbackTripKey;
    tripIds.add(tripId);

    acc.totalRevenue += Number(row[col.amount]) || 0;
    acc.totalFuel += Number(row[col.fuel]) || 0;
    acc.totalDriverIncome += Number(row[col.driver]) || 0;
    acc.totalCompanyProfit += Number(row[col.companyProfit]) || 0;
    return acc;
  }, {
    totalProfit: 0,
    tripCount: 0,
    totalFuel: 0,
    totalDriverIncome: 0,
    totalCompanyProfit: 0,
    totalRevenue: 0
  });

  totals.tripCount = tripIds.size;
  totals.totalProfit = totals.totalCompanyProfit;

  const chartMap = {};
  filtered.forEach(row => {
    const key = Utilities.formatDate(new Date(row[0]), Session.getScriptTimeZone(), 'yyyy-MM-dd');
    if (!chartMap[key]) chartMap[key] = { revenue: 0, profit: 0 };
    chartMap[key].revenue += Number(row[col.amount]) || 0;
    chartMap[key].profit += Number(row[col.companyProfit]) || 0;
  });

  const chart = Object.keys(chartMap)
    .sort()
    .map(date => [date, chartMap[date].revenue, chartMap[date].profit]);

  return {
    header,
    totals,
    chart,
    trips: filtered
  };
}

function getColumnIndexMap(header) {
  const byName = {};
  header.forEach((name, idx) => {
    byName[String(name).trim()] = idx;
  });

  return {
    tripId: byName['ID поездки'] ?? 1,
    amount: byName['Сумма'] ?? 1,
    fuel: byName['Топливо'] ?? 4,
    driver: byName['Водитель'] ?? 7,
    companyProfit: byName['Прибыль компании'] ?? 8
  };
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function ensureStructure() {
  const ss = SpreadsheetApp.getActive();

  [CONFIG.SHEETS.DASHBOARD, CONFIG.SHEETS.TRIPS, CONFIG.SHEETS.SETTINGS, CONFIG.SHEETS.CALCULATIONS]
    .forEach(name => {
      if (!ss.getSheetByName(name)) ss.insertSheet(name);
    });

  const tripsSheet = ss.getSheetByName(CONFIG.SHEETS.TRIPS);
  const tripHeaders = [
    'Дата',
    'ID поездки',
    'Название поездки',
    'Название груза',
    'Сумма',
    'Общий км',
    'Пустой км',
    'Груженый км',
    'Тип оплаты',
    'Топливо',
    'Лизинг',
    'Ремонт',
    'Водитель',
    'Прибыль компании',
    'Обратный груз'
  ];

  if (tripsSheet.getLastRow() === 0) {
    tripsSheet.appendRow(tripHeaders);
  }

  const settingsSheet = ss.getSheetByName(CONFIG.SHEETS.SETTINGS);
  if (settingsSheet.getLastRow() === 0) {
    settingsSheet.getRange(1, 1, 4, 2).setValues([
      ['Параметр', 'Значение'],
      ['Цена дизеля, руб/л', CONFIG.FUEL_PRICE],
      ['Расход, л/100км', CONFIG.CONSUMPTION_PER_100KM],
      ['Налог водителя', CONFIG.DRIVER_TAX_RATE]
    ]);
  }
}
