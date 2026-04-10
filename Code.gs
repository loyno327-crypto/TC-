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

function openDashboard() {
  ensureStructure();
  const html = HtmlService.createTemplateFromFile('dashboard')
    .evaluate()
    .setTitle('🚛 CRM Dashboard')
    .setWidth(420);

  SpreadsheetApp.getUi().showSidebar(html);
}

function showAddTripModal() {
  const html = HtmlService.createTemplateFromFile('addTrip')
    .evaluate()
    .setWidth(520)
    .setHeight(640);

  SpreadsheetApp.getUi().showModalDialog(html, 'Добавить поездку');
}

function saveTrip(data) {
  ensureStructure();
  const calc = calculateTrip(data);
  const tripsSheet = SpreadsheetApp.getActive().getSheetByName(CONFIG.SHEETS.TRIPS);

  tripsSheet.appendRow([
    new Date(),
    calc.amount,
    calc.km,
    calc.paymentType,
    calc.fuelTotal,
    calc.leasing,
    calc.repair,
    calc.driverNet,
    calc.companyProfit,
    calc.hasBackhaul ? 'Да' : 'Нет'
  ]);

  return {
    success: true,
    message: 'Поездка сохранена успешно',
    trip: calc,
    dashboard: getDashboardData()
  };
}

function calculateTrip(data) {
  const amount = Number(data.amount) || 0;
  const km = Number(data.km) || 0;
  const paymentType = String(data.paymentType || 'cash').toLowerCase();
  const hasBackhaul = String(data.hasBackhaul || 'yes').toLowerCase() === 'yes';

  const fuel100 = CONFIG.CONSUMPTION_PER_100KM * CONFIG.FUEL_PRICE;
  const fuel1km = fuel100 / 100;
  const fuelTotal = km * fuel1km;

  const profitBase = amount - fuelTotal;
  const leasing = profitBase * 0.2;
  const repair = profitBase * 0.4;

  let driverGross = profitBase * 0.4;
  let driverTax = 0;
  let driverNet = driverGross;

  if (paymentType === 'vat') {
    driverTax = driverGross * CONFIG.DRIVER_TAX_RATE;
    driverNet = driverGross - driverTax;
  }

  let companyProfit = leasing + repair;

  if (!hasBackhaul) {
    const extraFuel = (km * fuel1km) / 2;
    driverNet -= extraFuel / 2;
    companyProfit -= extraFuel / 2;
  }

  return {
    date: new Date(),
    amount,
    km,
    paymentType: paymentType === 'vat' ? 'С НДС' : 'Нал',
    hasBackhaul,
    fuel100,
    fuel1km,
    fuelTotal,
    profitBase,
    leasing,
    repair,
    driverGross,
    driverTax,
    driverNet,
    companyProfit
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

  const totals = filtered.reduce((acc, row) => {
    acc.tripCount += 1;
    acc.totalRevenue += Number(row[1]) || 0;
    acc.totalFuel += Number(row[4]) || 0;
    acc.totalDriverIncome += Number(row[7]) || 0;
    acc.totalCompanyProfit += Number(row[8]) || 0;
    return acc;
  }, {
    totalProfit: 0,
    tripCount: 0,
    totalFuel: 0,
    totalDriverIncome: 0,
    totalCompanyProfit: 0,
    totalRevenue: 0
  });

  totals.totalProfit = totals.totalCompanyProfit;

  const chartMap = {};
  filtered.forEach(row => {
    const key = Utilities.formatDate(new Date(row[0]), Session.getScriptTimeZone(), 'yyyy-MM-dd');
    if (!chartMap[key]) chartMap[key] = { revenue: 0, profit: 0 };
    chartMap[key].revenue += Number(row[1]) || 0;
    chartMap[key].profit += Number(row[8]) || 0;
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
  const tripHeaders = ['Дата', 'Сумма', 'Км', 'Тип оплаты', 'Топливо', 'Лизинг', 'Ремонт', 'Водитель', 'Прибыль компании', 'Обратный груз'];
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
