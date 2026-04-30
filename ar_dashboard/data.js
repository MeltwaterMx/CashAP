// Este es tu archivo de configuración de datos.
// Puedes editar estos valores para que el dashboard muestre información real o actualizada.
// Solo asegúrate de respetar el formato (comillas, comas, etc).

const DASHBOARD_DATA = {
    // --- MÉTRICAS DSO (Days Sales Outstanding) ---
    dso: {
        actual: 38.4,   // DSO del mes actual
        prev: 36.1,     // DSO del mes anterior
        best: 24.0,     // Mejor DSO posible dadas las condiciones actuales
        target: 36.3    // Objetivo fijado para el DSO
    },

    // --- REPORTE DE ANTIGÜEDAD (Aging) - Valores en Moneda ---
    aging: {
        current: 1842000, // Deuda al Corriente (No Vencida)
        d30: 987500,      // Vencido 1 a 30 días
        d60: 643200,      // Vencido 31 a 60 días
        d90: 471350,      // Vencido 61 a 90 días
        d90p: 343300      // Vencido más de 90 días
    },

    // --- MÉTRICAS DE RESUMEN EJECUTIVO ---
    totalAR: 4287350,   // Cartera Total ($) - Cuentas por Cobrar Total
    collected: 1923100, // Lo recaudado en el mes hasta el momento ($) MTD

    // --- GRÁFICA TENDENCIA DSO (Últimos 6 meses) ---
    months: ['Oct', 'Nov', 'Dic', 'Ene', 'Feb', 'Mar'], // Etiquetas de los meses
    dsoHistory: [34.2, 35.8, 33.9, 37.1, 36.1, 38.4], // Valores reales históricos de DSO

    // --- LISTA DE CLIENTES ---
    // name: Nombre, balance: Deuda Total, overdue: Días Vencido, limitExc: Límite de Crédito Excedido %
    // trend: Tendencia (up: Deteriorando, down: Mejorando, stable: Estable)
    // score: Calificación de Riesgo (0-100), seg: Segmentación (strategic, alert, stable, lowrisk)
    clients: [
        { name: 'Constructora Alfa', balance: 487200, overdue: 91, limitExc: 34, trend: 'up', score: 91, seg: 'alert' },
        { name: 'Tech Solutions S.A.', balance: 312500, overdue: 67, limitExc: 0, trend: 'up', score: 78, seg: 'alert' },
        { name: 'Distribuidora Beta', balance: 198400, overdue: 78, limitExc: 22, trend: 'up', score: 85, seg: 'lowrisk' },
        { name: 'Grupo Inversiones C.', balance: 156800, overdue: 45, limitExc: 0, trend: 'stable', score: 52, seg: 'stable' },
        { name: 'Importaciones Delta', balance: 134200, overdue: 93, limitExc: 18, trend: 'up', score: 88, seg: 'lowrisk' },
        { name: 'Metalúrgica Norte', balance: 98700, overdue: 31, limitExc: 0, trend: 'down', score: 32, seg: 'stable' },
        { name: 'Farmacorp Ltda.', balance: 87600, overdue: 60, limitExc: 11, trend: 'stable', score: 61, seg: 'alert' },
        { name: 'Megasuper Retail', balance: 743000, overdue: 12, limitExc: 0, trend: 'down', score: 18, seg: 'strategic' },
        { name: 'Industrias Omega', balance: 621500, overdue: 8, limitExc: 0, trend: 'down', score: 12, seg: 'strategic' },
        { name: 'Servicios Globales', balance: 412000, overdue: 22, limitExc: 0, trend: 'stable', score: 28, seg: 'strategic' },
        { name: 'Banco Regional', balance: 312000, overdue: 5, limitExc: 0, trend: 'down', score: 9, seg: 'strategic' },
        { name: 'MicroEmpresas CR', balance: 87000, overdue: 14, limitExc: 0, trend: 'stable', score: 22, seg: 'stable' },
        { name: 'Panadería Central', balance: 42000, overdue: 36, limitExc: 8, trend: 'up', score: 58, seg: 'lowrisk' },
        { name: 'Auto Repuestos JM', balance: 63000, overdue: 88, limitExc: 15, trend: 'up', score: 80, seg: 'lowrisk' },
    ],

    // --- PROYECCIONES DE RECAUDO ---
    // client: Cliente, week: Fecha Estimada (Semana), amount: Monto esperado ($), prob: Probabilidad (high, med, low)
    projection: [
        { client: 'Megasuper Retail', week: 'Mar 04', amount: 284000, prob: 'high' },
        { client: 'Industrias Omega', week: 'Mar 06', amount: 210000, prob: 'high' },
        { client: 'Servicios Globales', week: 'Mar 07', amount: 156000, prob: 'high' },
        { client: 'Banco Regional', week: 'Mar 08', amount: 190000, prob: 'high' },
        { client: 'Metalúrgica Norte', week: 'Mar 10', amount: 98700, prob: 'high' },
        { client: 'Grupo Inversiones C.', week: 'Mar 12', amount: 156800, prob: 'high' },
        { client: 'MicroEmpresas CR', week: 'Mar 14', amount: 87000, prob: 'high' },
        { client: 'Farmacorp Ltda.', week: 'Mar 15', amount: 87600, prob: 'med' },
        { client: 'Tech Solutions S.A.', week: 'Mar 18', amount: 113000, prob: 'med' },
        { client: 'Constructora Alfa', week: 'Mar 20', amount: 172000, prob: 'med' },
        { client: 'Distribuidora Beta', week: 'Mar 22', amount: 134440, prob: 'med' },
        { client: 'Importaciones Delta', week: 'Mar 25', amount: 91510, prob: 'low' },
        { client: 'Panadería Central', week: 'Mar 26', amount: 42000, prob: 'low' },
        { client: 'Auto Repuestos JM', week: 'Mar 28', amount: 63000, prob: 'low' },
        { client: 'Constructora Alfa', week: 'Mar 31', amount: 77840, prob: 'low' },
    ],

    // --- CASH APPLICATIONS & REFUNDS ---
    cashapp: {
        kpis: { unapplied: 345000, suspense: 55000, autoMatch: 82.5, manTime: 4 },
        suspense: { noRef: 45, invalidAmt: 25, noClient: 20, doublePay: 10 },
        items: [
            { ref: 'TRF-9921', amount: 45000, date: '12 Feb 2026', days: 19, client: 'Industrias Omega?', status: 'Investigando' },
            { ref: 'DEP-0023', amount: 12500, date: '14 Feb 2026', days: 17, client: 'Desconocido', status: 'Pendiente' },
            { ref: 'TRF-9812', amount: 8900, date: '15 Feb 2026', days: 16, client: 'Tech Solutions S.A.', status: 'Contactado' },
            { ref: 'CHK-4412', amount: 5600, date: '15 Feb 2026', days: 16, client: 'Metalúrgica Norte', status: 'En Análisis' },
            { ref: 'TRF-1022', amount: 120000, date: '16 Feb 2026', days: 15, client: 'Megasuper Retail?', status: 'Pendiente' }
        ],
        refunds: {
            total: 125400,
            history: [12000, 15000, 11000, 18000, 14000, 16000],
            // Simulated historical data for interannual comparison
            interannualHistory: {
                currentYear: [85000, 92000, 78000, 110000, 95000, 120000],
                prevYear: [65000, 72000, 58000, 80000, 75000, 90000],
                year2024: [55000, 62000, 48000, 70000, 65000, 80000],
                year2023: [45000, 52000, 38000, 60000, 55000, 70000]
            },
            items: [
                { rftNumber: 'RFT - 94', created: '05-mar-26', subsidiary: '311 - Germany', amount: 32188.71, currency: 'EUR', age: 55, responsable: 'Andrea', link: 'https://meltwater.atlassian.net/browse/RFT-94', status: 'Pendiente' },
                { rftNumber: 'RFT - 110', created: '31-mar-26', subsidiary: '102 - US', amount: 19800.00, currency: 'USD', age: 29, responsable: 'Andrea', link: 'https://meltwater.atlassian.net/browse/RFT-110', status: 'Validando' },
                { rftNumber: 'RFT - 129', created: '28-abr-26', subsidiary: '421 - Finland', amount: 8332.80, currency: 'EUR', age: 1, responsable: 'Andrea', link: 'https://meltwater.atlassian.net/browse/RFT-129', status: 'Pendiente' },
                { rftNumber: 'RFT - 130', created: '28-abr-26', subsidiary: '241 - Singapore', amount: 548.10, currency: 'USD', age: 1, responsable: 'Andrea', link: 'https://meltwater.atlassian.net/browse/RFT-130', status: 'Pendiente' }
            ]
        },
        // --- COMPARATIVA ANUAL DE APLICACIÓN DE EFECTIVO ---
        appliedCashHistory: {
            currentYear: [1850000, 1920000, 1780000, 2100000, 1950000, 2200000], // Montos aplicados por mes (2026)
            prevYear: [1650000, 1720000, 1580000, 1800000, 1750000, 1900000],    // Montos aplicados por mes (2025)
            year2024: [1550000, 1620000, 1480000, 1700000, 1650000, 1800000],    // Montos aplicados por mes (2024)
            year2023: [1450000, 1520000, 1380000, 1600000, 1550000, 1700000]     // Montos aplicados por mes (2023)
        }
    }
};
