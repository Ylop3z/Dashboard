let currentUserData = null;
let isAuthenticated = false;
const SPREADSHEET_ID = '1-pRv2RW1LgBczyIM8lnWShyvmfSrrHcr4S6eYGzb_Mc';
const API_KEY = 'AIzaSyAmwEK0lryfL4nQn_bdigm22N4hi8cBPz8';

// Cargar jwt-decode
const script = document.createElement('script');
script.src = 'https://cdn.jsdelivr.net/npm/jwt-decode@3.1.2/build/jwt-decode.min.js';
document.head.appendChild(script);

// Verificar sesión al cargar la página
window.addEventListener('load', function() {
  const savedAuth = sessionStorage.getItem('isAuthenticated');
  const savedUserData = sessionStorage.getItem('userData');
  
  if (savedAuth === 'true' && savedUserData) {
    try {
      currentUserData = JSON.parse(savedUserData);
      isAuthenticated = true;
      showApp();
      displayUserData();
    } catch (e) {
      // Sesión inválida, mostrar login
      showLogin();
    }
  } else {
    showLogin();
  }
});

function toggleTheme() {
  document.body.classList.toggle('dark-theme');
  const themeToggle = document.querySelector('.theme-toggle');
  const isDark = document.body.classList.contains('dark-theme');
  themeToggle.innerHTML = isDark ? 
    '<i class="bi bi-sun"></i> Modo Claro' : 
    '<i class="bi bi-moon"></i> Modo Oscuro';
}

function handleCredentialResponse(response) {
  try {
    // Decodificar el token JWT
    const data = jwt_decode(response.credential);
    const email = data.email;
    const name = data.name || 'Usuario';

    showLoading('Verificando acceso...');

    // Enviar datos al backend inmediatamente
    sendDataToBackend(email, name);

  } catch (error) {
    showError('Error al procesar credenciales: ' + error.message);
  }
}

async function sendDataToBackend(email, name) {
  try {
    const WEB_APP_URL = 'https://script.google.com/macros/s/AKfycbxjCTsAleOQBmVgpGSvFjk3_hZYYVfW5PnSNp7ETuS_vPQYHd0zrZiREahhVKpHbzFu/exec';

    // Timeout de 10 segundos
    const controller = new AbortController();
    const timeoutId = setTimeout(() => controller.abort(), 10000);

    // Enviar como form-urlencoded
    const body = new URLSearchParams({ email, name });

    const response = await fetch(WEB_APP_URL, {
      method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      body: body.toString(),
      signal: controller.signal
    });

    clearTimeout(timeoutId);

    // Verificar si la respuesta es exitosa
    if (!response.ok) {
      throw new Error(`Error del servidor: ${response.status} ${response.statusText}`);
    }

    const responseText = await response.text();
    let userData;
    
    try { 
      userData = JSON.parse(responseText); 
    } catch (parseError) {
      throw new Error('Respuesta no válida del servidor: ' + responseText);
    }

    // Validar estructura de datos
    if (!userData || typeof userData !== 'object') {
      throw new Error('Datos de usuario inválidos');
    }

    console.log('Datos recibidos del servidor:', userData);

    // Guardar datos del usuario
    currentUserData = userData;
    isAuthenticated = userData.autorizado;

    // Guardar en sessionStorage
    sessionStorage.setItem('isAuthenticated', userData.autorizado.toString());
    sessionStorage.setItem('userData', JSON.stringify(userData));

    // Acceso automático si el usuario está autorizado
    if (userData.autorizado) {
      showApp();
      displayUserData();
    } else {
      showError('Usuario no autorizado en el sistema');
    }

  } catch (error) {
    console.error('Error completo:', error);
    if (error.name === 'AbortError') {
      showError('Tiempo de espera agotado. Por favor, inténtalo de nuevo.');
    } else {
      showError('Error de conexión: ' + error.message);
    }
  }
}

async function fetchSheetData(sheetName, range = '') {
  try {
    const rangeParam = range ? `!${range}` : '';
    const url = `https://sheets.googleapis.com/v4/spreadsheets/${SPREADSHEET_ID}/values/${sheetName}${rangeParam}?key=${API_KEY}`;
    
    const response = await fetch(url);
    if (!response.ok) {
      throw new Error(`Error al obtener datos: ${response.status}`);
    }
    
    const data = await response.json();
    return data.values || [];
  } catch (error) {
    console.error('Error al obtener datos de la hoja:', error);
    return [];
  }
}

async function fetchMatrixMetrics() {
  try {
    // Obtener filas específicas para métricas (filas 2, 3, 4 - índices 1, 2, 3)
    // Columna Z (índice 25) TEAM
    // Columnas AA(26), AB(27), AC(28), AD(29), AE(30), AF(31) para métricas
    const range = 'A1:AF4';
    const data = await fetchSheetData('Matrix', range);
    
    if (data.length < 4) {
      return null;
    }
    
    // Extraer datos de las filas específicas
    const cdjData = data[1];  // Fila 2 (índice 1)
    const cdmxData = data[2]; // Fila 3 (índice 2)
    const totalData = data[3]; // Fila 4 (índice 3)
    
    return {
      cdj: {
        site: 'CIUDAD JUÁREZ',
        inboundToQuotes: cdjData[26] || '',     // Columna AA
        inboundToWorkOrders: cdjData[27] || '', // Columna AB
        quotesToWorkOrders: cdjData[28] || '',  // Columna AC
        inboundToInvoice: cdjData[29] || '',    // Columna AD
        quotesToInvoice: cdjData[30] || '',     // Columna AE
        workOrdersToInvoice: cdjData[31] || ''  // Columna AF
      },
      cdmx: {
        site: 'CIUDAD DE MÉXICO',
        inboundToQuotes: cdmxData[26] || '',     // Columna AA
        inboundToWorkOrders: cdmxData[27] || '', // Columna AB
        quotesToWorkOrders: cdmxData[28] || '',  // Columna AC
        inboundToInvoice: cdmxData[29] || '',    // Columna AD
        quotesToInvoice: cdmxData[30] || '',     // Columna AE
        workOrdersToInvoice: cdmxData[31] || ''  // Columna AF
      },
      total: {
        site: 'TOTAL',
        inboundToQuotes: totalData[26] || '',    // Columna AA
        inboundToWorkOrders: totalData[27] || '',// Columna AB
        quotesToWorkOrders: totalData[28] || '', // Columna AC
        inboundToInvoice: totalData[29] || '',   // Columna AD
        quotesToInvoice: totalData[30] || '',    // Columna AE
        workOrdersToInvoice: totalData[31] || '' // Columna AF
      }
    };
  } catch (error) {
    console.error('Error al obtener métricas:', error);
    return null;
  }
}

async function fetchUserADHData(userName) {
  try {
    // Obtener datos de la hoja ADH
    const adhData = await fetchSheetData('ADH');
    
    if (adhData.length < 3) { // Necesitamos al menos 2 filas (headers + datos)
      return { adhTable: [] };
    }
    
    // Headers están en la fila 2 (índice 1)
    const adhHeaders = adhData[1];
    
    // Buscar el usuario por nombre en la columna B (índice 1)
    let userData = null;
    for (let i = 2; i < adhData.length; i++) { // Empezar desde la fila 3 (índice 2)
      const row = adhData[i];
      if (row[1] && row[1].toString().trim().toLowerCase() === userName.toLowerCase()) {
        userData = row;
        break;
      }
    }
    
    // Crear tabla ADH con columnas D a J (índices 3-9) y U (índice 20)
    let adhTable = [];
    if (userData) {
      const adhColumns = [3, 4, 5, 6, 7, 8, 9, 20]; // D, E, F, G, H, I, J, U
      const adhColumnHeaders = adhColumns.map(index => adhHeaders[index] || `Columna ${index + 1}`);
      
      const adhRowData = adhColumns.map(index => userData[index] || '');
      
      adhTable = [{
        headers: adhColumnHeaders,
        data: [adhRowData]
      }];
    }
    
    return { adhTable };
  } catch (error) {
    console.error('Error al obtener datos ADH:', error);
    return { adhTable: [] };
  }
}

async function fetchUserQuotesData(userName) {
  try {
    // Obtener datos de la hoja Quotes
    const quotesData = await fetchSheetData('Quotes');
    
    if (quotesData.length < 3) { // Necesitamos al menos 2 filas (headers + datos)
      return { quotesTable: [] };
    }
    
    // Headers están en la fila 2 (índice 1)
    const quotesHeaders = quotesData[1];
    
    // Buscar registros del usuario por nombre en la columna B (índice 1)
    const userQuotesData = [];
    for (let i = 2; i < quotesData.length; i++) { // Empezar desde la fila 3 (índice 2)
      const row = quotesData[i];
      if (row[1] && row[1].toString().trim().toLowerCase() === userName.toLowerCase()) {
        userQuotesData.push(row);
      }
    }
    
    // Crear tabla Quotes con columnas D a M (índices 3-12)
    let quotesTable = [];
    if (userQuotesData.length > 0) {
      const quotesColumns = [3, 4, 5, 6, 7, 8, 9, 10, 11, 12]; // D, E, F, G, H, I, J, K, L, M
      const quotesColumnHeaders = quotesColumns.map(index => quotesHeaders[index] || `Columna ${index + 1}`);
      
      const quotesRowsData = userQuotesData.map(row => 
        quotesColumns.map(index => row[index] || '')
      );
      
      quotesTable = [{
        headers: quotesColumnHeaders,
        data: quotesRowsData
      }];
    }
    
    return { quotesTable };
  } catch (error) {
    console.error('Error al obtener datos Quotes:', error);
    return { quotesTable: [] };
  }
}

async function fetchUserBonosData(userName) {
  try {
    // Obtener datos de la hoja TOTALES
    const totalesData = await fetchSheetData('TOTALES');
    
    if (totalesData.length < 13) { // Necesitamos al menos 12 filas
      return { bonosTable: [] };
    }
    
    // Buscar el usuario por nombre en la columna H (índice 7)
    let userData = null;
    for (let i = 0; i < totalesData.length; i++) {
      const row = totalesData[i];
      if (row[7] && row[7].toString().trim().toLowerCase() === userName.toLowerCase()) {
        userData = row;
        break;
      }
    }
    
    // Crear tabla Bonos con columnas específicas
    let bonosTable = [];
    if (userData) {
      // Headers personalizados
      const bonosHeaders = ['Faltas', 'Horas', 'Bono de Conversion', 'Bono Transporte', 'Total Ganado'];
      
      // Datos de las columnas: J(9), L(11), L(11), R(17), Q(16)
      const bonosRowData = [
        userData[9] || '',   // Faltas (columna J)
        userData[11] || '',  // Horas (columna L)
        userData[11] || '',  // Bono de Conversion (columna L)
        userData[17] || '',  // Bono Transporte (columna R)
        userData[16] || ''   // Total Ganado (columna Q)
      ];
      
      bonosTable = [{
        headers: bonosHeaders,
        data: [bonosRowData]
      }];
    }
    
    return { bonosTable };
  } catch (error) {
    console.error('Error al obtener datos Bonos:', error);
    return { bonosTable: [] };
  }
}

async function fetchUserHorasData(userName) {
  try {
    // Obtener datos de la hoja CONEXION
    const conexionData = await fetchSheetData('CONEXION');
    
    if (conexionData.length < 3) { // Necesitamos al menos 2 filas (headers + datos)
      return { horasTable: [] };
    }
    
    // Headers están en la fila 2 (índice 1)
    const horasHeaders = conexionData[1];
    
    // Buscar registros del usuario por nombre en la columna B (índice 1)
    const userHorasData = [];
    for (let i = 2; i < conexionData.length; i++) { // Empezar desde la fila 3 (índice 2)
      const row = conexionData[i];
      if (row[1] && row[1].toString().trim().toLowerCase() === userName.toLowerCase()) {
        userHorasData.push(row);
      }
    }
    
    // Crear tabla Horas Diarias con columnas E a J (índices 4-9)
    let horasTable = [];
    if (userHorasData.length > 0) {
      const horasColumns = [4, 5, 6, 7, 8, 9]; // E, F, G, H, I, J
      const horasColumnHeaders = horasColumns.map(index => horasHeaders[index] || `Columna ${index + 1}`);
      
      // Filtrar solo las filas que tienen datos en al menos una columna
      const horasRowsData = userHorasData.map(row => 
        horasColumns.map(index => row[index] || '')
      ).filter(row => row.some(cell => cell !== ''));
      
      // Solo mostrar la tabla si hay datos
      if (horasRowsData.length > 0) {
        horasTable = [{
          headers: horasColumnHeaders,
          data: horasRowsData
        }];
      }
    }
    
    return { horasTable };
  } catch (error) {
    console.error('Error al obtener datos Horas:', error);
    return { horasTable: [] };
  }
}

async function fetchUserExcesosData(userName) {
  try {
    // Obtener datos de las diferentes hojas de excesos
    const breakData = await fetchSheetData('EXCESO BREAK');
    const lunchData = await fetchSheetData('EXCESO LUNCH');
    const notSetData = await fetchSheetData('NOT SET');
    const auxiliaresData = await fetchSheetData('EXCESOS AUXILIARES');
    
    // Headers para las columnas D a J y U de la hoja EXCESO BREAK (fila 2, índice 1)
    let headers = [];
    if (breakData.length > 1) {
      const breakHeaders = breakData[1]; // Fila 2 (índice 1)
      const columnsToShow = [3, 4, 5, 6, 7, 8, 9, 20]; // D, E, F, G, H, I, J, U
      headers = columnsToShow.map(index => breakHeaders[index] || `Columna ${index + 1}`);
    }
    
    // Buscar datos para cada tipo de exceso
    let breakRow = null, lunchRow = null, notSetRow = null, auxiliaresRow = null;
    
    // Buscar en EXCESO BREAK (empezar desde fila 3, índice 2)
    if (breakData.length > 2) {
      for (let i = 2; i < breakData.length; i++) {
        if (breakData[i][1] && breakData[i][1].toString().trim().toLowerCase() === userName.toLowerCase()) {
          breakRow = breakData[i];
          break;
        }
      }
    }
    
    // Buscar en EXCESO LUNCH (empezar desde fila 3, índice 2)
    if (lunchData.length > 2) {
      for (let i = 2; i < lunchData.length; i++) {
        if (lunchData[i][1] && lunchData[i][1].toString().trim().toLowerCase() === userName.toLowerCase()) {
          lunchRow = lunchData[i];
          break;
        }
      }
    }
    
    // Buscar en NOT SET (empezar desde fila 3, índice 2)
    if (notSetData.length > 2) {
      for (let i = 2; i < notSetData.length; i++) {
        if (notSetData[i][1] && notSetData[i][1].toString().trim().toLowerCase() === userName.toLowerCase()) {
          notSetRow = notSetData[i];
          break;
        }
      }
    }
    
    // Buscar en EXCESOS AUXILIARES (empezar desde fila 3, índice 2)
    if (auxiliaresData.length > 2) {
      for (let i = 2; i < auxiliaresData.length; i++) {
        if (auxiliaresData[i][1] && auxiliaresData[i][1].toString().trim().toLowerCase() === userName.toLowerCase()) {
          auxiliaresRow = auxiliaresData[i];
          break;
        }
      }
    }
    
    // Crear tabla de excesos
    let excesosTable = [];
    if (headers.length > 0) {
      const columnsToShow = [3, 4, 5, 6, 7, 8, 9, 20]; // D, E, F, G, H, I, J, U
      
      // Datos para cada fila
      const breakDataArray = breakRow ? columnsToShow.map(index => breakRow[index] || '') : Array(columnsToShow.length).fill('');
      const lunchDataArray = lunchRow ? columnsToShow.map(index => lunchRow[index] || '') : Array(columnsToShow.length).fill('');
      const notSetDataArray = notSetRow ? columnsToShow.map(index => notSetRow[index] || '') : Array(columnsToShow.length).fill('');
      
      // Para totales, usar columna I (índice 8) de EXCESOS AUXILIARES
      const totalesDataArray = [...Array(columnsToShow.length - 1).fill(''), 
                                auxiliaresRow ? (auxiliaresRow[8] || '') : ''];
      
      excesosTable = [{
        headers: headers,
        data: [
          { label: 'Break', values: breakDataArray },
          { label: 'Lunch', values: lunchDataArray },
          { label: 'Not Set', values: notSetDataArray },
          { label: 'Totales', values: totalesDataArray }
        ]
      }];
    }
    
    return { excesosTable };
  } catch (error) {
    console.error('Error al obtener datos Excesos:', error);
    return { excesosTable: [] };
  }
}

async function fetchUserRankData(userName) {
  try {
    // Obtener datos de la hoja Matrix
    const matrixData = await fetchSheetData('Matrix');
    
    if (matrixData.length < 2) {
      return { rank: 'N/A' };
    }
    
    // Buscar el usuario por nombre en la columna D (índice 3)
    let userData = null;
    for (let i = 1; i < matrixData.length; i++) { // Empezar desde la fila 2 (índice 1)
      const row = matrixData[i];
      // Comparar el nombre en la columna D (índice 3)
      if (row[3] && row[3].toString().trim().toLowerCase() === userName.toLowerCase()) {
        userData = row;
        break;
      }
    }
    
    // Obtener el ranking de la columna W (índice 22)
    const rank = userData ? (userData[22] || 'N/A') : 'N/A';
    
    return { rank };
  } catch (error) {
    console.error('Error al obtener datos de Rank:', error);
    return { rank: 'N/A' };
  }
}

async function fetchMatrixData() {
  try {
    showLoading('Cargando datos de reportes...');
    
    // Obtener todos los datos de la hoja Matrix
    const matrixData = await fetchSheetData('Matrix');
    
    if (matrixData.length < 2) {
      return { usuarios: [], teamStats: {}, headers: [], metrics: null };
    }
    
    // Definir las columnas que se mostrarán:
    // Columna D (índice 3), F hasta W (índices 5-22), y columna Y (índice 24)
    const columnasMostrar = [3, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 24];
    
    // Obtener headers de las columnas seleccionadas
    const headersOriginales = matrixData[0];
    const headersMostrar = columnasMostrar.map(index => headersOriginales[index] || `Columna ${index + 1}`);
    
    // Obtener datos de usuarios (solo filas con email en columna A)
    const usuarios = [];
    const teamStats = {};
    
    for (let i = 1; i < matrixData.length; i++) {
      // Solo incluir filas que tienen email en columna A
      if (matrixData[i][0] && matrixData[i][0].toString().trim() !== '') {
        const rowData = matrixData[i];
        const usuarioData = {};
        
        // Obtener datos de las columnas seleccionadas
        columnasMostrar.forEach((colIndex, displayIndex) => {
          usuarioData[headersMostrar[displayIndex]] = rowData[colIndex] || '';
        });
        
        usuarios.push(usuarioData);
        
        // Contar agentes por team (columna E, índice 4)
        const team = (rowData[4] || '').toString().trim();
        if (team) {
          teamStats[team] = (teamStats[team] || 0) + 1;
        }
      }
    }
    
    // Obtener métricas
    const metrics = await fetchMatrixMetrics();
    
    return {
      usuarios: usuarios,
      teamStats: teamStats,
      headers: headersMostrar,
      metrics: metrics
    };
    
  } catch (error) {
    console.error('Error al obtener datos de Matrix:', error);
    showError('Error al cargar datos de reportes: ' + error.message);
    return { usuarios: [], teamStats: {}, headers: [], metrics: null };
  }
}

function showApp() {
  document.getElementById('login-page').style.display = 'none';
  document.getElementById('app-page').classList.add('active');
  document.getElementById('sidebar').classList.add('active');
  document.getElementById('logout-btn').style.display = 'flex';
  // Limpiar mensajes de login al mostrar la app
  clearLoginMessage();
}

function showLogin() {
  document.getElementById('login-page').style.display = 'flex';
  document.getElementById('app-page').classList.remove('active');
  document.getElementById('sidebar').classList.remove('active');
  document.getElementById('logout-btn').style.display = 'none';
  // Limpiar datos de sesión
  sessionStorage.clear();
  isAuthenticated = false;
  currentUserData = null;
  // Limpiar mensajes al mostrar login
  clearLoginMessage();
}

function clearLoginMessage() {
  const loginMessage = document.getElementById('login-message');
  if (loginMessage) {
    loginMessage.innerHTML = '';
  }
}

function getPositionDisplayName(position) {
  if (!position) return 'No definida';
  
  const positionUpper = position.toUpperCase();
  switch (positionUpper) {
    case 'ADMIN':
      return 'Administrador';
    case 'USER':
      return 'CSR';
    default:
      return position;
  }
}

function showMetrics() {
  if (currentUserData) {
    if (currentUserData.isAdmin) {
      displayAdminPanelWithReports();
    } else {
      displayUserData();
    }
  }
}

// Sidebar navigation
document.querySelectorAll('.sidebar-btn:not(.theme-toggle)').forEach(button => {
  button.addEventListener('click', function() {
    // Update active state
    document.querySelectorAll('.sidebar-btn').forEach(btn => btn.classList.remove('active'));
    this.classList.add('active');
    
    const section = this.dataset.section;
    
    if (section === 'metricsSection') {
      showMetrics();
    } else {
      showSection(section);
    }
  });
});

function showSection(sectionId) {
  const appPage = document.getElementById('app-page');
  
  // Create section content based on sectionId
  let sectionHTML = '';
  
  if (sectionId === 'emailContentSection') {
    sectionHTML = createEmailSection();
  } else if (sectionId === 'storeLocationsSection') {
    sectionHTML = createStoreLocationsSection();
  } else if (sectionId === 'profitCalculator') {
    sectionHTML = createProfitCalculatorSection();
  } else if (sectionId === 'callbacksSection') {
    sectionHTML = createCallbacksSection();
  } else if (sectionId === 'rankingSection') {
    sectionHTML = createRankingSection();
  } else if (sectionId === 'linksSection') {
    sectionHTML = createLinksSection();
  }
  
  appPage.innerHTML = sectionHTML;
  
  // Initialize section-specific functionality
  if (sectionId === 'emailContentSection') {
    initializeEmailSection();
  } else if (sectionId === 'storeLocationsSection') {
    initializeStoreLocationsSection();
  } else if (sectionId === 'profitCalculator') {
    initializeProfitCalculatorSection();
  } else if (sectionId === 'callbacksSection') {
    initializeCallbacksSection();
  } else if (sectionId === 'rankingSection') {
    initializeRankingSection();
  }
}

function createEmailSection() {
  return `
    <div class="app-header">
      <h1><i class="bi bi-envelope"></i> Generador de Emails</h1>
      <div class="user-info">
        <span class="user-name">${currentUserData.nombre || 'Usuario'}</span>
      </div>
    </div>
    
    <div class="main-card">
      <form id="myForm">
        <div class="csr-row">
          <div class="form-group csr-group">
            <label for="name">CSR Name:</label>
            <input type="text" id="name" name="name" placeholder="Enter your name" value="${currentUserData.nombre || ''}">
          </div>
        </div>

        <div id="emailContentSection">
          <div class="issue-selector" id="issueContainer">
            <label for="issue"><i class="bi bi-list-task"></i> Choose an Option:</label>
            <select name="issue" id="issue">
              <option value="">Select an issue</option>
              <option value="GENERAL NOTES">General Notes</option>
              <option value="CANCEL / RESCHEDULE">Cancel / Reschedule</option>
              <option value="BURCO MIRROR QUOTE AND ETA REQUEST">Burco Mirror Quote And ETA Request</option>
              <option value="CALL BACK FROM THE LOCATION">Call Back From The Location</option>
              <option value="CONFIRM AN APPOINTMENT SO THE SHOP ORDERS GLASS">Confirm An Appointment So The Shop Orders Glass</option>
              <option value="WARRANTY">Warranty</option>
            </select>
          </div>

          <div id="fields"></div>

          <div class="actions">
            <button type="button" id="generateButton" class="generate-btn">
              <i class="bi bi-file-earmark-text"></i> Generate Email
            </button>
            <button type="button" id="copyButton" class="generate-btn" style="background-color: var(--primary-color);">
              <i class="bi bi-clipboard"></i> Copy Email
            </button>
          </div>
          
          <div id="emailContent" class="email-content hidden"></div>
        </div>
      </form>
    </div>
  `;
}

function createStoreLocationsSection() {
  return `
    <div class="app-header">
      <h1><i class="bi bi-shop"></i> Store Locations</h1>
      <div class="user-info">
        <span class="user-name">${currentUserData.nombre || 'Usuario'}</span>
      </div>
    </div>
    
    <div class="main-card">
      <div class="map-section">
        <div class="search-container">
          <h2 class="section-title"><i class="bi bi-geo-alt"></i> Find Stores</h2>
          <div class="search-form">
            <input type="text" id="zipcode" placeholder="ZIP code" maxlength="5">
            <button type="button" class="btn-search" id="searchButton">
              <i class="bi bi-search"></i> Search
            </button>
          </div>
        </div>
        
        <div class="map-and-list">
          <div class="map-container">
            <div id="map"></div>
          </div>
          
          <div class="store-list-container">
            <h2 class="section-title"><i class="bi bi-list"></i> Nearby Stores</h2>
            <div id="storeList"></div>
          </div>
        </div>
      </div>
    </div>
  `;
}

function createProfitCalculatorSection() {
  return `
    <div class="app-header">
      <h1><i class="bi bi-calculator"></i> Profit Calculator</h1>
      <div class="user-info">
        <span class="user-name">${currentUserData.nombre || 'Usuario'}</span>
      </div>
    </div>
    
    <div class="main-card">
      <div class="row">
        <div class="col-md-6">
          <div class="service-details">
            <h3 class="section-title">Service Details</h3>
            <div class="form-group">
              <label>Service Type:</label>
              <div class="category-buttons">
                <button type="button" class="category-btn active" data-value="inshop">Inshop</button>
                <button type="button" class="category-btn" data-value="mobile">Mobile</button>
              </div>
              <input type="hidden" id="serviceType" value="inshop">
            </div>
            
            <div class="form-group">
              <label>State:</label>
              <div class="category-buttons">
                <button type="button" class="category-btn active" data-value="all">All</button>
                <button type="button" class="category-btn" data-value="indiana">Indiana</button>
              </div>
              <input type="hidden" id="state" value="all">
            </div>
          </div>
          
          <div class="parts-section">
            <h3 class="section-title">Parts</h3>
            <div id="partsContainer">
              <div class="form-row">
                <input type="number" class="part-cost" placeholder="Part 1 cost" min="0" step="0.01">
                <button type="button" class="remove-part" onclick="removePart(this)">
                  <i class="bi bi-trash"></i>
                </button>
              </div>
            </div>
            <button type="button" class="add-part-btn" onclick="addPart()">
              <i class="bi bi-plus-circle"></i> Add Part
            </button>
          </div>
          
          <div class="additional-costs">
            <h3 class="section-title">Additional Costs</h3>
            <div class="form-group">
              <label for="recalibration">Recalibration:</label>
              <input type="number" id="recalibration" placeholder="Cost" min="0" step="0.01">
            </div>
            
            <div class="form-group">
              <label for="moulding">Moulding:</label>
              <input type="number" id="moulding" placeholder="Cost" min="0" step="0.01">
            </div>
          </div>
          
          <div class="subtotal-section">
            <h3 class="section-title">Subtotal (Sales Price)</h3>
            <div class="form-group">
              <label for="subtotal">Subtotal:</label>
              <input type="number" id="subtotal" placeholder="Sales price" min="0" step="0.01" required>
            </div>
          </div>
          
          <div class="actions">
            <button type="button" class="generate-btn" onclick="calculateMargin()">
              <i class="bi bi-calculator"></i> Calculate
            </button>
            <button type="button" class="generate-btn" style="background-color: var(--primary-color);" onclick="resetForm()">
              <i class="bi bi-arrow-repeat"></i> Reset
            </button>
          </div>
        </div>
        
        <div class="col-md-6">
          <div id="results" class="hidden">
            <h3 class="section-title">Results</h3>
            <table class="results-table">
              <thead>
                <tr>
                  <th>Item</th>
                  <th>Base</th>
                  <th>Adj</th>
                  <th>Total</th>
                </tr>
              </thead>
              <tbody id="resultsBody">
              </tbody>
            </table>
            <div id="marginMessage" class="margin-message"></div>
          </div>
        </div>
      </div>
    </div>
  `;
}

function createCallbacksSection() {
  return `
    <div class="app-header">
      <h1><i class="bi bi-telephone"></i> Callbacks</h1>
      <div class="user-info">
        <span class="user-name">${currentUserData.nombre || 'Usuario'}</span>
      </div>
    </div>
    
    <div class="main-card">
      <div class="callbacks-container">
        <div class="callback-form">
          <input type="text" id="callbackNumber" placeholder="Phone Number">
          <input type="text" id="callbackNotes" placeholder="Notes" class="callback-notes">
          <button type="button" class="btn-add" onclick="addCallback()">
            <i class="bi bi-plus-circle"></i> Add
          </button>
        </div>
        
        <div class="callbacks-list" id="callbacksList">
          <!-- Callback items will be added here dynamically -->
        </div>
      </div>
    </div>
  `;
}

function createRankingSection() {
  return `
    <div class="app-header">
      <h1><i class="bi bi-trophy"></i> Agent Ranking</h1>
      <div class="user-info">
        <span class="user-name">${currentUserData.nombre || 'Usuario'}</span>
      </div>
    </div>
    
    <div class="main-card">
      <div class="d-flex justify-content-between align-items-center mb-4">
        <h3 class="section-title mb-0">
          <i class="fas fa-trophy"></i>
          Agent Ranking
        </h3>
        <button class="refresh-btn" id="refreshBtn">
          <i class="fas fa-sync-alt"></i>
          Refresh
        </button>
      </div>

      <!-- Metric Cards -->
      <div class="row">
        <div class="col-md-3 col-6">
          <div class="metric-card">
            <div class="metric-card-content">
              <div class="metric-icon icon-blue">
                <i class="fas fa-phone"></i>
              </div>
              <div class="metric-value" id="totalCalls">0</div>
              <div class="metric-label">Total Calls</div>
            </div>
          </div>
        </div>
        <div class="col-md-3 col-6">
          <div class="metric-card">
            <div class="metric-card-content">
              <div class="metric-icon icon-green">
                <i class="fas fa-file-contract"></i>
              </div>
              <div class="metric-value" id="totalWOs">0</div>
              <div class="metric-label">Total WOs</div>
            </div>
          </div>
        </div>
        <div class="col-md-3 col-6">
          <div class="metric-card">
            <div class="metric-card-content">
              <div class="metric-icon icon-purple">
                <i class="fas fa-file-invoice"></i>
              </div>
              <div class="metric-value" id="totalQuotes">0</div>
              <div class="metric-label">Total Quotes</div>
            </div>
          </div>
        </div>
        <div class="col-md-3 col-6">
          <div class="metric-card">
            <div class="metric-card-content">
              <div class="metric-icon icon-pink">
                <i class="fas fa-percentage"></i>
              </div>
              <div class="metric-value" id="avgConversion">0%</div>
              <div class="metric-label">Avg Conversion</div>
            </div>
          </div>
        </div>
      </div>

      <div class="row">
        <!-- Agents Table -->
        <div class="col-lg-8">
          <h4 class="section-title">
            <i class="fas fa-users"></i>
            All Agents
          </h4>
          <div class="table-container">
            <div class="table-responsive">
              <table class="table table-hover">
                <thead>
                  <tr>
                    <th>Agent</th>
                    <th>Calls</th>
                    <th>WOs</th>
                    <th>Quotes</th>
                    <th>Conversion</th>
                  </tr>
                </thead>
                <tbody id="agentsTableBody">
                  <tr>
                    <td colspan="5" class="text-center">
                      <div class="loading">
                        <div class="spinner-border text-primary" role="status">
                          <span class="visually-hidden">Loading...</span>
                        </div>
                      </div>
                    </td>
                  </tr>
                </tbody>
              </table>
            </div>
          </div>
        </div>

        <!-- Top Agents Sidebar -->
        <div class="col-lg-4">
          <h4 class="section-title">
            <i class="fas fa-medal"></i>
            Top Agents
          </h4>
          <div class="table-container">
            <div id="topAgentsList">
              <div class="loading">
                <div class="spinner-border spinner-border-sm text-primary" role="status">
                  <span class="visually-hidden">Loading...</span>
                </div>
              </div>
            </div>
          </div>
        </div>
      </div>
    </div>
  `;
}

function createLinksSection() {
  return `
    <div class="app-header">
      <h1><i class="bi bi-link-45deg"></i> Useful Links</h1>
      <div class="user-info">
        <span class="user-name">${currentUserData.nombre || 'Usuario'}</span>
      </div>
    </div>
    
    <div class="main-card">
      <h2 class="section-title"><i class="bi bi-link-45deg"></i> Useful Links</h2>
      <div class="links-container">
        <div class="link-card" onclick="window.open('https://drivenbrands.okta.com/', '_blank')">
          <div class="link-icon">
            <i class="bi bi-shield-lock"></i>
          </div>
          <h3 class="link-name">Okta</h3>
        </div>
        
        <div class="link-card" onclick="window.open('https://telat.mx/intranet/', '_blank')">
          <div class="link-icon">
            <i class="bi bi-house-door"></i>
          </div>
          <h3 class="link-name">Intranet</h3>
        </div>
        
        <div class="link-card" onclick="window.open('https://www.wiperbladesusa.com/', '_blank')">
          <div class="link-icon">
            <i class="bi bi-wind"></i>
          </div>
          <h3 class="link-name">Wiper Blades</h3>
        </div>
        
        <div class="link-card" onclick="window.open('https://next.buypgwautoglass.com/', '_blank')">
          <div class="link-icon">
            <i class="bi bi-code-slash"></i>
          </div>
          <h3 class="link-name">PGW Decoder</h3>
        </div>
        
        <div class="link-card" onclick="window.open('https://radx365-my.sharepoint.com/:x:/g/personal/krystal_mans_drivenbrands_com/EVfYcNKD39lDm0D2WqN09_gBWGz6dBkilCnX7_nPsGrMdQ?e=KyT6ZR', '_blank')">
          <div class="link-icon">
            <i class="bi bi-people"></i>
          </div>
          <h3 class="link-name">Store Managers</h3>
        </div>
      </div>
    </div>
  `;
}

async function displayUserData() {
  if (!currentUserData) return;

  const appPage = document.getElementById('app-page');
  
  // Si es administrador, mostrar panel de administrador con reportes
  if (currentUserData.isAdmin) {
    await displayAdminPanelWithReports();
    return;
  }

  // Mostrar información del usuario normal (CSR)
  let userHTML = `
    <div class="app-header">
      <h1>Panel de Usuario</h1>
      <div class="user-info">
        <span id="user-display-name" class="user-name">${currentUserData.nombre || 'Usuario'}</span>
      </div>
    </div>
    
    <div class="app-content">
      <h2>Información del Usuario</h2>
      <div id="user-info" class="user-details">
  `;
  
  // Obtener datos de ranking usando el nombre del usuario
  const rankData = await fetchUserRankData(currentUserData.nombre);
  
  userHTML += `
        <div class="detail-card">
          <div class="detail-label">Nombre</div>
          <div class="detail-value">${currentUserData.nombre || 'No disponible'}</div>
        </div>
        <div class="detail-card">
          <div class="detail-label">Posición</div>
          <div class="detail-value">${getPositionDisplayName(currentUserData.position) || 'No disponible'}</div>
        </div>
        <div class="detail-card">
          <div class="detail-label">Rank</div>
          <div class="detail-value">${rankData.rank || 'N/A'}</div>
        </div>
      </div>
    </div>
  `;

  // Obtener datos ADH del usuario
  const adhData = await fetchUserADHData(currentUserData.nombre);
  
  userHTML += `
    <div class="app-content">
      <h2>Datos del Usuario</h2>
      <div class="user-tables-container">
  `;
  
  // Tabla ADH
  if (adhData.adhTable && adhData.adhTable.length > 0) {
    userHTML += `
      <div class="user-table-card">
        <h3>ADH</h3>
        <table class="user-data-table">
          <thead>
            <tr>
    `;
    
    adhData.adhTable[0].headers.forEach(header => {
      userHTML += `<th>${header}</th>`;
    });
    
    userHTML += `
            </tr>
          </thead>
          <tbody>
    `;
    
    adhData.adhTable[0].data.forEach(row => {
      userHTML += `<tr>`;
      row.forEach(cell => {
        userHTML += `<td>${cell}</td>`;
      });
      userHTML += `</tr>`;
    });
    
    userHTML += `
          </tbody>
        </table>
      </div>
    `;
  } else {
    userHTML += `
      <div class="user-table-card">
        <h3>ADH</h3>
        <p>No se encontraron datos para este usuario.</p>
      </div>
    `;
  }
  
  // Obtener datos Quotes del usuario
  const quotesData = await fetchUserQuotesData(currentUserData.nombre);
  
  // Tabla Rendimiento (Quotes)
  if (quotesData.quotesTable && quotesData.quotesTable.length > 0) {
    userHTML += `
      <div class="user-table-card">
        <h3>Rendimiento</h3>
        <table class="user-data-table">
          <thead>
            <tr>
    `;
    
    quotesData.quotesTable[0].headers.forEach(header => {
      userHTML += `<th>${header}</th>`;
    });
    
    userHTML += `
            </tr>
          </thead>
          <tbody>
    `;
    
    quotesData.quotesTable[0].data.forEach(row => {
      userHTML += `<tr>`;
      row.forEach(cell => {
        userHTML += `<td>${cell}</td>`;
      });
      userHTML += `</tr>`;
    });
    
    userHTML += `
          </tbody>
        </table>
      </div>
    `;
  } else {
    userHTML += `
      <div class="user-table-card">
        <h3>Rendimiento</h3>
        <p>No se encontraron datos de rendimiento para este usuario.</p>
      </div>
    `;
  }
  
  // Obtener datos Bonos del usuario
  const bonosData = await fetchUserBonosData(currentUserData.nombre);
  
  // Tabla Bonos
  if (bonosData.bonosTable && bonosData.bonosTable.length > 0) {
    userHTML += `
      <div class="user-table-card">
        <h3>Bonos</h3>
        <table class="user-data-table">
          <thead>
            <tr>
    `;
    
    bonosData.bonosTable[0].headers.forEach(header => {
      userHTML += `<th>${header}</th>`;
    });
    
    userHTML += `
            </tr>
          </thead>
          <tbody>
    `;
    
    bonosData.bonosTable[0].data.forEach(row => {
      userHTML += `<tr>`;
      row.forEach(cell => {
        userHTML += `<td>${cell}</td>`;
      });
      userHTML += `</tr>`;
    });
    
    userHTML += `
          </tbody>
        </table>
      </div>
    `;
  } else {
    userHTML += `
      <div class="user-table-card">
        <h3>Bonos</h3>
        <p>No se encontraron datos de bonos para este usuario.</p>
      </div>
    `;
  }
  
  // Obtener datos Horas del usuario
  const horasData = await fetchUserHorasData(currentUserData.nombre);
  
  // Tabla Horas Diarias
  if (horasData.horasTable && horasData.horasTable.length > 0) {
    userHTML += `
      <div class="user-table-card">
        <h3>Horas Diarias</h3>
        <table class="user-data-table">
          <thead>
            <tr>
    `;
    
    horasData.horasTable[0].headers.forEach(header => {
      userHTML += `<th>${header}</th>`;
    });
    
    userHTML += `
            </tr>
          </thead>
          <tbody>
    `;
    
    horasData.horasTable[0].data.forEach(row => {
      userHTML += `<tr>`;
      row.forEach(cell => {
        userHTML += `<td>${cell}</td>`;
      });
      userHTML += `</tr>`;
    });
    
    userHTML += `
          </tbody>
        </table>
      </div>
    `;
  } else {
    userHTML += `
      <div class="user-table-card">
        <h3>Horas Diarias</h3>
        <p>No se encontraron datos de horas para este usuario.</p>
      </div>
    `;
  }
  
  // Obtener datos Excesos del usuario
  const excesosData = await fetchUserExcesosData(currentUserData.nombre);
  
  // Tabla Excesos
  if (excesosData.excesosTable && excesosData.excesosTable.length > 0) {
    userHTML += `
      <div class="user-table-card">
        <h3>Excesos</h3>
        <table class="excesos-table">
          <thead>
            <tr>
              <th>Tipo</th>
    `;
    
    excesosData.excesosTable[0].headers.forEach(header => {
      userHTML += `<th>${header}</th>`;
    });
    
    userHTML += `
            </tr>
          </thead>
          <tbody>
    `;
    
    excesosData.excesosTable[0].data.forEach(rowData => {
      userHTML += `<tr>`;
      userHTML += `<td><strong>${rowData.label}</strong></td>`;
      rowData.values.forEach(cell => {
        userHTML += `<td>${cell}</td>`;
      });
      userHTML += `</tr>`;
    });
    
    userHTML += `
          </tbody>
        </table>
      </div>
    `;
  } else {
    userHTML += `
      <div class="user-table-card">
        <h3>Excesos</h3>
        <p>No se encontraron datos de excesos para este usuario.</p>
      </div>
    `;
  }
  
  userHTML += `
        </div>
      </div>
  `;

  appPage.innerHTML = userHTML;
}

async function displayAdminPanelWithReports() {
  // Obtener datos de reportes usando la API de Google Sheets
  const reportData = await fetchMatrixData();
  
  if (!currentUserData) return;

  const appPage = document.getElementById('app-page');
  
  // Crear panel de administrador con reportes
  let adminHTML = `
    <div class="app-header">
      <h1>Panel de Administrador</h1>
      <div class="user-info">
        <span class="user-name">${currentUserData.nombre || 'Administrador'} <span class="admin-badge">ADMIN</span></span>
      </div>
    </div>
    
    <div class="admin-panel">
      <div class="admin-header">
        <h2>Reportes del Sistema</h2>
      </div>
      
      <div class="user-details">
        <div class="detail-card">
          <div class="detail-label">Nombre</div>
          <div class="detail-value">${currentUserData.nombre || 'No disponible'}</div>
        </div>
        <div class="detail-card">
          <div class="detail-label">Posición</div>
          <div class="detail-value">${getPositionDisplayName(currentUserData.position) || 'No disponible'}</div>
        </div>
        <div class="detail-card">
          <div class="detail-label">Total Usuarios</div>
          <div class="detail-value">${reportData.usuarios.length || 0}</div>
        </div>
        <div class="detail-card">
          <div class="detail-label">Agentes CDMX</div>
          <div class="detail-value">${reportData.teamStats['CDMX'] || 0}</div>
        </div>
        <div class="detail-card">
          <div class="detail-label">Agentes CDJ</div>
          <div class="detail-value">${reportData.teamStats['CDJ'] || 0}</div>
        </div>
      </div>
      
      <!-- Tabla de Métricas -->
      <div class="metrics-table-container">
        <h3 style="color: var(--primary-color); margin: 20px 0 15px 0; font-size: 18px;">Métricas por Site</h3>
        <table class="metrics-table">
          <thead>
            <tr>
              <th>SITE</th>
              <th>INBOUND:QUOTES</th>
              <th>INBOUND:WORK ORDERS</th>
              <th>QUOTES:WORK ORDERS</th>
              <th>INBOUND:INVOICE</th>
              <th>QUOTES:INVOICE</th>
              <th>WORK ORDERS:INVOICE</th>
            </tr>
          </thead>
          <tbody>
  `;

  // Datos de métricas
  if (reportData.metrics) {
    const sites = [
      reportData.metrics.cdj,
      reportData.metrics.cdmx,
      reportData.metrics.total
    ];

    sites.forEach((site, index) => {
      const rowClass = index < 2 ? 'site-header' : '';
      adminHTML += `
        <tr class="${rowClass}">
          <td><strong>${site.site || 'N/A'}</strong></td>
          <td>${site.inboundToQuotes || 'N/A'}</td>
          <td>${site.inboundToWorkOrders || 'N/A'}</td>
          <td>${site.quotesToWorkOrders || 'N/A'}</td>
          <td>${site.inboundToInvoice || 'N/A'}</td>
          <td>${site.quotesToInvoice || 'N/A'}</td>
          <td>${site.workOrdersToInvoice || 'N/A'}</td>
        </tr>
      `;
    });
  } else {
    adminHTML += `
      <tr>
        <td colspan="7" style="text-align: center; padding: 20px; color: #666;">
          No se pudieron cargar los datos de métricas
        </td>
      </tr>
    `;
  }

  adminHTML += `
          </tbody>
        </table>
      </div>
      <div >
      </div>
      
      <div class="users-table-container">
  `;

  // Mostrar tabla de usuarios con datos de reporte (sin email)
  if (reportData.usuarios && reportData.usuarios.length > 0) {
    adminHTML += `
      <table class="users-table" id="usersTable">
        <thead>
          <tr>
    `;
    
    // Crear headers de la tabla (excluyendo email si está presente)
    const displayHeaders = reportData.headers.filter(header => 
      !header.toLowerCase().includes('email') && 
      header !== 'Email' && 
      header !== 'EMAIL'
    );
    displayHeaders.forEach(header => {
      adminHTML += `<th>${header}</th>`;
    });
    
    adminHTML += `
          </tr>
        </thead>
        <tbody id="usersTableBody">
    `;
    
    // Crear filas con datos de todos los usuarios (sin email)
    reportData.usuarios.forEach(usuario => {
      adminHTML += '<tr>';
      displayHeaders.forEach(header => {
        const value = usuario[header] || '';
        adminHTML += `<td>${value}</td>`;
      });
      adminHTML += '</tr>';
    });
    
    adminHTML += `
        </tbody>
      </table>
    `;
  } else {
    adminHTML += `
      <div class="no-results">
        <p>No se encontraron usuarios registrados.</p>
      </div>
    `;
  }

  adminHTML += `
        </div>
      </div>
      
      <!-- Contenedor para tablas de usuario seleccionado -->
      <div id="selectedUserTables" style="display: none; margin-top: 30px;">
        <div class="admin-panel">
          <div class="admin-header">
            <h2>Datos del Usuario Seleccionado</h2>
            <button class="btn btn-primary" onclick="hideUserTables()">Volver</button>
          </div>
          <div id="userTablesContent"></div>
        </div>
      </div>
  `;

  appPage.innerHTML = adminHTML;
  
  // Agregar event listeners a las filas de la tabla
  const tableRows = document.querySelectorAll('#usersTableBody tr');
  tableRows.forEach(row => {
    row.addEventListener('click', function() {
      showUserTables(this);
    });
    row.style.cursor = 'pointer';
  });
}

// Función para mostrar las tablas de un usuario seleccionado por el administrador
async function showUserTables(rowElement) {
  // Obtener el nombre del usuario de la fila seleccionada (primera columna)
  const userName = rowElement.cells[0].textContent.trim();
  
  if (!userName) {
    alert('No se pudo obtener el nombre del usuario');
    return;
  }
  
  // Mostrar contenedor de tablas
  const selectedUserTables = document.getElementById('selectedUserTables');
  const userTablesContent = document.getElementById('userTablesContent');
  selectedUserTables.style.display = 'block';
  
  // Mostrar mensaje de carga
  userTablesContent.innerHTML = '<div class="loading">Cargando datos del usuario...</div>';
  
  try {
    // Obtener todas las tablas del usuario
    const adhData = await fetchUserADHData(userName);
    const quotesData = await fetchUserQuotesData(userName);
    const bonosData = await fetchUserBonosData(userName);
    const horasData = await fetchUserHorasData(userName);
    const excesosData = await fetchUserExcesosData(userName);
    const rankData = await fetchUserRankData(userName);
    
    // Generar HTML para las tablas
    let tablesHTML = `
      <div class="user-details" style="margin-bottom: 30px;">
        <div class="detail-card">
          <div class="detail-label">Nombre</div>
          <div class="detail-value">${userName}</div>
        </div>
        <div class="detail-card">
          <div class="detail-label">Rank</div>
          <div class="detail-value">${rankData.rank || 'N/A'}</div>
        </div>
      </div>
      
      <div class="user-tables-container">
    `;
    
    // Tabla ADH
    if (adhData.adhTable && adhData.adhTable.length > 0) {
      tablesHTML += `
        <div class="user-table-card">
          <h3>ADH</h3>
          <table class="user-data-table">
            <thead>
              <tr>
      `;
      
      adhData.adhTable[0].headers.forEach(header => {
        tablesHTML += `<th>${header}</th>`;
      });
      
      tablesHTML += `
              </tr>
            </thead>
            <tbody>
      `;
      
      adhData.adhTable[0].data.forEach(row => {
        tablesHTML += `<tr>`;
        row.forEach(cell => {
          tablesHTML += `<td>${cell}</td>`;
        });
        tablesHTML += `</tr>`;
      });
      
      tablesHTML += `
            </tbody>
          </table>
        </div>
      `;
    } else {
      tablesHTML += `
        <div class="user-table-card">
          <h3>ADH</h3>
          <p>No se encontraron datos para este usuario.</p>
        </div>
      `;
    }
    
    // Tabla Rendimiento (Quotes)
    if (quotesData.quotesTable && quotesData.quotesTable.length > 0) {
      tablesHTML += `
        <div class="user-table-card">
          <h3>Rendimiento</h3>
          <table class="user-data-table">
            <thead>
              <tr>
      `;
      
      quotesData.quotesTable[0].headers.forEach(header => {
        tablesHTML += `<th>${header}</th>`;
      });
      
      tablesHTML += `
              </tr>
            </thead>
            <tbody>
      `;
      
      quotesData.quotesTable[0].data.forEach(row => {
        tablesHTML += `<tr>`;
        row.forEach(cell => {
          tablesHTML += `<td>${cell}</td>`;
        });
        tablesHTML += `</tr>`;
      });
      
      tablesHTML += `
            </tbody>
          </table>
        </div>
      `;
    } else {
      tablesHTML += `
        <div class="user-table-card">
          <h3>Rendimiento</h3>
          <p>No se encontraron datos de rendimiento para este usuario.</p>
        </div>
      `;
    }
    
    // Tabla Bonos
    if (bonosData.bonosTable && bonosData.bonosTable.length > 0) {
      tablesHTML += `
        <div class="user-table-card">
          <h3>Bonos</h3>
          <table class="user-data-table">
            <thead>
              <tr>
      `;
      
      bonosData.bonosTable[0].headers.forEach(header => {
        tablesHTML += `<th>${header}</th>`;
      });
      
      tablesHTML += `
              </tr>
            </thead>
            <tbody>
      `;
      
      bonosData.bonosTable[0].data.forEach(row => {
        tablesHTML += `<tr>`;
        row.forEach(cell => {
          tablesHTML += `<td>${cell}</td>`;
        });
        tablesHTML += `</tr>`;
      });
      
      tablesHTML += `
            </tbody>
          </table>
        </div>
      `;
    } else {
      tablesHTML += `
        <div class="user-table-card">
          <h3>Bonos</h3>
          <p>No se encontraron datos de bonos para este usuario.</p>
        </div>
      `;
    }
    
    // Tabla Horas Diarias
    if (horasData.horasTable && horasData.horasTable.length > 0) {
      tablesHTML += `
        <div class="user-table-card">
          <h3>Horas Diarias</h3>
          <table class="user-data-table">
            <thead>
              <tr>
      `;
      
      horasData.horasTable[0].headers.forEach(header => {
        tablesHTML += `<th>${header}</th>`;
      });
      
      tablesHTML += `
              </tr>
            </thead>
            <tbody>
      `;
      
      horasData.horasTable[0].data.forEach(row => {
        tablesHTML += `<tr>`;
        row.forEach(cell => {
          tablesHTML += `<td>${cell}</td>`;
        });
        tablesHTML += `</tr>`;
      });
      
      tablesHTML += `
            </tbody>
          </table>
        </div>
      `;
    } else {
      tablesHTML += `
        <div class="user-table-card">
          <h3>Horas Diarias</h3>
          <p>No se encontraron datos de horas para este usuario.</p>
        </div>
      `;
    }
    
    // Tabla Excesos
    if (excesosData.excesosTable && excesosData.excesosTable.length > 0) {
      tablesHTML += `
        <div class="user-table-card">
          <h3>Excesos</h3>
          <table class="excesos-table">
            <thead>
              <tr>
                <th>Tipo</th>
      `;
      
      excesosData.excesosTable[0].headers.forEach(header => {
        tablesHTML += `<th>${header}</th>`;
      });
      
      tablesHTML += `
              </tr>
            </thead>
            <tbody>
      `;
      
      excesosData.excesosTable[0].data.forEach(rowData => {
        tablesHTML += `<tr>`;
        tablesHTML += `<td><strong>${rowData.label}</strong></td>`;
        rowData.values.forEach(cell => {
          tablesHTML += `<td>${cell}</td>`;
        });
        tablesHTML += `</tr>`;
      });
      
      tablesHTML += `
            </tbody>
          </table>
        </div>
      `;
    } else {
      tablesHTML += `
        <div class="user-table-card">
          <h3>Excesos</h3>
          <p>No se encontraron datos de excesos para este usuario.</p>
        </div>
      `;
    }
    
    tablesHTML += `
        </div>
      `;
    
    userTablesContent.innerHTML = tablesHTML;
    
    // Desplazarse hacia las tablas
    selectedUserTables.scrollIntoView({ behavior: 'smooth' });
    
  } catch (error) {
    console.error('Error al cargar datos del usuario:', error);
    userTablesContent.innerHTML = `
      <div class="error-message">
        Error al cargar los datos del usuario: ${error.message}
      </div>
    `;
  }
}

// Función para ocultar las tablas de usuario y volver a la vista principal
function hideUserTables() {
  const selectedUserTables = document.getElementById('selectedUserTables');
  selectedUserTables.style.display = 'none';
}

// Función para filtrar usuarios en el panel de administrador
function filterUsers() {
  const searchInput = document.getElementById('userSearch');
  const filter = searchInput.value.toLowerCase();
  const table = document.getElementById('usersTable');
  const tbody = document.getElementById('usersTableBody');
  
  if (!tbody) return;
  
  const rows = tbody.getElementsByTagName('tr');
  
  for (let i = 0; i < rows.length; i++) {
    const cells = rows[i].getElementsByTagName('td');
    let found = false;
    
    for (let j = 0; j < cells.length; j++) {
      if (cells[j].textContent.toLowerCase().indexOf(filter) > -1) {
        found = true;
        break;
      }
    }
    
    rows[i].style.display = found ? '' : 'none';
  }
}

function logout() {
  // Limpiar todo antes de cerrar sesión
  clearLoginMessage();
  showLogin();
}

function showLoading(message) {
  document.getElementById('login-message').innerHTML = `
    <div class="loading">
      <div class="spinner"></div>
      <span>${message}</span>
    </div>
  `;
}

function showError(message) {
  document.getElementById('login-message').innerHTML = `
    <div class="error-message">
      ${message}
    </div>
  `;
}

function showSuccess(message) {
  document.getElementById('login-message').innerHTML = `
    <div class="success-message">
      ${message}
    </div>
  `;
}

// Initialize email section functionality
function initializeEmailSection() {
  const fieldsData = {
    "GENERAL NOTES": [
      { label: "Notes", id: "notes", type: "textarea" },
    ],
    "CANCEL / RESCHEDULE": [
      { label: "Customer Name", id: "customerName", type: "text" },
      { label: "Job #", id: "jobNumber", type: "text" },
      { label: "Phone Number", id: "phoneNumber", type: "text" },
      { label: "Appointment Date", id: "date", type: "date"},
      { label: "New Appointment Date", id: "newDate", type: "date"},
      { label: "Situation", id: "situation", type: "select", options: ["cancel", "reschedule"] },
      { label: "Reason", id: "reason", type: "textarea" }
    ],
    "BURCO MIRROR QUOTE AND ETA REQUEST": [
      { label: "Customer Name", id: "customerName", type: "text" },
      { label: "Mirror", id: "mirror", type: "select", options: ["driver side", "passenger side"] },
      { label: "Job #", id: "jobNumber", type: "text" },
      { label: "VIN", id: "vin", type: "text" },
      { label: "Phone Number", id: "phoneNumber", type: "text" },
      { label: "Burco #", id: "burcoNumber", type: "text" }
    ],
    "WARRANTY": [
      { label: "Customer Name", id: "customerName", type: "text" },
      { label: "Complaint", id: "complaint", type: "textarea" },
      { label: "Job #", id: "jobNumber", type: "text" },
      { label: "VIN", id: "vin", type: "text" },
      { label: "Phone Number", id: "phoneNumber", type: "text" },
      { label: "Email", id: "email", type: "text" },
      { label: "Year Make and Model", id: "yearMakeModel", type: "text" }
    ]
  };

  function createField(field) {
    const fieldDiv = document.createElement("div");
    fieldDiv.classList.add("form-group");

    const label = document.createElement("label");
    label.setAttribute("for", field.id);
    label.textContent = field.label + ":";

    let input;
    switch (field.type) {
      case "select":
        input = document.createElement("select");
        input.id = field.id;
        input.name = field.id;
        const defaultOption = document.createElement("option");
        defaultOption.value = "";
        defaultOption.textContent = `Select ${field.label}`;
        input.appendChild(defaultOption);
        field.options.forEach(option => {
          const optionElement = document.createElement("option");
          optionElement.value = option;
          optionElement.textContent = option;
          input.appendChild(optionElement);
        });

        if (field.subOptions) {
          const subField = document.createElement("select");
          subField.id = field.id + "-sub";
          subField.name = field.id + "-sub";
          subField.style.display = 'none';
          subField.classList.add("form-control");

          Object.keys(field.subOptions).forEach(option => {
            const subOptions = field.subOptions[option];
            subOptions.forEach(subOption => {
              const subOptionElement = document.createElement("option");
              subOptionElement.value = subOption;
              subOptionElement.textContent = subOption;
              subField.appendChild(subOptionElement);
            });
          });

          input.addEventListener('change', function() {
            if (field.subOptions[this.value]) {
              subField.style.display = 'block';
            } else {
              subField.style.display = 'none';
            }
          });

          fieldDiv.appendChild(subField);
        }
        
        // Add event listener for situation field to toggle newDate visibility
        if (field.id === "situation") {
          input.addEventListener('change', function() {
            const newDateField = document.getElementById("newDate");
            const newDateLabel = document.querySelector('label[for="newDate"]');
            if (newDateField && newDateLabel) {
              if (this.value === "cancel") {
                newDateField.parentElement.style.display = "none";
              } else {
                newDateField.parentElement.style.display = "flex";
              }
            }
          });
        }
        break;
      case "textarea":
        input = document.createElement("textarea");
        input.id = field.id;
        input.name = field.id;
        input.rows = 3;
        input.classList.add("form-control");
        break;
      case "date":
        input = document.createElement("input");
        input.type = "text";
        input.id = field.id;
        input.name = field.id;
        input.classList.add("form-control");
        flatpickr(input, {
          dateFormat: "m/d/Y"
        });
        break;
      case "number":
        input = document.createElement("input");
        input.type = "number";
        input.id = field.id;
        input.name = field.id;
        input.classList.add("form-control");
        break;
      case "text":
      default:
        input = document.createElement("input");
        input.type = "text";
        input.id = field.id;
        input.name = field.id;
        input.classList.add("form-control");
        break;
    }

    fieldDiv.appendChild(label);
    fieldDiv.appendChild(input);
    return fieldDiv;
  }

  // Event listeners para los cambios en la categoría, el issue y el botón de generar
  document.getElementById("issue").addEventListener("change", updateFields);
  document.getElementById("generateButton").addEventListener("click", generateEmail);
  document.getElementById("copyButton").addEventListener("click", copyEmail);

  function updateFields() {
    const issue = document.getElementById("issue").value;
    const fieldsContainer = document.getElementById("fields");
    const emailContent = document.getElementById("emailContent");

    // Limpia los campos y el contenido del email
    fieldsContainer.innerHTML = "";
    emailContent.innerHTML = "";
    emailContent.classList.add("hidden");

    if (fieldsData[issue]) {
      fieldsData[issue].forEach(field => {
        const fieldDiv = createField(field);
        fieldsContainer.appendChild(fieldDiv);
      });
      fieldsContainer.classList.remove("hidden");
    } else {
      fieldsContainer.classList.add("hidden");
    }
  }

  function clearForm() {
    // Limpiar todos los campos generados dinámicamente excepto los select
    const fieldsContainer = document.getElementById("fields");
    const inputs = fieldsContainer.querySelectorAll("input, textarea");
    inputs.forEach(input => {
      input.value = "";
    });
  }
  
  function validateFields() {
    const issue = document.getElementById("issue").value;
    const requiredFields = document.querySelectorAll('#fields input, #fields select, #fields textarea');
    
    for (const field of requiredFields) {
      // Skip validation for newDate if situation is cancel
      if (issue === "CANCEL / RESCHEDULE" && field.id === "newDate") {
        const situation = document.getElementById("situation");
        if (situation && situation.value === "cancel") {
          continue;
        }
      }
      
      if (field.value.trim() === '') {
        field.classList.add('is-invalid');
        field.scrollIntoView({ behavior: 'smooth', block: 'center' });
        return false;
      } else {
        field.classList.remove('is-invalid');
      }
    }
    return true;
  }

  function generateEmail() {
    if (!validateFields()) {
      Swal.fire({
        icon: 'warning',
        title: 'Incomplete Form',
        text: 'Please fill in all required fields before generating the email.'
      });
      return; // Detiene la generación si falta algún campo
    }
    const issue = document.getElementById("issue").value;
    const csrName = document.getElementById("name").value;
    const customerName = document.getElementById("customerName") ? document.getElementById("customerName").value : '';
    const jobNumber = document.getElementById("jobNumber") ? document.getElementById("jobNumber").value : '';
    const vin = document.getElementById("vin") ? document.getElementById("vin").value : '';
    const phoneNumber = document.getElementById("phoneNumber") ? document.getElementById("phoneNumber").value : '';
    const zipCode = document.getElementById("zipCode") ? document.getElementById("zipCode").value : '';
    const email = document.getElementById("email") ? document.getElementById("email").value : '';
    const yearMakeModel = document.getElementById("yearMakeModel") ? document.getElementById("yearMakeModel").value : '';
    const part = document.getElementById("part") ? document.getElementById("part").value : '';

    let emailText = `Hi, this is ${csrName} from the call center. See full customer’s information at the end of the email.\n\n`;

    if (issue === "BURCO MIRROR QUOTE AND ETA REQUEST") {
      const mirror = document.getElementById("mirror") ? document.getElementById("mirror").value : '';
      const burcoNumber = document.getElementById("burcoNumber") ? document.getElementById("burcoNumber").value : '';
      emailText += `Mr/Mrs ${customerName} called in because they would like to know if we can get a ${mirror} mirror. Can you help me get a price and ETA from your vendors and let me know how much should the quote be?\n\n`;
      emailText += `Thanks\n\n`;
      emailText += `Job #: ${jobNumber}\nVIN: ${vin}\nPhone number: ${phoneNumber}\nBurco Redi-Cut#${burcoNumber} ${mirror}\n`;
    } else if (issue === "WARRANTY") {
      const complaint = document.getElementById("complaint") ? document.getElementById("complaint").value : '';
      emailText += `Customer is having an issue with the job previously performed at our shop and wants to apply their warranty. They are claiming that ${complaint}.\n\n`;
      emailText += `Could you please let me know if I can follow up with this claim or contact the customer.\n\n`;
      emailText += `Much appreciated.\n\n`;
      emailText += `Job #: ${jobNumber}\nName: ${customerName}\nVIN: ${vin}\nPhone number: ${phoneNumber}\nYear/make/model: ${yearMakeModel}\nEmail: ${email}\n`;
    } else if (issue === "CANCEL / RESCHEDULE") {
      const situation = document.getElementById("situation") ? document.getElementById("situation").value : '';
      const reason = document.getElementById("reason") ? document.getElementById("reason").value : '';
      const date = document.getElementById("date") ? document.getElementById("date").value : '';
      const newDate = document.getElementById("newDate") ? document.getElementById("newDate").value : '';
      
      emailText += `Customer ${customerName} would like to ${situation} their appointment.\n\n`;
      emailText += `Reason: ${reason}\n\n`;
      emailText += `Original appointment date: ${date}\n`;
      if (situation === "reschedule") {
        emailText += `New requested appointment date: ${newDate}\n\n`;
      }
      emailText += `Job #: ${jobNumber}\nPhone number: ${phoneNumber}\n`;
    } else if (issue === "GENERAL NOTES") {
      const notes = document.getElementById("notes") ? document.getElementById("notes").value : '';
      emailText += `General notes regarding customer interaction:\n\n${notes}\n\n`;
      emailText += `Job #: ${jobNumber}\nCustomer Name: ${customerName}\nPhone number: ${phoneNumber}\n`;
    }

    const emailContentElement = document.getElementById("emailContent");
    emailContentElement.textContent = emailText;
    emailContentElement.classList.remove("hidden");
    
    // Limpiar todos los campos del formulario
    clearForm();
  }

  function copyEmail() {
    const emailContent = document.getElementById("emailContent");
    if (emailContent.classList.contains("hidden")) {
      alert("No email content to copy. Please generate an email first.");
      return;
    }
    
    const textToCopy = emailContent.textContent;
    
    navigator.clipboard.writeText(textToCopy).then(() => {
      // Mostrar mensaje de éxito
      const copyMessage = document.getElementById("copyMessage");
      copyMessage.style.display = "block";
      
      // Ocultar mensaje después de 3 segundos
      setTimeout(() => {
        copyMessage.style.display = "none";
      }, 3000);
    }).catch(err => {
      console.error("Error copying text: ", err);
      alert("Error copying text. Please try again.");
    });
  }
}

// Initialize store locations section
function initializeStoreLocationsSection() {
  let map;
  let markers = [];
  const LOCATIONIQ_API_KEY = 'pk.07d987a87bff1e4be589385ba1001be3';
  const GOOGLE_SHEETS_API_KEY = 'AIzaSyC8el04dXppUCvNp2Ok9R6mOUmyP_oonnY';
  const SHEET_ID = '1Q4JL8aeUiri9bwJ-hrxf6wdZtW9PLxGcBLDPRK1FIMM';
  const MAX_TRAVEL_TIME = 170; // 2 horas y 50 minutos en minutos

  // Inicializar el mapa
  function initMap() {
    if (!map) {
      map = L.map('map').setView([39.8283, -98.5795], 4);
      L.tileLayer('https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png', {
        attribution: '© OpenStreetMap contributors'
      }).addTo(map);
    }
  }

  // Buscar tiendas
  async function searchStores() {
    const zipcode = document.getElementById('zipcode').value;
    if(!zipcode || zipcode.length !== 5) {
      Swal.fire({
        icon: 'error',
        title: 'Invalid ZIP Code',
        text: 'Please enter a valid 5-digit ZIP code'
      });
      return;
    }

    try {
      const locationResponse = await fetch(`https://us1.locationiq.com/v1/search.php?key=${LOCATIONIQ_API_KEY}&q=${zipcode},USA&format=json`);
      const locationData = await locationResponse.json();

      if(!locationData[0]) {
        Swal.fire({
          icon: 'error',
          title: 'Invalid ZIP Code',
          text: 'No location found for this ZIP code'
        });
        return;
      }

      const userLat = parseFloat(locationData[0].lat);
      const userLon = parseFloat(locationData[0].lon);

      const sheetResponse = await fetch(
        `https://sheets.googleapis.com/v4/spreadsheets/${SHEET_ID}/values/Locations?key=${GOOGLE_SHEETS_API_KEY}`
      );
      const sheetData = await sheetResponse.json();

      const stores = sheetData.values.slice(1).map(row => ({
        name: row[2],
        phone: row[7],
        address: row[12],
        lat: parseFloat(row[10]),
        lon: parseFloat(row[11]),
        distance: calculateDistance(userLat, userLon, parseFloat(row[10]), parseFloat(row[11]))
      }));

      const maxDistance = (MAX_TRAVEL_TIME / 60) * 50;
      const nearbyStores = stores
        .filter(store => store.distance <= maxDistance)
        .sort((a, b) => a.distance - b.distance);

      displayStores(nearbyStores, userLat, userLon);
    } catch (error) {
      console.error('Error:', error);
      Swal.fire({
        icon: 'error',
        title: 'Error',
        text: 'Error finding stores. Please try again.'
      });
    }
  }

  function calculateDistance(lat1, lon1, lat2, lon2) {
    const R = 3959;
    const dLat = toRad(lat2 - lat1);
    const dLon = toRad(lon2 - lon1);
    const a = Math.sin(dLat/2) * Math.sin(dLat/2) +
            Math.cos(toRad(lat1)) * Math.cos(toRad(lat2)) * 
            Math.sin(dLon/2) * Math.sin(dLon/2);
    const c = 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1-a));
    return R * c;
  }

  function toRad(degrees) {
    return degrees * (Math.PI/180);
  }

  function calculateTravelTime(distance) {
    const speedMPH = 50;
    const hours = distance / speedMPH;
    const minutes = Math.round(hours * 60);
    
    if (minutes < 60) {
      return `${minutes} min`;
    } else {
      const wholeHours = Math.floor(hours);
      const remainingMinutes = Math.round((hours - wholeHours) * 60);
      if (remainingMinutes === 0) {
        return `${wholeHours} hr`;
      }
      return `${wholeHours}h ${remainingMinutes}m`;
    }
  }

  function displayStores(stores, userLat, userLon) {
    markers.forEach(marker => map.removeLayer(marker));
    markers = [];

    map.setView([userLat, userLon], 9);

    const userIcon = L.icon({
      iconUrl: 'https://raw.githubusercontent.com/pointhi/leaflet-color-markers/master/img/marker-icon-2x-red.png',
      iconSize: [25, 41],
      iconAnchor: [12, 41],
      popupAnchor: [1, -34],
      shadowUrl: 'https://cdnjs.cloudflare.com/ajax/libs/leaflet/0.7.7/images/marker-shadow.png',
      shadowSize: [41, 41]
    });

    const userMarker = L.marker([userLat, userLon], { icon: userIcon }).addTo(map);
    markers.push(userMarker);

    const storeList = document.getElementById('storeList');
    storeList.innerHTML = '';

    if(stores.length === 0) {
      storeList.innerHTML = '<div class="text-center p-3">No stores found within 2h 50m.</div>';
      return;
    }

    stores.forEach(store => {
      const travelTime = calculateTravelTime(store.distance);
      
      const marker = L.marker([store.lat, store.lon])
        .bindPopup(
          `<b>${store.name}</b><br>
          <small>${store.address}</small><br>
          <i class="bi bi-telephone"></i> ${store.phone}<br>
          <i class="bi bi-geo-alt"></i> ${store.distance.toFixed(1)} mi<br>
          <i class="bi bi-clock"></i> ${travelTime}`
        )
        .addTo(map);
      markers.push(marker);

      const storeDiv = document.createElement('div');
      storeDiv.className = 'store-item';
      
      const viewButton = document.createElement('button');
      viewButton.className = 'btn-map';
      viewButton.innerHTML = '<i class="bi bi-map"></i> View';
      viewButton.type = 'button';
      
      viewButton.addEventListener('click', (e) => {
        e.preventDefault();
        map.setView([store.lat, store.lon], 13);
      });
      
      storeDiv.innerHTML = 
        `<div class="store-info">
          <strong>${store.name}</strong><br>
          <small class="text-muted">${store.address}</small><br>
          <span><i class="bi bi-telephone"></i> ${store.phone}</span><br>
          <span><i class="bi bi-geo-alt"></i> ${store.distance.toFixed(1)} mi</span>
          <div class="estimated-time">
            <i class="bi bi-clock"></i> ${travelTime}
          </div>
        </div>`;
      
      storeDiv.appendChild(viewButton);
      storeList.appendChild(storeDiv);
    });
  }

  // Initialize with map
  initMap();

  // Agregar manejador para el botón de búsqueda
  const searchButton = document.getElementById('searchButton');
  searchButton.addEventListener('click', async (event) => {
    event.preventDefault();
    await searchStores();
  });

  // Permitir búsqueda al presionar Enter
  const searchInput = document.getElementById('zipcode');
  searchInput.addEventListener('keypress', (event) => {
    if (event.key === 'Enter') {
      event.preventDefault();
      searchStores();
    }
  });
}

// Initialize profit calculator section
function initializeProfitCalculatorSection() {
  // Toggle button handlers
  document.querySelectorAll('.category-btn').forEach(button => {
    button.addEventListener('click', function() {
      // Get parent container
      const parent = this.parentElement;
      
      // Remove active class from siblings
      parent.querySelectorAll('.category-btn').forEach(btn => {
        btn.classList.remove('active');
      });
      
      // Add active class to clicked button
      this.classList.add('active');
      
      // Update hidden input
      const hiddenInput = parent.parentElement.querySelector('input[type="hidden"]');
      hiddenInput.value = this.dataset.value;
    });
  });

  // Add part function
  window.addPart = function() {
    const partsContainer = document.getElementById('partsContainer');
    const partNumber = partsContainer.children.length + 1;
    
    const formRow = document.createElement('div');
    formRow.className = 'form-row';
    formRow.innerHTML = `
      <input type="number" class="part-cost" placeholder="Part ${partNumber} cost" min="0" step="0.01">
      <button type="button" class="remove-part" onclick="removePart(this)">
        <i class="bi bi-trash"></i>
      </button>
    `;
    
    partsContainer.appendChild(formRow);
  }

  // Remove part function
  window.removePart = function(button) {
    if (document.querySelectorAll('.form-row').length > 1) {
      button.parentElement.remove();
      // Renumber remaining parts
      const parts = document.querySelectorAll('.part-cost');
      parts.forEach((part, index) => {
        part.placeholder = `Part ${index + 1} cost`;
      });
    }
  }

  // Reset form function
  window.resetForm = function() {
    // Clear all numeric inputs
    document.querySelectorAll('input[type="number"]').forEach(input => {
      input.value = '';
    });
    
    // Reset toggles
    document.querySelectorAll('.category-btn').forEach(button => {
      button.classList.remove('active');
    });
    
    // Set default toggles to active
    document.querySelectorAll('.category-btn[data-value="inshop"], .category-btn[data-value="all"]')
      .forEach(button => {
        button.classList.add('active');
      });
    
    // Reset hidden inputs
    document.getElementById('serviceType').value = 'inshop';
    document.getElementById('state').value = 'all';
    
    // Reset to only one part
    const partsContainer = document.getElementById('partsContainer');
    partsContainer.innerHTML = `
      <div class="form-row">
        <input type="number" class="part-cost" placeholder="Part 1 cost" min="0" step="0.01">
        <button type="button" class="remove-part" onclick="removePart(this)">
          <i class="bi bi-trash"></i>
        </button>
      </div>
    `;
    
    // Hide results
    document.getElementById('results').classList.add('hidden');
  }

  // Calculate margin function
  window.calculateMargin = function() {
    const subtotal = parseFloat(document.getElementById('subtotal').value);
    if (isNaN(subtotal)) {
      alert('Please enter a valid subtotal amount');
      return;
    }
    
    const serviceType = document.getElementById('serviceType').value;
    const state = document.getElementById('state').value;
    
    const parts = document.querySelectorAll('.part-cost');
    const recalibration = parseFloat(document.getElementById('recalibration').value) || 0;
    const moulding = parseFloat(document.getElementById('moulding').value) || 0;
    
    const resultsBody = document.getElementById('resultsBody');
    resultsBody.innerHTML = '';
    
    let totalCost = 0;
    
    // Calculate part costs
    parts.forEach((partInput, index) => {
      const baseCost = parseFloat(partInput.value) || 0;
      let adjustment = 0;
      
      if (state === 'indiana') {
        adjustment = 150;
      } else if (state === 'all') {
        adjustment = serviceType === 'inshop' ? 225 : 275;
      }
      
      const total = baseCost + adjustment;
      totalCost += total;
      
      const row = document.createElement('tr');
      row.innerHTML = `
        <td>Part ${index + 1}</td>
        <td>$${baseCost.toFixed(2)}</td>
        <td>$${adjustment.toFixed(2)}</td>
        <td>$${total.toFixed(2)}</td>
      `;
      resultsBody.appendChild(row);
    });
    
    // Add recalibration
    totalCost += recalibration;
    if (recalibration > 0) {
      const row = document.createElement('tr');
      row.innerHTML = `
        <td>Recal</td>
        <td>$${recalibration.toFixed(2)}</td>
        <td>$0.00</td>
        <td>$${recalibration.toFixed(2)}</td>
      `;
      resultsBody.appendChild(row);
    }
    
    // Add moulding cost
    let mouldingTotal = moulding;
    if (moulding > 0) {
      mouldingTotal += 20;
      totalCost += mouldingTotal;
      
      const row = document.createElement('tr');
      row.innerHTML = `
        <td>Mould</td>
        <td>$${moulding.toFixed(2)}</td>
        <td>$20.00</td>
        <td>$${mouldingTotal.toFixed(2)}</td>
      `;
      resultsBody.appendChild(row);
    }
    
    // Add total costs row
    const totalRow = document.createElement('tr');
    totalRow.innerHTML = `
      <td><strong>Total</strong></td>
      <td></td>
      <td></td>
      <td><strong>$${totalCost.toFixed(2)}</strong></td>
    `;
    resultsBody.appendChild(totalRow);
    
    // Add subtotal row
    const subtotalRow = document.createElement('tr');
    subtotalRow.innerHTML = `
      <td><strong>Subtotal</strong></td>
      <td></td>
      <td></td>
      <td><strong>$${subtotal.toFixed(2)}</strong></td>
    `;
    resultsBody.appendChild(subtotalRow);
    
    // Calculate margin
    const margin = subtotal - totalCost;
    const differenceRow = document.createElement('tr');
    differenceRow.innerHTML = `
      <td><strong>Margin</strong></td>
      <td></td>
      <td></td>
      <td><strong>$${margin.toFixed(2)}</strong></td>
    `;
    resultsBody.appendChild(differenceRow);
    
    // Show results
    document.getElementById('results').classList.remove('hidden');
    
    // Display margin message
    const marginMessage = document.getElementById('marginMessage');
    if (margin < 0) {
      marginMessage.textContent = `Negative margin of $${Math.abs(margin).toFixed(2)} – No discount`;
      marginMessage.className = "margin-message error";
    } else {
      marginMessage.textContent = `Positive margin of $${margin.toFixed(2)} – Discount available`;
      marginMessage.className = "margin-message success";
    }
  }
}

// Initialize callbacks section
function initializeCallbacksSection() {
  let callbacks = [];

  // Load callbacks from localStorage
  function loadCallbacks() {
    const savedCallbacks = localStorage.getItem('callbacks');
    if (savedCallbacks) {
      callbacks = JSON.parse(savedCallbacks);
    } else {
      callbacks = [];
    }
    renderCallbacks();
  }

  // Save callbacks to localStorage
  function saveCallbacks() {
    localStorage.setItem('callbacks', JSON.stringify(callbacks));
  }

  window.addCallback = function() {
    const number = document.getElementById('callbackNumber').value.trim();
    const notes = document.getElementById('callbackNotes').value.trim();
    
    if (!number) {
      Swal.fire({
        icon: 'warning',
        title: 'Incomplete Form',
        text: 'Please enter at least phone number.'
      });
      return;
    }
    
    const callback = {
      id: Date.now(),
      number: number,
      notes: notes,
      timestamp: new Date()
    };
    
    callbacks.push(callback);
    saveCallbacks();
    renderCallbacks();
    
    // Clear form fields
    document.getElementById('callbackNumber').value = '';
    document.getElementById('callbackNotes').value = '';
  }

  function removeCallback(id) {
    callbacks = callbacks.filter(callback => callback.id !== id);
    saveCallbacks();
    renderCallbacks();
  }

  function renderCallbacks() {
    const callbacksList = document.getElementById('callbacksList');
    callbacksList.innerHTML = '';
    
    if (callbacks.length === 0) {
      callbacksList.innerHTML = '<div class="text-center p-3 text-muted">No callbacks scheduled</div>';
      return;
    }
    
    callbacks.forEach(callback => {
      const callbackItem = document.createElement('div');
      callbackItem.className = 'callback-item';
      callbackItem.innerHTML = `
        <div class="callback-info">
          <strong>${callback.number}</strong>
          ${callback.notes ? `<br><small class="text-muted">${callback.notes}</small>` : ''}
        </div>
        <button type="button" class="callback-check" onclick="removeCallback(${callback.id})">
          <i class="bi bi-check"></i>
        </button>
      `;
      callbacksList.appendChild(callbackItem);
    });
  }

  // Load callbacks when section is shown
  loadCallbacks();
}

// Initialize ranking section
function initializeRankingSection() {
  // CONFIGURACIÓN DE LA API
  const RANKING_SHEET_ID = "1vTRE0WMWqc-_GgCKDlYX6YRkXeasKuYIU1Rza2nvI9s";
  const RANKING_API_KEY = "AIzaSyCy5c6rOI8YmuzskW_IunhwY_q8P8HW3xs";
  const RANKING_RANGE = "Update!B2:F20";
  const GOOGLE_SHEETS_URL = `https://sheets.googleapis.com/v4/spreadsheets/${RANKING_SHEET_ID}/values/${RANKING_RANGE}?key=${RANKING_API_KEY}`;

  // Variables globales
  let agentsData = [];

  // Event listener para el botón de refresh
  document.getElementById('refreshBtn').addEventListener('click', function() {
    loadDashboardData();
  });

  // Funciones de datos
  async function loadDashboardData() {
    try {
      // Mostrar indicador de carga
      showLoadingIndicators();
      
      const response = await fetch(GOOGLE_SHEETS_URL);
      const data = await response.json();
      
      // Procesar datos de Google Sheets
      agentsData = processSheetData(data.values);
      
      // Filtrar agentes con al menos 1 llamada
      agentsData = agentsData.filter(agent => agent.calls >= 1);
      
      // Calcular métricas
      calculateMetrics();
      
      // Renderizar UI
      renderMetrics();
      renderAgentsTable();
      renderTopAgents();
      
    } catch (error) {
      console.error('Error loading data:', error);
      // Mostrar mensaje de error
      document.getElementById('agentsTableBody').innerHTML = `
        <tr>
          <td colspan="5" class="text-center text-danger">
            Error loading data. Please try again.
          </td>
        </tr>
      `;
    }
  }

  function showLoadingIndicators() {
    document.getElementById('agentsTableBody').innerHTML = `
      <tr>
        <td colspan="5" class="text-center">
          <div class="loading">
            <div class="spinner-border text-primary" role="status">
              <span class="visually-hidden">Loading...</span>
            </div>
          </div>
        </td>
      </tr>
    `;
    
    document.getElementById('topAgentsList').innerHTML = `
      <div class="loading">
        <div class="spinner-border spinner-border-sm text-primary" role="status">
          <span class="visually-hidden">Loading...</span>
        </div>
      </div>
    `;
  }

  function processSheetData(rows) {
    if (!rows || rows.length === 0) return [];
    
    return rows.map((row, index) => ({
      id: index + 1,
      name: row[0] || '',
      calls: parseInt(row[1]) || 0,
      wos: parseInt(row[2]) || 0,
      quotes: parseInt(row[3]) || 0
    }));
  }

  function calculateMetrics() {
    const totalCalls = agentsData.reduce((sum, agent) => sum + agent.calls, 0);
    const totalWOs = agentsData.reduce((sum, agent) => sum + agent.wos, 0);
    const totalQuotes = agentsData.reduce((sum, agent) => sum + agent.quotes, 0);
    
    // Calcular conversión promedio (WOs / Llamadas)
    const conversionRates = agentsData.map(agent => 
      agent.calls > 0 ? (agent.wos / agent.calls) * 100 : 0
    );
    const avgConversion = conversionRates.length > 0 
      ? conversionRates.reduce((sum, rate) => sum + rate, 0) / conversionRates.length 
      : 0;
    
    // Guardar métricas globales
    window.metrics = {
      totalCalls,
      totalWOs,
      totalQuotes,
      avgConversion: avgConversion.toFixed(1)
    };
  }

  // Funciones de renderizado
  function renderMetrics() {
    document.getElementById('totalCalls').textContent = window.metrics.totalCalls;
    document.getElementById('totalWOs').textContent = window.metrics.totalWOs;
    document.getElementById('totalQuotes').textContent = window.metrics.totalQuotes;
    document.getElementById('avgConversion').textContent = window.metrics.avgConversion + '%';
  }

  function renderAgentsTable() {
    const tbody = document.getElementById('agentsTableBody');
    tbody.innerHTML = '';
    
    if (agentsData.length === 0) {
      tbody.innerHTML = `
        <tr>
          <td colspan="5" class="text-center">
            No agents with calls recorded
          </td>
        </tr>
      `;
      return;
    }
    
    agentsData.forEach(agent => {
      // Calcular conversión como WOs / Llamadas
      const conversionRate = agent.calls > 0 ? ((agent.wos / agent.calls) * 100).toFixed(1) : 0;
      const row = document.createElement('tr');
      
      row.innerHTML = `
        <td>${agent.name}</td>
        <td>${agent.calls}</td>
        <td>${agent.wos}</td>
        <td>${agent.quotes}</td>
        <td>
          ${conversionRate}%
          <div class="conversion-bar" style="width: ${Math.min(conversionRate, 100)}%"></div>
        </td>
      `;
      
      tbody.appendChild(row);
    });
  }

  function renderTopAgents() {
    const topAgentsList = document.getElementById('topAgentsList');
    topAgentsList.innerHTML = '';
    
    if (agentsData.length === 0) {
      topAgentsList.innerHTML = '<div class="text-center p-3">No agents with calls recorded</div>';
      return;
    }
    
    // Ordenar por WOs y Quotes (combinación de ambos)
    const agentsWithScore = agentsData.map(agent => ({
      ...agent,
      score: agent.wos + agent.quotes // Puntuación basada en WOs + Quotes
    })).sort((a, b) => b.score - a.score).slice(0, 5);
    
    agentsWithScore.forEach((agent, index) => {
      const agentElement = document.createElement('div');
      agentElement.className = 'top-agent';
      agentElement.innerHTML = `
        <div class="d-flex justify-content-between align-items-center">
          <div>
            <div class="fw-bold">${agent.name}</div>
            <div class="small text-muted">${agent.wos} WOs | ${agent.quotes} Quotes</div>
          </div>
          <div class="badge bg-primary rounded-pill">${index + 1}</div>
        </div>
      `;
      topAgentsList.appendChild(agentElement);
    });
  }

  // Load initial data
  loadDashboardData();
}
