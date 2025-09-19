// auto-fill.js
const fs = require("fs");
const XLSX = require("xlsx");
const puppeteer = require("puppeteer-core");

const CFG = {
  EXCEL_FILE: "usuarios.xlsx",
  SHEET_NAME: null,
  TARGET_URL: "https://crimsonstrauss.xyz/expansao",
  CLICK_SUBMIT: true,
  HEADLESS: false,
  NAVIGATION_TIMEOUT: 20000,
  MIN_DELAY_MS: 2000,
  MAX_DELAY_MS: 10000,
  WAIT_AFTER_CLICK_MS: 25000,
  LOG_FILE: "autofill_log.csv",
  MAX_RETRIES: 3,
  MAX_ACTIVITY_WAIT_MINUTES: 15, // Aumentado para 15 minutos
};

// Cores para logs
const colors = {
  reset: '\x1b[0m',
  bright: '\x1b[1m',
  red: '\x1b[31m',
  green: '\x1b[32m',
  yellow: '\x1b[33m',
  blue: '\x1b[34m',
  magenta: '\x1b[35m',
  cyan: '\x1b[36m',
  white: '\x1b[37m'
};

function sleep(ms) {
  return new Promise((res) => setTimeout(res, ms));
}

function randBetween(min, max) {
  return Math.floor(Math.random() * (max - min + 1)) + min;
}

function detectColumn(keys, possibles) {
  const lower = keys.map((k) => k.toLowerCase());
  for (const p of possibles) {
    const idx = lower.findIndex((k) => k.includes(p));
    if (idx !== -1) return keys[idx];
  }
  return null;
}

function safeText(v) {
  return v === undefined || v === null ? "" : String(v).trim();
}

function appendLog(line) {
  fs.appendFileSync(CFG.LOG_FILE, line + "\n");
}

function logUser(userId, message, color = 'white') {
  console.log(`${colors.bright}${colors.blue}[USER ${userId}]${colors.reset} ${colors[color]}${message}${colors.reset}`);
}

function logCourse(userId, courseName, message, color = 'white') {
  console.log(`${colors.bright}${colors.blue}[USER ${userId}]${colors.reset} ${colors.green}[${courseName}]${colors.reset} ${colors[color]}${message}${colors.reset}`);
}

function logError(userId, message) {
  console.log(`${colors.bright}${colors.blue}[USER ${userId}]${colors.reset} ${colors.red}❌ ${message}${colors.reset}`);
}

function logSuccess(userId, message) {
  console.log(`${colors.bright}${colors.blue}[USER ${userId}]${colors.reset} ${colors.green}✅ ${message}${colors.reset}`);
}

function logWarning(userId, message) {
  console.log(`${colors.bright}${colors.blue}[USER ${userId}]${colors.reset} ${colors.yellow}⚠️ ${message}${colors.reset}`);
}

async function reopenCoursesModal(mainPage, userId, retries = CFG.MAX_RETRIES) {
  for (let attempt = 1; attempt <= retries; attempt++) {
    try {
      logUser(userId, `Tentativa ${attempt}/${retries} de abrir modal de cursos...`, 'cyan');
      
      await mainPage.evaluate(() => {
        const btn = document.querySelector("#buscarCursosBtn");
        if (btn) btn.click();
      });
      
      await mainPage.waitForSelector("#coursesModal", {
        visible: true,
        timeout: 10000
      });
      
      await sleep(2000);
      logUser(userId, "Modal reaberto com sucesso!", 'green');
      return true;
      
    } catch (error) {
      logWarning(userId, `Tentativa ${attempt} falhou: ${error.message}`);
      
      await mainPage.evaluate(() => {
        const closeBtn = document.querySelector("#closeCoursesModal, .close-modal");
        if (closeBtn) closeBtn.click();
      });
      
      await sleep(2000);
      
      if (attempt === retries) {
        logError(userId, "Não foi possível reabrir o modal após todas as tentativas");
        return false;
      }
    }
  }
}
async function waitForActivitiesCompletion(mainPage, userId, courseName) {
  let startTime = Date.now();
  const maxWaitTime = CFG.MAX_ACTIVITY_WAIT_MINUTES * 60 * 1000;
  let lastProgress = '';
  let consecutiveCompletions = 0;
  let activityCount = 0;
  let completedCount = 0;

  logCourse(userId, courseName, "Aguardando conclusão das atividades...", 'yellow');

  while (Date.now() - startTime < maxWaitTime) {
    try {
      // Aguarda um tempo mínimo antes da primeira verificação
      if (Date.now() - startTime < 30000) { // 30 segundos mínimos
        await sleep(5000);
        continue;
      }

      // Verifica o progresso atual de forma mais específica
      const progressInfo = await mainPage.evaluate(() => {
        // Procura por elementos que mostram progresso
        const progressElements = [
          document.querySelector('[class*="progress"]'),
          document.querySelector('[class*="status"]'),
          document.querySelector('[class*="completion"]'),
          document.querySelector('body').textContent.match(/\d+\s*\/\s*\d+/g)
        ].filter(Boolean);

        // Procura por mensagens de conclusão
        const completionIndicators = [
          'processamento concluído',
          'todas as atividades concluídas',
          '100% completo',
          'curso finalizado',
          'atividades finalizadas'
        ];

        const pageText = document.body.textContent.toLowerCase();
        let isCompleted = false;
        
        for (const indicator of completionIndicators) {
          if (pageText.includes(indicator)) {
            isCompleted = true;
            break;
          }
        }

        return {
          progress: progressElements[0] ? progressElements[0].textContent.trim() : '',
          isCompleted: isCompleted,
          pageText: pageText
        };
      });

      // Log do progresso se for diferente do último
      if (progressInfo.progress && progressInfo.progress !== lastProgress) {
        logCourse(userId, courseName, `Progresso: ${progressInfo.progress}`, 'cyan');
        lastProgress = progressInfo.progress;
        
        // Tenta extrair números do progresso (ex: "2/8")
        const progressMatch = progressInfo.progress.match(/(\d+)\s*\/\s*(\d+)/);
        if (progressMatch) {
          completedCount = parseInt(progressMatch[1]);
          activityCount = parseInt(progressMatch[2]);
          logCourse(userId, courseName, `Atividades: ${completedCount}/${activityCount}`, 'cyan');
        }
      }

      // Verifica se está realmente completo
      if (progressInfo.isCompleted) {
        consecutiveCompletions++;
        logCourse(userId, courseName, `Conclusão detectada (${consecutiveCompletions}/3)`, 'green');
        
        // Espera por 3 verificações consecutivas para confirmar
        if (consecutiveCompletions >= 3) {
          logSuccess(userId, `Todas as atividades de ${courseName} concluídas!`);
          return true;
        }
      } else {
        consecutiveCompletions = 0; // Reseta se não estiver completo
      }

      // Verifica se há erro
      const hasError = await mainPage.evaluate(() => {
        const errorSelectors = [
          '[class*="error"]',
          '[class*="fail"]',
          '[class*="alert"]',
          '[class*="danger"]',
          '.error-message',
          '.alert-danger'
        ];
        
        for (const selector of errorSelectors) {
          const elements = document.querySelectorAll(selector);
          for (const el of elements) {
            const text = el.textContent.toLowerCase();
            if (text.includes('erro') || text.includes('error') || text.includes('falha')) {
              return true;
            }
          }
        }
        return false;
      });

      if (hasError) {
        logError(userId, `Erro detectado no processamento de ${courseName}`);
        return false;
      }

      // Verifica se a página mudou (curso foi redirecionado ou finalizado)
      const currentUrl = await mainPage.url();
      if (!currentUrl.includes(CFG.TARGET_URL)) {
        logCourse(userId, courseName, `Redirecionamento detectado, assumindo conclusão`, 'green');
        return true;
      }

      // Aguarda antes de verificar novamente
      await sleep(10000); // Verifica a cada 10 segundos

    } catch (error) {
      logWarning(userId, `Erro ao verificar progresso: ${error.message}`);
      await sleep(5000);
      
      // Se houver muitos erros, assume que terminou
      if (Date.now() - startTime > 600000) { // 10 minutos
        logWarning(userId, `Muitos erros, assumindo conclusão de ${courseName}`);
        return true;
      }
    }
  }

  logWarning(userId, `Tempo máximo excedido para ${courseName}. Continuando...`);
  return false;
}

async function processCourse(mainPage, course, raVal, senhaVal, userId) {
  let success = false;
  let attempts = 0;
  
  while (!success && attempts < CFG.MAX_RETRIES) {
    attempts++;
    try {
      logCourse(userId, course.name, `Tentativa ${attempts}/${CFG.MAX_RETRIES}`, 'cyan');
      
      const modalOpened = await reopenCoursesModal(mainPage, userId);
      if (!modalOpened) {
        logWarning(userId, `Não foi possível abrir o modal para: ${course.name}`);
        continue;
      }

      const courseElement = await mainPage.$(`.task-list-item[data-course-id="${course.id}"]`);
      if (courseElement) {
        logCourse(userId, course.name, "Clicando no curso...", 'cyan');
        await courseElement.click();
        await sleep(5000); // Aumentado para 5 segundos

        // Verifica qual modal apareceu
        const modalResult = await mainPage.evaluate(() => {
          // Primeiro verifica modal de quiz
          const quizModal = document.querySelector('#quizConfirmationModal');
          if (quizModal && getComputedStyle(quizModal).display !== 'none') {
            const processBtn = document.querySelector('#processQuizDefaultTimeBtn');
            if (processBtn) {
              processBtn.click();
              return 'quiz_processing';
            }
          }
          
          // Depois verifica modal normal
          const normalModal = document.querySelector('#confirmationModal');
          if (normalModal && getComputedStyle(normalModal).display !== 'none') {
            const processBtn = document.querySelector('#processCourseBtn, button:contains("Processar"), button:contains("Confirmar")');
            if (processBtn) {
              processBtn.click();
              return 'course_processing';
            }
            return 'no_process_button';
          }
          
          // Verifica se já está processando
          const processingElements = document.querySelectorAll('[class*="processing"], [class*="progress"]');
          if (processingElements.length > 0) {
            return 'already_processing';
          }
          
          return 'no_modal_found';
        });

        if (modalResult === 'quiz_processing' || modalResult === 'course_processing' || modalResult === 'already_processing') {
          if (modalResult === 'quiz_processing') {
            logCourse(userId, course.name, "Questionário iniciado", 'green');
          } else if (modalResult === 'course_processing') {
            logCourse(userId, course.name, "Processamento do curso iniciado", 'green');
          } else {
            logCourse(userId, course.name, "Processamento já em andamento", 'green');
          }
          
          // ESPERA TODAS AS ATIVIDADES SEREM CONCLUÍDAS
          logCourse(userId, course.name, "Aguardando conclusão de TODAS as atividades...", 'yellow');
          const completed = await waitForActivitiesCompletion(mainPage, userId, course.name);
          
          if (completed) {
            logSuccess(userId, `Curso concluído: ${course.name}`);
            success = true;
            
            // Aguarda um tempo extra para garantir
            await sleep(5000);
          } else {
            logWarning(userId, `Curso pode não estar totalmente concluído: ${course.name}`);
            // Mesmo assim marca como sucesso para continuar
            success = true;
          }
          
        } else if (modalResult === 'no_process_button') {
          logWarning(userId, `Botão de processamento não encontrado para: ${course.name}`);
        } else {
          logWarning(userId, `Modal não encontrado para: ${course.name}`);
        }

      } else {
        logWarning(userId, `Curso não encontrado no modal: ${course.name}`);
      }

      // Fecha qualquer modal ou notificação
      await mainPage.evaluate(() => {
        const closeBtns = [
          "#closeCoursesModal",
          ".close-modal",
          ".notification-close",
          ".btn-secondary",
          'button:contains("Fechar")',
          'button:contains("Close")'
        ];
        
        for (const selector of closeBtns) {
          try {
            const btn = document.querySelector(selector);
            if (btn) btn.click();
          } catch (e) {}
        }
      });
      await sleep(2000);
      
    } catch (error) {
      logError(userId, `Erro na tentativa ${attempts}: ${error.message}`);
      await sleep(2000);
    }
  }
  
  if (!success) {
    logError(userId, `Não foi possível processar: ${course.name} após ${attempts} tentativas`);
  } else {
    logSuccess(userId, `Curso processado: ${course.name}`);
  }
  
  return success;
}

async function processUser(browser, row, keys, colRa, colDigito, colEstado, colSenha, index, total) {
  const mainPage = await browser.newPage();
  mainPage.setDefaultNavigationTimeout(CFG.NAVIGATION_TIMEOUT);
  
  let raVal = "";
  if (colRa) raVal = safeText(row[colRa]);
  if (colDigito && row[colDigito] !== "") raVal += safeText(row[colDigito]);
  if (colEstado && row[colEstado] !== "") raVal += safeText(row[colEstado]);
  if (!raVal) {
    const alt = detectColumn(keys, ["id", "user", "login"]);
    if (alt) raVal = safeText(row[alt]);
  }
  let senhaVal = "";
  if (colSenha) senhaVal = safeText(row[colSenha]);
  if (!senhaVal) {
    const altPwd = detectColumn(keys, ["senha", "password", "pass", "pwd"]);
    if (altPwd) senhaVal = safeText(row[altPwd]);
  }

  const userId = raVal.substring(0, 8) + '...'; // Mostra apenas os primeiros caracteres por segurança
  const ts = new Date().toISOString();
  appendLog(`${ts},${raVal},${senhaVal},start,linha_${index + 1}`);
  
  console.log(`\n${colors.bright}${colors.magenta}=== PROCESSANDO USUÁRIO ${index + 1}/${total} ===${colors.reset}`);
  logUser(userId, `Iniciando processamento`, 'magenta');

  try {
    await mainPage.goto(CFG.TARGET_URL, { waitUntil: "domcontentloaded" });
    await mainPage.waitForSelector("#ra", { timeout: 8000 });
    await mainPage.waitForSelector("#senha", { timeout: 8000 });

    await mainPage.evaluate(
      (r, s) => {
        const raInput = document.querySelector("#ra");
        const senhaInput = document.querySelector("#senha");
        if (raInput) {
          raInput.focus();
          raInput.value = "";
          raInput.value = r;
          raInput.dispatchEvent(new Event("input", { bubbles: true }));
        }
        if (senhaInput) {
          senhaInput.focus();
          senhaInput.value = "";
          senhaInput.value = s;
          senhaInput.dispatchEvent(new Event("input", { bubbles: true }));
        }
      },
      raVal,
      senhaVal
    );

    appendLog(`${new Date().toISOString()},${raVal},${senhaVal},filled,ok`);
    logUser(userId, "Campos preenchidos", 'green');

    if (CFG.CLICK_SUBMIT) {
      try {
        await mainPage.evaluate(() => {
          const btn = document.querySelector("#buscarCursosBtn");
          if (btn) btn.click();
        });
        appendLog(
          `${new Date().toISOString()},${raVal},${senhaVal},clicked,ok`
        );
        logUser(userId, "Botão de buscar cursos clicado", 'green');

        logUser(userId, "Aguardando modal carregar...", 'cyan');
        await mainPage.waitForSelector("#coursesModal", {
          visible: true,
          timeout: 30000,
        });
        
        await sleep(3000);
        
        logUser(userId, "Modal carregado. Obtendo cursos...", 'green');
        
        const coursesData = await mainPage.evaluate(() => {
          const courses = [];
          const courseElements = document.querySelectorAll('#coursesList .task-list-item');
          
          courseElements.forEach((el, index) => {
            const courseId = el.getAttribute('data-course-id');
            const courseName = el.getAttribute('data-course-name') || 
                             el.querySelector('label')?.textContent?.split('(')[0]?.trim() || 
                             `Curso ${index + 1}`;
            courses.push({ id: courseId, name: courseName, index });
          });
          
          return courses;
        });
        
        logUser(userId, `Encontrados ${coursesData.length} cursos`, 'green');
        coursesData.forEach((course, i) => {
          logUser(userId, `${i + 1}. ${course.name}`, 'white');
        });
        
        await mainPage.evaluate(() => {
          const closeBtn = document.querySelector("#closeCoursesModal");
          if (closeBtn) closeBtn.click();
        });
        await sleep(1000);
        
        // PROCESSAMENTO SEQUENCIAL DOS CURSOS
        logUser(userId, "Iniciando processamento dos cursos...", 'magenta');
        for (const course of coursesData) {
          await processCourse(mainPage, course, raVal, senhaVal, userId);
        }
        
        logSuccess(userId, "Todos os cursos processados!");
        
      } catch (errClick) {
        appendLog(
          `${new Date().toISOString()},${raVal},${senhaVal},clicked,error:${
            errClick.message
          }`
        );
        logError(userId, `Falha ao clicar: ${errClick.message}`);
      }
    }

    const delay = randBetween(CFG.MIN_DELAY_MS, CFG.MAX_DELAY_MS);
    logUser(userId, `Aguardando ${delay}ms antes do próximo usuário...`, 'yellow');
    await sleep(delay);
    appendLog(`${new Date().toISOString()},${raVal},${senhaVal},done,ok`);
    
  } catch (err) {
    console.error(`Erro na linha ${index + 1}:`, err.message || err);
    appendLog(
      `${new Date().toISOString()},${raVal},${senhaVal},error,${(
        err.message || ""
      ).replace(/[\r\n,]/g, " ")}`
    );
    logError(userId, `Erro geral: ${err.message}`);
  } finally {
    await mainPage.close();
    logUser(userId, "Sessão finalizada", 'blue');
  }
}

async function main() {
  console.log(`${colors.bright}${colors.magenta}=== INICIANDO PROCESSAMENTO ===${colors.reset}`);
  
  if (!fs.existsSync(CFG.EXCEL_FILE)) {
    console.error(`Arquivo não encontrado: ${CFG.EXCEL_FILE}`);
    process.exit(1);
  }
  const wb = XLSX.readFile(CFG.EXCEL_FILE);
  const sheetName = CFG.SHEET_NAME || wb.SheetNames[0];
  const sheet = wb.Sheets[sheetName];
  const rows = XLSX.utils.sheet_to_json(sheet, { defval: "" });
  if (!rows || rows.length === 0) {
    console.error("Planilha vazia.");
    process.exit(1);
  }

  if (!fs.existsSync(CFG.LOG_FILE))
    fs.writeFileSync(CFG.LOG_FILE, "timestamp,ra,senha,action,details\n");

  const keys = Object.keys(rows[0]);
  const colRa = detectColumn(keys, ["ra", "registro", "matricula"]);
  const colDigito = detectColumn(keys, ["digito", "dv"]);
  const colEstado = detectColumn(keys, ["estado", "uf"]);
  const colSenha = detectColumn(keys, ["senha", "password", "pass", "pwd"]);
  
  console.log(`${colors.cyan}Colunas detectadas:${colors.reset}`);
  console.log(`${colors.cyan}- RA: ${colRa || 'Não encontrado'}${colors.reset}`);
  console.log(`${colors.cyan}- Dígito: ${colDigito || 'Não encontrado'}${colors.reset}`);
  console.log(`${colors.cyan}- Estado: ${colEstado || 'Não encontrado'}${colors.reset}`);
  console.log(`${colors.cyan}- Senha: ${colSenha || 'Não encontrado'}${colors.reset}`);
  console.log(`${colors.cyan}Total de usuários: ${rows.length}${colors.reset}\n`);

  const browser = await puppeteer.launch({
    headless: CFG.HEADLESS,
    executablePath: "C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe",
    defaultViewport: { width: 1200, height: 800 },
    args: ["--no-sandbox", "--disable-setuid-sandbox"],
  });

  // PROCESSAMENTO PARALELO DE USUÁRIOS
  const userPromises = rows.map((row, index) => 
    processUser(browser, row, keys, colRa, colDigito, colEstado, colSenha, index, rows.length)
  );
  
  await Promise.all(userPromises);

  await browser.close();
  console.log(`\n${colors.bright}${colors.magenta}=== PROCESSAMENTO CONCLUÍDO ===${colors.reset}`);
  console.log(`${colors.green}Log salvo em: ${CFG.LOG_FILE}${colors.reset}`);
}

main().catch((err) => {
  console.error(`${colors.red}Erro fatal: ${err}${colors.reset}`);
});