const { app, BrowserWindow, ipcMain, Tray, Menu } = require("electron");
const XLSX = require("xlsx");
const path = require("path");
const nodemailer = require("nodemailer");
const cron = require("node-cron");
const fs = require("fs");

let splashWindow, mainWindow, tray;
const EMAIL_TEMPLATE_PATH = path.join(__dirname, "email_template.html");

const EMAIL_CONFIG = {
  host: "smtp.gmail.com",
  port: 465,
  secure: true,
  user: "jimsfavurboy@gmail.com", // Ganti sesuai email-mu!
  pass: "yosfoohlhxqqwtyk", // Ganti sesuai password/mta-mu!
  subject: "Selamat Ulang Tahun 🎉 dari HRD",
};

const DEFAULT_EMAIL_TEMPLATE = `
<div style="font-size:1.1rem;">
  Hai <b>{{NAMA}}</b>,<br><br>
  Selamat ulang tahun yang ke-<b>{{UMUR}}</b>! 🎉<br>
  Semoga sehat dan sukses selalu.<br><br>
  Salam,<br>HRD
</div>
`.trim();

function excelDateToJSDate(serial) {
  if (!serial) return "";
  const utc_days = Math.floor(serial - 25569);
  const utc_value = utc_days * 86400;
  const date_info = new Date(utc_value * 1000);
  const day = String(date_info.getUTCDate()).padStart(2, "0");
  const month = String(date_info.getUTCMonth() + 1).padStart(2, "0");
  const year = date_info.getUTCFullYear();
  return `${day}-${month}-${year}`;
}

function getTodayBirthday(data) {
  let today = new Date();
  let dd = String(today.getDate()).padStart(2, "0");
  let mm = String(today.getMonth() + 1).padStart(2, "0");
  function getAgeFromTanggalLahir(tglLahir) {
    if (!tglLahir) return "";
    const [d, m, y] = tglLahir.split("-");
    const birthDate = new Date(`${y}-${m}-${d}`);
    let age = today.getFullYear() - birthDate.getFullYear();
    const mDiff = today.getMonth() - birthDate.getMonth();
    if (mDiff < 0 || (mDiff === 0 && today.getDate() < birthDate.getDate()))
      age--;
    return age;
  }
  return data
    .map((row) => {
      let tgl = row["Tanggal Lahir"];
      if (typeof tgl === "number") tgl = excelDateToJSDate(tgl);
      let [d, m] = (tgl || "").split("-");
      const isUltah = d === dd && m === mm;
      const umur = getAgeFromTanggalLahir(tgl);
      return { ...row, "Tanggal Lahir": tgl, umur, isUltah };
    })
    .filter((row) => row.isUltah);
}

// =========================
// HANDLER AREA (NO USER ADMIN!)
// =========================

ipcMain.handle("load-excel-db", (event) => {
  try {
    const EXCEL_PATH = path.join(__dirname, "data_karyawan.xlsx");
    const workbook = XLSX.readFile(EXCEL_PATH);
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    const data = XLSX.utils.sheet_to_json(worksheet);
    return data;
  } catch (e) {
    console.error("[Excel Load Error]", e);
    return [];
  }
});
ipcMain.handle("load-email-template", (event) => {
  try {
    if (fs.existsSync(EMAIL_TEMPLATE_PATH)) {
      return fs.readFileSync(EMAIL_TEMPLATE_PATH, "utf8");
    }
    return DEFAULT_EMAIL_TEMPLATE;
  } catch (e) {
    console.error("[Load Template Error]", e);
    return DEFAULT_EMAIL_TEMPLATE;
  }
});
ipcMain.handle("save-email-template", (event, template) => {
  try {
    fs.writeFileSync(EMAIL_TEMPLATE_PATH, template, "utf8");
    return { success: true };
  } catch (e) {
    console.error("[Save Template Error]", e);
    return { success: false, error: e.message };
  }
});
ipcMain.handle(
  "broadcast-birthday",
  async (event, data, emailConfig, emailTemplate) => {
    const resultList = [];
    try {
      let transporter = nodemailer.createTransport({
        host: emailConfig.host,
        port: emailConfig.port,
        secure: emailConfig.secure,
        auth: {
          user: emailConfig.user,
          pass: emailConfig.pass,
        },
      });

      for (const row of data) {
        try {
          let htmlBody = emailTemplate
            .replace(/\{\{NAMA\}\}/g, row.Nama || "rekan")
            .replace(/\{\{TANGGAL\}\}/g, row["Tanggal Lahir"] || "")
            .replace(/\{\{UMUR\}\}/g, row.umur || "");

          await transporter.sendMail({
            from: `"HRD" <${emailConfig.user}>`,
            to: row.Email,
            subject: emailConfig.subject || "Selamat Ulang Tahun 🎉",
            html: htmlBody,
          });

          resultList.push({ email: row.Email, status: "success" });
          await new Promise((res) => setTimeout(res, 700));
        } catch (errSend) {
          console.error("[SendMail Error]", row.Email, errSend);
          resultList.push({
            email: row.Email,
            status: "fail",
            error: errSend.message,
          });
        }
      }
      return { success: true, results: resultList };
    } catch (e) {
      console.error("[Broadcast Handler Error]", e);
      return { success: false, error: e.message, results: resultList };
    }
  }
);

// ================
// WINDOW FUNCTION
// ================
function createSplashWindow() {
  splashWindow = new BrowserWindow({
    width: 400,
    height: 340,
    frame: false,
    transparent: true,
    alwaysOnTop: true,
    resizable: true,
    center: true,
    webPreferences: { nodeIntegration: true, contextIsolation: false },
  });
  splashWindow.loadFile("splash.html");
  splashWindow.on("closed", () => (splashWindow = null));
}

// Main window (langsung index.html)
function createMainWindow() {
  mainWindow = new BrowserWindow({
    width: 900,
    height: 650,
    show: false,
    webPreferences: { nodeIntegration: true, contextIsolation: false },
  });
  mainWindow.loadFile("index.html");
  mainWindow.on("closed", () => (mainWindow = null));
  mainWindow.once("ready-to-show", () => mainWindow.show());
}

// URUTAN SPLASH → MAIN LANGSUNG
ipcMain.on("splash-done", () => {
  if (splashWindow) splashWindow.close();
  createMainWindow();
});

// TRAY SECTION
function createTray() {
  tray = new Tray(path.join(__dirname, "assets/logo.png"));
  const contextMenu = Menu.buildFromTemplate([
    {
      label: "Show App",
      click: () => {
        if (mainWindow) mainWindow.show();
      },
    },
    {
      label: "Quit",
      click: () => {
        app.isQuiting = true;
        app.quit();
      },
    },
  ]);
  tray.setToolTip("Birthday Broadcast App");
  tray.setContextMenu(contextMenu);
  tray.on("double-click", () => {
    if (mainWindow) mainWindow.show();
  });
}

// CRON AUTO MAIL
function autoBroadcastBirthday() {
  try {
    const EXCEL_PATH = path.join(__dirname, "data_karyawan.xlsx");
    const workbook = XLSX.readFile(EXCEL_PATH);
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    const data = XLSX.utils.sheet_to_json(worksheet);
    const ultahToday = getTodayBirthday(data);

    // Baca template dari file
    let emailTemplate = DEFAULT_EMAIL_TEMPLATE;
    if (fs.existsSync(EMAIL_TEMPLATE_PATH)) {
      try {
        emailTemplate = fs.readFileSync(EMAIL_TEMPLATE_PATH, "utf8");
      } catch (e) {
        console.error("[Read Email Template (auto) Error]", e);
      }
    }

    if (ultahToday.length > 0) {
      let transporter = nodemailer.createTransport({
        host: EMAIL_CONFIG.host,
        port: EMAIL_CONFIG.port,
        secure: EMAIL_CONFIG.secure,
        auth: { user: EMAIL_CONFIG.user, pass: EMAIL_CONFIG.pass },
      });
      ultahToday.forEach(async (row) => {
        let htmlBody = emailTemplate
          .replace(/\{\{NAMA\}\}/g, row.Nama || "rekan")
          .replace(/\{\{TANGGAL\}\}/g, row["Tanggal Lahir"] || "")
          .replace(/\{\{UMUR\}\}/g, row.umur || "");
        try {
          await transporter.sendMail({
            from: `"HRD" <${EMAIL_CONFIG.user}>`,
            to: row.Email,
            subject: EMAIL_CONFIG.subject,
            html: htmlBody,
          });
        } catch (err) {
          console.error("[Auto Broadcast Error]", row.Email, err);
        }
      });
      tray?.displayBalloon?.({
        title: "Birthday Broadcast",
        content: `Auto broadcast sukses ke ${ultahToday.length} orang!`,
      });
      console.log(
        `[AUTO BROADCAST] Sukses kirim ke ${ultahToday.length} orang`
      );
    } else {
      console.log("[AUTO BROADCAST] Tidak ada yang ultah hari ini");
    }
  } catch (e) {
    console.error("[AUTO BROADCAST ERROR]", e);
  }
}

// APP READY
app.whenReady().then(() => {
  createSplashWindow();
  createTray();
});
// CRON: jam 07:00 setiap hari
cron.schedule("0 7 * * *", () => {
  autoBroadcastBirthday();
});

app.on("window-all-closed", () => {
  if (process.platform !== "darwin") app.quit();
});
app.on("activate", () => {
  if (!mainWindow) createMainWindow();
});
