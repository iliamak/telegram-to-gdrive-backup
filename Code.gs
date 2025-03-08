/**
 * Скрипт для бэкапа файлов из Telegram бота на Google Drive
 * 
 * Инструкция по настройке:
 * 1. Создайте бота в Telegram через BotFather и получите токен
 * 2. Вставьте токен в константу TELEGRAM_BOT_TOKEN ниже
 * 3. Запустите функцию createFolders() один раз для создания структуры папок
 * 4. Перезагрузите таблицу, чтобы увидеть новое меню
 * 5. Пересылайте сообщения из избранного вашему боту
 * 6. Нажмите "Сделать бэкап" в меню
 */

// Константы для настройки
const TELEGRAM_BOT_TOKEN = 'YOUR_TELEGRAM_BOT_TOKEN'; // Замените на токен вашего бота
const ROOT_FOLDER_NAME = 'Telegram Backup'; // Название корневой папки на Google Drive
const MAX_FILES_TO_PROCESS = 30; // Максимальное количество файлов для обработки

// Создаем меню при открытии таблицы
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Telegram Backup')
    .addItem('Сделать бэкап', 'startBackup')
    .addItem('Настроить папки', 'createFolders')
    .addToUi();
}

/**
 * Создает структуру папок на Google Drive, если она не существует
 */
function createFolders() {
  try {
    const rootFolder = findOrCreateFolder(DriveApp.getRootFolder(), ROOT_FOLDER_NAME);
    findOrCreateFolder(rootFolder, 'Фото');
    findOrCreateFolder(rootFolder, 'Видео');
    findOrCreateFolder(rootFolder, 'Документы');
    
    SpreadsheetApp.getUi().alert('Папки успешно созданы на Google Drive');
    
    // Инициализируем лист логов, если он не существует
    initLogSheet();
    
    return true;
  } catch (error) {
    Logger.log('Ошибка создания папок: ' + error.message);
    SpreadsheetApp.getUi().alert('Ошибка: ' + error.message);
    return false;
  }
}

/**
 * Находит папку по имени или создает новую, если она не существует
 */
function findOrCreateFolder(parentFolder, folderName) {
  const folders = parentFolder.getFoldersByName(folderName);
  
  if (folders.hasNext()) {
    return folders.next();
  } else {
    return parentFolder.createFolder(folderName);
  }
}

/**
 * Инициализирует лист логов в таблице
 */
function initLogSheet() {
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let logSheet = ss.getSheetByName('Backup Logs');
  
  if (!logSheet) {
    logSheet = ss.insertSheet('Backup Logs');
    
    // Настраиваем заголовки
    logSheet.appendRow([
      'Дата', 'Время', 'Файл', 'Тип', 'Размер (KB)', 'Статус'
    ]);
    
    // Форматируем заголовки
    logSheet.getRange(1, 1, 1, 6).setFontWeight('bold').setBackground('#f3f3f3');
    logSheet.setFrozenRows(1);
    
    // Настраиваем ширину столбцов
    logSheet.setColumnWidth(1, 100); // Дата
    logSheet.setColumnWidth(2, 80);  // Время
    logSheet.setColumnWidth(3, 300); // Файл
    logSheet.setColumnWidth(4, 80);  // Тип
    logSheet.setColumnWidth(5, 100); // Размер
    logSheet.setColumnWidth(6, 100); // Статус
  }
  
  return logSheet;
}

/**
 * Получает последние обновления от бота
 */
function getUpdatesFromBot() {
  const url = `https://api.telegram.org/bot${TELEGRAM_BOT_TOKEN}/getUpdates`;
  
  try {
    const response = UrlFetchApp.fetch(url);
    const data = JSON.parse(response.getContentText());
    
    if (!data.ok) {
      throw new Error('Telegram API вернул ошибку: ' + JSON.stringify(data));
    }
    
    return data.result;
  } catch (error) {
    Logger.log('Ошибка получения обновлений: ' + error.message);
    throw error;
  }
}

/**
 * Получает информацию о файле по его ID
 */
function getFileInfo(fileId) {
  const url = `https://api.telegram.org/bot${TELEGRAM_BOT_TOKEN}/getFile?file_id=${fileId}`;
  
  try {
    const response = UrlFetchApp.fetch(url);
    const data = JSON.parse(response.getContentText());
    
    if (!data.ok) {
      throw new Error('Не удалось получить информацию о файле: ' + JSON.stringify(data));
    }
    
    return data.result;
  } catch (error) {
    Logger.log('Ошибка получения информации о файле: ' + error.message);
    throw error;
  }
}

/**
 * Скачивает файл из Telegram и сохраняет его на Google Drive
 */
function downloadAndSaveFile(fileInfo, fileName, fileType) {
  // Получаем URL для скачивания файла
  const filePath = fileInfo.file_path;
  const downloadUrl = `https://api.telegram.org/file/bot${TELEGRAM_BOT_TOKEN}/${filePath}`;
  
  try {
    // Скачиваем файл
    const fileBlob = UrlFetchApp.fetch(downloadUrl).getBlob();
    fileBlob.setName(fileName);
    
    // Находим корневую папку для бэкапа
    const rootFolders = DriveApp.getRootFolder().getFoldersByName(ROOT_FOLDER_NAME);
    if (!rootFolders.hasNext()) {
      throw new Error('Корневая папка не найдена. Запустите функцию createFolders()');
    }
    const rootFolder = rootFolders.next();
    
    // Определяем в какую папку сохранять файл
    let targetFolderName;
    switch (fileType) {
      case 'photo':
        targetFolderName = 'Фото';
        break;
      case 'video':
        targetFolderName = 'Видео';
        break;
      default:
        targetFolderName = 'Документы';
        break;
    }
    
    // Находим целевую папку
    const targetFolders = rootFolder.getFoldersByName(targetFolderName);
    if (!targetFolders.hasNext()) {
      throw new Error(`Папка ${targetFolderName} не найдена. Запустите функцию createFolders()`);
    }
    const targetFolder = targetFolders.next();
    
    // Сохраняем файл
    const savedFile = targetFolder.createFile(fileBlob);
    
    return {
      fileName: fileName,
      fileType: fileType,
      fileSize: Math.round(fileInfo.file_size / 1024), // Размер в KB
      status: 'Сохранено',
      driveUrl: savedFile.getUrl()
    };
  } catch (error) {
    Logger.log('Ошибка скачивания и сохранения файла: ' + error.message);
    return {
      fileName: fileName,
      fileType: fileType,
      fileSize: fileInfo.file_size ? Math.round(fileInfo.file_size / 1024) : 0,
      status: 'Ошибка: ' + error.message,
      driveUrl: null
    };
  }
}

/**
 * Извлекает информацию о медиафайле из сообщения
 */
function extractFileData(message) {
  if (!message) return null;
  
  // Проверяем наличие фото
  if (message.photo && message.photo.length > 0) {
    // Берем фото с наилучшим качеством (последнее в массиве)
    const photo = message.photo[message.photo.length - 1];
    return {
      fileId: photo.file_id,
      fileName: `photo_${message.message_id}.jpg`,
      fileType: 'photo'
    };
  }
  
  // Проверяем наличие документа
  if (message.document) {
    return {
      fileId: message.document.file_id,
      fileName: message.document.file_name || `document_${message.message_id}.${message.document.mime_type.split('/')[1] || 'file'}`,
      fileType: 'document'
    };
  }
  
  // Проверяем наличие видео
  if (message.video) {
    return {
      fileId: message.video.file_id,
      fileName: `video_${message.message_id}.mp4`,
      fileType: 'video'
    };
  }
  
  // Проверяем наличие аудио
  if (message.audio) {
    return {
      fileId: message.audio.file_id,
      fileName: message.audio.title ? `${message.audio.title}.mp3` : `audio_${message.message_id}.mp3`,
      fileType: 'document'
    };
  }
  
  // Проверяем наличие голосового сообщения
  if (message.voice) {
    return {
      fileId: message.voice.file_id,
      fileName: `voice_${message.message_id}.ogg`,
      fileType: 'document'
    };
  }
  
  // Проверяем наличие видеосообщения
  if (message.video_note) {
    return {
      fileId: message.video_note.file_id,
      fileName: `video_note_${message.message_id}.mp4`,
      fileType: 'video'
    };
  }
  
  // Если нет медиа, возвращаем null
  return null;
}

/**
 * Добавляет запись в лог
 */
function logBackupResult(result) {
  const now = new Date();
  const dateStr = Utilities.formatDate(now, Session.getScriptTimeZone(), 'dd.MM.yyyy');
  const timeStr = Utilities.formatDate(now, Session.getScriptTimeZone(), 'HH:mm:ss');
  
  const logSheet = initLogSheet();
  
  logSheet.appendRow([
    dateStr,
    timeStr,
    result.fileName,
    result.fileType,
    result.fileSize,
    result.status
  ]);
}

/**
 * Основная функция для запуска процесса бэкапа
 */
function startBackup() {
  try {
    // Проверка настроек и структуры папок
    if (!checkSetup()) {
      return;
    }
    
    // Получаем обновления от бота
    const updates = getUpdatesFromBot();
    
    // Если нет обновлений
    if (!updates || updates.length === 0) {
      showStatusMessage('Нет новых сообщений у бота. Перешлите избранные сообщения боту и попробуйте снова.');
      return;
    }
    
    // Извлекаем сообщения с файлами
    const fileMessages = [];
    for (let i = 0; i < updates.length; i++) {
      const update = updates[i];
      
      if (update.message && (
          update.message.photo ||
          update.message.document ||
          update.message.video ||
          update.message.audio ||
          update.message.voice ||
          update.message.video_note
        )) {
        fileMessages.push(update.message);
      }
      
      // Проверяем, не превышаем ли лимит файлов
      if (fileMessages.length >= MAX_FILES_TO_PROCESS) {
        break;
      }
    }
    
    // Если нет сообщений с файлами
    if (fileMessages.length === 0) {
      showStatusMessage('Не найдено сообщений с файлами. Перешлите избранные сообщения с медиафайлами боту и попробуйте снова.');
      return;
    }
    
    // Показываем прогресс
    showProgressAlert(fileMessages.length);
    
    // Обрабатываем каждый файл
    let successCount = 0;
    let errorCount = 0;
    
    for (let i = 0; i < fileMessages.length; i++) {
      try {
        const fileData = extractFileData(fileMessages[i]);
        
        if (fileData) {
          // Получаем информацию о файле
          const fileInfo = getFileInfo(fileData.fileId);
          
          // Скачиваем и сохраняем файл
          const result = downloadAndSaveFile(fileInfo, fileData.fileName, fileData.fileType);
          
          // Логируем результат
          logBackupResult(result);
          
          if (result.status.startsWith('Ошибка')) {
            errorCount++;
          } else {
            successCount++;
          }
        }
      } catch (error) {
        errorCount++;
        Logger.log('Ошибка обработки файла: ' + error.message);
        logBackupResult({
          fileName: `Неизвестный файл ${i}`,
          fileType: 'unknown',
          fileSize: 0,
          status: 'Ошибка: ' + error.message
        });
      }
    }
    
    // Показываем результаты
    const message = `Бэкап завершен!\n\nФайлов обработано: ${fileMessages.length}\nУспешно сохранено: ${successCount}\nОшибок: ${errorCount}\n\nРезультаты вы можете увидеть на листе "Backup Logs"`;
    SpreadsheetApp.getUi().alert(message);
    
  } catch (error) {
    Logger.log('Общая ошибка бэкапа: ' + error.message);
    SpreadsheetApp.getUi().alert('Произошла ошибка: ' + error.message);
  }
}

/**
 * Проверяет настройку скрипта и структуру папок
 */
function checkSetup() {
  // Проверяем токен бота
  if (TELEGRAM_BOT_TOKEN === 'YOUR_TELEGRAM_BOT_TOKEN') {
    SpreadsheetApp.getUi().alert('Пожалуйста, установите токен вашего Telegram бота в скрипте (TELEGRAM_BOT_TOKEN).');
    return false;
  }
  
  // Проверяем наличие корневой папки
  const rootFolders = DriveApp.getRootFolder().getFoldersByName(ROOT_FOLDER_NAME);
  if (!rootFolders.hasNext()) {
    const response = SpreadsheetApp.getUi().alert(
      'Структура папок не найдена. Создать папки сейчас?',
      SpreadsheetApp.getUi().ButtonSet.YES_NO
    );
    
    if (response === SpreadsheetApp.getUi().Button.YES) {
      return createFolders();
    } else {
      return false;
    }
  }
  
  return true;
}

/**
 * Показывает сообщение о статусе операции
 */
function showStatusMessage(message) {
  SpreadsheetApp.getUi().alert(message);
}

/**
 * Показывает информацию о прогрессе
 */
function showProgressAlert(fileCount) {
  SpreadsheetApp.getUi().alert(
    `Начинаем бэкап ${fileCount} файлов.\n\nЭто может занять некоторое время. Пожалуйста, дождитесь завершения операции.`
  );
}
