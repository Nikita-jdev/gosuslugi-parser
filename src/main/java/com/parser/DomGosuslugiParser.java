package com.parser;

import lombok.Setter;
import org.apache.poi.common.usermodel.HyperlinkType;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Hyperlink;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.NoSuchElementException;
import org.openqa.selenium.TimeoutException;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.Select;
import org.openqa.selenium.support.ui.WebDriverWait;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.time.Duration;
import java.time.LocalDate;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.concurrent.CompletableFuture;
import java.util.concurrent.ExecutionException;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;
import java.util.concurrent.TimeUnit;
import java.util.concurrent.atomic.AtomicBoolean;

import static java.lang.Thread.sleep;


public class DomGosuslugiParser {
    private WebDriver driver;
    private WebDriverWait wait;
    private final List<Company> companies = new ArrayList<>();

    private static final String TARGET_URL = "https://dom.gosuslugi.ru/#!/organizations?orgType=1&orgType=19&orgType=22&orgType=21&orgType=20&doSearch=false&restore=false";
    private static final int TIMEOUT_SECONDS = 30;
    private static final String BASE_URL = "https://dom.gosuslugi.ru";

    private int startPage = 1;
    private int currentPage = 1;

    @Setter
    private ProgressListener listener;
    @Setter
    private String region = "Санкт-Петербург";
    private AtomicBoolean cancelRequested = new AtomicBoolean(false);

    public void setCancellationFlag(AtomicBoolean cancelRequested) {
        this.cancelRequested = (cancelRequested != null) ? cancelRequested : new AtomicBoolean(false);
    }

    private void notifyStatus(String text) {
        if (listener != null) listener.onStatus(text);
        System.out.println(text);
    }

    private void notifyPageProgress(int current, int total) {
        if (listener != null) listener.onPageProgress(current, total);
    }

    private void notifyLog(String text) {
        if (listener != null) listener.log(text);
        System.out.println(text);
    }

    private void notifyFinished(boolean success, String message) {
        if (listener != null) listener.onFinished(success, message);
    }

    private void checkCancelled() throws InterruptedException {
        if (cancelRequested != null && cancelRequested.get()) {
            throw new InterruptedException("Операция отменена пользователем");
        }
    }

    public void setStartPage(int startPage) {
        this.startPage = Math.max(1, startPage);
    }

    public void parseOrganizations() {
        try {
            checkSeleniumSetup();

            notifyStatus("Запуск драйвера...");
            initDriver();
            notifyLog("🚀 Запуск парсера управляющих компаний...");

            driver.get(TARGET_URL);
            wait.until(ExpectedConditions.presenceOfElementLocated(By.cssSelector("body")));
            sleep(5000);
            checkCancelled();

            selectRegionFilter();
//            selectSpbFilter();
            checkCancelled();

            clickSearchButton();
            sleep(1000);
            checkCancelled();

            selectItemsPerPage("100");

            parseAllPages(startPage);

            notifyLog("📊 Всего найдено компаний: " + companies.size());

            notifyStatus("Сохранение в Excel...");
            saveToExcel();

            notifyFinished(true, "Успешно: сохранено " + companies.size() + " записей");
        } catch (InterruptedException ie) {
            notifyLog("⏹️ " + ie.getMessage());
            notifyFinished(false, ie.getMessage());
        } catch (Exception e) {
            notifyLog("❌ Ошибка при парсинге: " + e.getMessage());
            notifyFinished(false, "Ошибка: " + e.getMessage());
        } finally {
            if (driver != null) {
                driver.quit();
                notifyLog("🔴 Браузер закрыт");
            }
        }
    }

    public void initDriver() {
        try {
            String chromeDriverPath = "chromedriver.exe";
            File chromeDriverFile = new File(chromeDriverPath);
            if (chromeDriverFile.exists()) {
                System.setProperty("webdriver.chrome.driver", chromeDriverPath);
                notifyLog("✅ ChromeDriver найден: " + chromeDriverFile.getAbsolutePath());
            } else {
                notifyLog("⚠️ ChromeDriver не найден по пути: " + chromeDriverFile.getAbsolutePath());
                notifyLog("📥 Поместите chromedriver.exe в ту же папку, где находится программа");
                throw new RuntimeException("ChromeDriver не найден. Путь: " + chromeDriverFile.getAbsolutePath());
            }
        } catch (Exception e) {
            notifyLog("❌ Ошибка настройки ChromeDriver: " + e.getMessage());
            throw new RuntimeException("Не удалось настроить ChromeDriver", e);
        }

        try {
            driver = new ChromeDriver(createChromeOptions());
            wait = new WebDriverWait(driver, Duration.ofSeconds(TIMEOUT_SECONDS));
            notifyLog("🚀 Драйвер успешно инициализирован");
        } catch (Exception e) {
            notifyLog("❌ Ошибка инициализации драйвера: " + e.getMessage());
            throw new RuntimeException("Не удалось запустить ChromeDriver", e);
        }
    }

    private ChromeOptions createChromeOptions() {
        ChromeOptions options = new ChromeOptions();
        options.addArguments("--headless=new");
        options.addArguments("--window-size=1024,768");
        options.addArguments("--disable-blink-features=AutomationControlled");
        options.addArguments("--user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36");
        options.addArguments("--disable-gpu");
        options.addArguments("--no-sandbox");
        options.addArguments("--disable-dev-shm-usage");
        options.addArguments("--remote-allow-origins=*");
        options.addArguments("--disable-extensions");
        options.addArguments("--disable-plugins");
        options.addArguments("--disable-images");
        options.addArguments("--disable-javascript");
        options.addArguments("--memory-pressure-off");
        return options;
    }

    private void checkSeleniumSetup() {
        String chromeDriverPath = "chromedriver.exe";
        File chromeDriverFile = new File(chromeDriverPath);

        if (!chromeDriverFile.exists()) {
            notifyLog("❌ ВНИМАНИЕ: ChromeDriver не найден!");
            notifyLog("📂 Требуемый путь: " + chromeDriverFile.getAbsolutePath());
            notifyLog("💡 Действие: Поместите chromedriver.exe в ту же папку, где находится программа");
        } else {
            notifyLog("✅ ChromeDriver доступен: " + chromeDriverFile.getAbsolutePath());
        }
    }

    private void selectRegionFilter() {
        try {
            List<WebElement> selects = driver.findElements(By.cssSelector("select"));
            if (!selects.isEmpty()) {
                Select dropdown = new Select(selects.get(0));

                List <String> regions = dropdown.getOptions().stream()
                        .map(WebElement::getText)
                        .toList();

                String selectedRegion = listener.showRegionSelectionDialog(regions);

                if (selectedRegion == null) {
                    throw new InterruptedException("Пользователь отменил выбор региона");
                }

                boolean regionFound = false;

                for (WebElement option : dropdown.getOptions()) {
                    if (option.getText().contains(selectedRegion)) {
                        dropdown.selectByVisibleText(option.getText());
                        regionFound = true;
                        notifyLog("✅ Выбран регион: " + selectedRegion);
                        break;
                    }
                }

                if (!regionFound) {
                    notifyLog("⚠️ Регион '" + selectedRegion + "' не найден в списке, используется первый доступный");
                    // Выбираем первый доступный регион
                    if (dropdown.getOptions().size() > 1) {
                        dropdown.selectByIndex(1); // пропускаем "Все регионы" если есть
                    }
                }
            }
            sleep(1000);
        } catch (Exception e) {
            notifyLog("❌ Ошибка выбора региона: " + e.getMessage());
        }
    }

//    private void selectSpbFilter() {
//        try {
//            List<WebElement> selects = driver.findElements(By.cssSelector("select"));
//            if (!selects.isEmpty()) {
//                Select dropdown = new Select(selects.get(0));
//                for (WebElement option : dropdown.getOptions()) {
//                    if (option.getText().contains("Санкт-Петербург")) {
//                        dropdown.selectByVisibleText(option.getText());
//                        break;
//                    }
//                }
//            }
//            sleep(1000);
//        } catch (Exception e) {
//            notifyLog("Ошибка выбора фильтра: " + e.getMessage());
//        }
//    }

    private void clickSearchButton() {
        try {
            // Поиск кнопки только по атрибутам
            WebElement button = driver.findElement(By.cssSelector("button[type='submit'][class*='btn-prime']"));

            // Простой клик без лишних проверок
            ((JavascriptExecutor) driver).executeScript("arguments[0].click();", button);
            sleep(1000);

        } catch (Exception e) {
            notifyLog("❌ Ошибка при нажатии кнопки 'Найти': " + e.getMessage());
        }
    }

    private void selectItemsPerPage(String countPerPage) {
        try {
            // Ждем появления элемента "Выводить по"
            wait.until(ExpectedConditions.presenceOfElementLocated(
                    By.cssSelector("select.output-by_select, .output-by select, select[ng-model*='itemsPerPage']")
            ));

            // Ищем селект по различным возможным селекторам
            WebElement selectElement = null;
            String[] selectors = {
                    "select.output-by_select",
                    ".output-by select",
                    "select[ng-model*='itemsPerPage']",
                    "select[ng-model*='pagination']",
                    "select#count",
                    "select[title*='Babojatts']"
            };

            for (String selector : selectors) {
                try {
                    List<WebElement> elements = driver.findElements(By.cssSelector(selector));
                    if (!elements.isEmpty()) {
                        selectElement = elements.get(0);
                        break;
                    }
                } catch (Exception e) {
                    continue;
                }
            }

            if (selectElement != null) {
                Select dropdown = new Select(selectElement);

                try {
                    dropdown.selectByValue(countPerPage);

                } catch (Exception e) {
                    notifyLog("Ошибка выбора количества элементов на странице: " + e.getMessage());
                }

                // Ждем обновления контента после выбора
                sleep(1500);
            } else {
                notifyLog("⚠️ Элемент 'Выводить по' не найден, используем стандартные настройки");
            }
        } catch (Exception e) {
            notifyLog("❌ Ошибка при выборе количества элементов: " + e.getMessage());
        }
    }

    private void parseAllPages(int startPage) throws InterruptedException {
        int totalPages = getTotalPages();
        notifyLog("Общее количество страниц: " + totalPages);

        currentPage = startPage;

        if (startPage > 1) {
            if (startPage > totalPages) {
                notifyLog("❌ Стартовая страница " + startPage + " превышает общее количество страниц " + totalPages);
                return;
            }
            notifyLog("⏩ Переход к странице " + startPage);
            goToPage(startPage);
        }

        try {
            while (true) {
                // Проверка отмены в начале каждой страницы
                if (cancelRequested.get()) {
                    throw new InterruptedException("Операция отменена пользователем");
                }

                notifyPageProgress(currentPage, totalPages);
                notifyLog("📄 Обработка страницы " + currentPage + " из " + totalPages);

                parseCurrentPage();

                // СОХРАНЕНИЕ ПОСЛЕ КАЖДОЙ СТРАНИЦЫ
                if (!companies.isEmpty()) {
                    notifyStatus("Сохранение данных страницы " + currentPage + "...");
                    saveIntermediateResults();
                    cleanupMemory();
                }

                // Проверка отмены перед переходом на следующую страницу
                if (cancelRequested.get()) {
                    throw new InterruptedException("Операция отменена пользователем");
                }

                if (!goToNextPage()) {
                    notifyLog("✅ Достигнута последняя страница");
                    break;
                }

                currentPage++;
            }
        } catch (InterruptedException ie) {
            // Сохраняем прогресс при прерывании
            if (!companies.isEmpty()) {
                notifyStatus("Сохранение данных перед остановкой...");
                saveIntermediateResults();
            }
            throw ie;
        } catch (Exception e) {
            notifyLog("Ошибка парсинга страниц: " + e.getMessage());
        }
    }

    private void goToPage(int pageNumber) {
        try {
            int choicePage = 1;
            int countingPage = pageNumber;

            while (countingPage > 2) {
                WebElement pageLink = driver.findElement(By.xpath("//a[text()='" + (choicePage + 2) + "']"));
                if (pageLink != null && pageLink.isEnabled()) {
                    ((JavascriptExecutor) driver).executeScript("arguments[0].scrollIntoView(true);", pageLink);
                    sleep(1000);
                    pageLink.click();

                    countingPage -= 2;
                    choicePage += 2;
                }
            }

            if (countingPage == 2) {
                WebElement pageLink = driver.findElement(By.xpath("//a[text()='" + (choicePage + 1) + "']"));
                if (pageLink != null && pageLink.isEnabled()) {
                    ((JavascriptExecutor) driver).executeScript("arguments[0].scrollIntoView(true);", pageLink);
                    sleep(1000);
                    pageLink.click();
                }
            }

            // Ждем загрузки новой страницы
            wait.until(ExpectedConditions.presenceOfAllElementsLocatedBy(
                    By.cssSelector("ef-poch-ro-row[ng-repeat='organization in organizations'] .register-card")));
            sleep(2000);

            notifyLog("➡️ Переход на страницу " + pageNumber);
        } catch (Exception e) {
            notifyLog("❌ Ошибка перехода на страницу " + pageNumber + ": " + e.getMessage());
        }
    }

    private int getTotalPages() {
        try {
            // Способ 1: Ищем элемент с текстом "из" и следующую ссылку
            List<WebElement> totalPagesElements = driver.findElements(By.xpath(
                    "//span[contains(text(), 'из')]/following-sibling::a[contains(@ng-click, 'lastPage')]"
            ));

            if (!totalPagesElements.isEmpty()) {
                int totalPagesText = Integer.parseInt(totalPagesElements.get(0).getText().trim());
                if (totalPagesText >= 0) {
                    return totalPagesText;
                }
            }

            notifyLog("⚠️ Не удалось определить общее количество страниц");
            return 1;

        } catch (Exception e) {
            notifyLog("❌ Ошибка при получении количества страниц: " + e.getMessage());
            return 1;
        }
    }

    // Добавляем метод для промежуточного сохранения
    private void saveIntermediateResults() {
        if (companies.isEmpty()) {
            return;
        }

        try {
            String fileName = "Управляющие компании СПб " + LocalDate.now().getYear() + ".xlsx";
            boolean fileExists = new File(fileName).exists();

            Workbook workbook;
            Sheet sheet;

            if (fileExists) {
                try (FileInputStream fis = new FileInputStream(fileName)) {
                    workbook = new XSSFWorkbook(fis);
                }
                sheet = workbook.getSheet("Компании");
                if (sheet == null) {
                    sheet = workbook.createSheet("Компании");
                    createHeaders(sheet, workbook);
                }
            } else {
                workbook = new XSSFWorkbook();
                sheet = workbook.createSheet("Компании");
                createHeaders(sheet, workbook);
            }

            CellStyle defaultStyle = createDefaultStyle(workbook);
            CellStyle linkStyle = createLinkStyle(workbook);
            CreationHelper createHelper = workbook.getCreationHelper();

            // Получаем существующие компании из файла
            Map<String, Integer> existingCompanies = new HashMap<>();
            if (fileExists && sheet.getPhysicalNumberOfRows() > 1) {
                for (int i = 1; i <= sheet.getLastRowNum(); i++) {
                    Row row = sheet.getRow(i);
                    if (row != null && row.getCell(0) != null) {
                        String companyName = row.getCell(0).getStringCellValue();
                        if (companyName != null && !companyName.trim().isEmpty()) {
                            existingCompanies.put(companyName.trim(), i);
                        }
                    }
                }
            }

            int newRowsCount = 0;
            int updatedRowsCount = 0;

            // Добавляем/обновляем только новые компании
            for (Company company : companies) {
                if (company.getName() == null || company.getName().trim().isEmpty()) {
                    continue;
                }

                String companyName = company.getName().trim();
                Integer existingRowIndex = existingCompanies.get(companyName);

                if (existingRowIndex != null) {
                    updateCompanyRow(sheet.getRow(existingRowIndex), company, defaultStyle, linkStyle, createHelper);
                    updatedRowsCount++;
                } else {
                    int newRowIndex = sheet.getLastRowNum() + 1;
                    Row row = sheet.createRow(newRowIndex);
                    createCompanyRow(row, company, defaultStyle, linkStyle, createHelper);
                    newRowsCount++;
                    existingCompanies.put(companyName, newRowIndex);
                }
            }

            // Авто-размер колонок
            for (int i = 0; i < 11; i++) {
                sheet.autoSizeColumn(i);
                sheet.setColumnWidth(i, sheet.getColumnWidth(i) + 512);
            }

            // Авто-фильтр
            sheet.setAutoFilter(new CellRangeAddress(0, sheet.getLastRowNum(), 0, 10));

            try (FileOutputStream fos = new FileOutputStream(fileName)) {
                workbook.write(fos);
            }

            workbook.close();

            notifyLog("💾 Промежуточное сохранение: " + newRowsCount + " новых, " + updatedRowsCount + " обновлено");

        } catch (IOException e) {
            cancelRequested.set(true);
            notifyLog("❌ Ошибка промежуточного сохранения: " + e.getMessage());
        }
    }

    // Добавляем метод для очистки памяти
    private void cleanupMemory() {
        // Очищаем список компаний
        companies.clear();

        // Принудительный вызов сборщика мусора
        System.gc();

        notifyLog("🧹 Память очищена");
    }

    private void parseCurrentPage() throws InterruptedException {
        try {
            // Проверка отмены перед началом парсинга страницы
            if (cancelRequested.get()) {
                throw new InterruptedException("Операция отменена пользователем");
            }

            wait.until(ExpectedConditions.presenceOfAllElementsLocatedBy(
                    By.cssSelector("ef-poch-ro-row[ng-repeat='organization in organizations'] .register-card")));

            sleep(2000);

            // 1. ОДНОПОТОЧНЫЙ парсинг основных данных карточек
            List<Company> basicCompanies = new ArrayList<>();

            int cardCount = driver.findElements(
                    By.cssSelector("ef-poch-ro-row[ng-repeat='organization in organizations'] .register-card")).size();
            notifyLog("Найдено карточек на странице: " + cardCount);

            for (int i = 0; i < cardCount; i++) {
                // Проверка отмены перед каждой карточкой (только быстрая проверка флага)
                if (cancelRequested.get()) {
                    notifyLog("⏹️ Отмена запрошена, прерываем парсинг карточек");
                    break;
                }

                try {
                    List<WebElement> currentCards = driver.findElements(
                            By.cssSelector("ef-poch-ro-row[ng-repeat='organization in organizations'] .register-card"));

                    if (i < currentCards.size()) {
                        WebElement card = currentCards.get(i);
                        Company company = parseCompanyCard(card);
                        if (company != null && company.getProfileUrl() != null && !company.getProfileUrl().isEmpty()) {
                            basicCompanies.add(company);
                        }
                    }
                } catch (Exception e) {
                    if (e.getMessage().contains("stale element reference")) {
                        notifyLog("❌ STALE ЭЛЕМЕНТ при парсинге карточки " + (i + 1) + ", пропускаем");
                    }
                }
            }

            if (basicCompanies.isEmpty()) {
                notifyLog("⚠️ На странице не найдено компаний для парсинга");
                return;
            }

            // 2. МНОГОПОТОЧНЫЙ парсинг - проверка отмены перед запуском
            if (cancelRequested.get()) {
                throw new InterruptedException("Операция отменена пользователем");
            }

            ExecutorService executorService = Executors.newFixedThreadPool(3);
            List<CompletableFuture<Void>> futures = new ArrayList<>();

            for (Company basicCompany : basicCompanies) {
                // Проверка отмены перед добавлением каждой задачи
                if (cancelRequested.get()) {
                    notifyLog("⏹️ Отмена запрошена, прерываем запуск потоков");
                    break;
                }

                CompletableFuture<Void> future = CompletableFuture.runAsync(() -> {
                    // Проверка отмены в начале каждого потока
                    if (cancelRequested.get()) {
                        return;
                    }

                    WebDriver threadDriver = null;
                    try {
                        threadDriver = new ChromeDriver(createChromeOptions());

                        // Передаем флаг отмены в метод парсинга деталей
                        parseCompanyDetails(basicCompany, threadDriver);
                    } catch (Exception e) {
                        if (!cancelRequested.get()) {
                            notifyLog("❌ Ошибка парсинга деталей для " + basicCompany.getName() + ": " + e.getMessage());
                        }
                    } finally {
                        if (threadDriver != null) {
                            threadDriver.quit();
                        }
                    }
                }, executorService);
                futures.add(future);
            }

            // Ждем завершения с периодической проверкой отмены
            CompletableFuture<Void> allFutures = CompletableFuture.allOf(
                    futures.toArray(new CompletableFuture[0])
            );

            try {
                // Ждем с таймаутом и проверкой отмены каждую секунду
                for (int i = 0; i < 480; i++) { // 8 минут = 480 секунд
                    if (cancelRequested.get()) {
                        notifyLog("⏹️ Отмена запрошена, прерываем ожидание потоков");
                        futures.forEach(f -> f.cancel(true));
                        break;
                    }

                    if (allFutures.isDone()) {
                        break;
                    }

                    sleep(1000); // Ждем 1 секунду
                }

                if (!allFutures.isDone()) {
                    notifyLog("⚠️ Таймаут ожидания завершения потоков парсинга");
                    futures.forEach(f -> f.cancel(true));
                } else {
                    notifyLog("🎯 Парсинг страницы " + currentPage + " завершен");
                }
            } finally {
                executorService.shutdownNow(); // Принудительно завершаем executor
            }

            // Добавляем компании в общий список
            companies.addAll(basicCompanies);

        } catch (InterruptedException ie) {
            throw ie;
        } catch (TimeoutException te) {
            notifyLog("⚠️ Карточки организаций не появились: " + te.getMessage());
        } catch (Exception e) {
            notifyLog("❌ Ошибка парсинга страницы: " + e.getMessage());
        }
    }

    // Поля карточки списка: устойчивые селекторы для названия и ссылки
    private Company parseCompanyCard(WebElement card) {
        try {
            Company company = new Company();

            // 1) Название: пробуем несколько вариантов внутри заголовка
            String name = "";
            // текст заголовка
            List<WebElement> headerTitle = card.findElements(By.cssSelector(".register-card__header-title"));
            if (!headerTitle.isEmpty()) {
                name = safeTrim(headerTitle.get(0).getText());
            }
            // иногда название — это ссылка внутри заголовка
            List<WebElement> headerLinkCandidates = card.findElements(By.cssSelector(
                    ".register-card__header-title a, .register-card__header a, a.register-card__title"));
            if (!headerLinkCandidates.isEmpty()) {
                String t = safeTrim(headerLinkCandidates.get(0).getText());
                if (!t.isEmpty()) name = t;
            }
            company.setName(name);

            // 2) Ссылка на карточку: ищем в приоритетах ui-sref/ui-state/ng-href/href
            String url = findCardUrl(card);
            company.setProfileUrl(url);

            // 3) Вид организации (ng-repeat)
            List<WebElement> roleItems = card.findElements(By.cssSelector(
                    "[ng-repeat='role in organization.nsiOrganizationRoles'] .ng-binding"));
            if (!roleItems.isEmpty()) {
                List<String> roles = new ArrayList<>();
                for (WebElement it : roleItems) {
                    String val = safeTrim(it.getText()).replaceAll("\\s*;\\s*$", "");
                    if (!val.isEmpty()) roles.add(val);
                }
                if (!roles.isEmpty()) company.setType(String.join(System.lineSeparator(), roles));
            }

            // 4) Сайт (a[ng-href] либо обычный a с http)
            WebElement siteLink = firstOrNull(card, By.cssSelector("a[ng-href^='http'], a[href^='http']"));
            if (siteLink != null) {
                String siteText = safeTrim(siteLink.getText());
                String siteHref = siteLink.getAttribute("href");
                // исключить из сайта саму ссылку на профиль dom.gosuslugi (оставляем только неодоменные/внешние сайты)
                if (siteHref != null && !siteHref.contains("dom.gosuslugi.ru")) {
                    company.setWebsite(!siteText.isEmpty() ? siteText : siteHref);
                }
            }

            // 5) Адрес / Телефон по лейблам (fallback)
            List<WebElement> tables = card.findElements(By.cssSelector(".register-card__table"));
            for (WebElement table : tables) {
                for (WebElement tr : table.findElements(By.tagName("tr"))) {
                    List<WebElement> tds = tr.findElements(By.tagName("td"));
                    if (tds.size() < 2) continue;
                    String labelText = safeTrim(tds.get(0).getText());
                    String valueText = safeTrim(tds.get(1).getText());

                    if ("Фактический адрес:".equals(labelText)) {
                        company.setAddress(valueText);
                    } else if ("Контактный телефон:".equals(labelText)) {
                        company.setPhone(valueText);
                    } else if ("Сайт в сети Интернет:".equals(labelText) && company.getWebsite() == null) {
                        company.setWebsite(valueText);
                    }
                }
            }

            // минимальная валидация
            if ((company.getName() == null || company.getName().isBlank()) &&
                (company.getProfileUrl() == null || company.getProfileUrl().isBlank())) {
                notifyLog("⚠️ Карточка без названия/ссылки пропущена");
                return null;
            }
            return company;

        } catch (Exception e) {
            notifyLog("❌ Ошибка детального парсинга карточки: " + e.getMessage());
            return null;
        }
    }

    // Поиск ссылки на профиль организации внутри карточки
    private String findCardUrl(WebElement card) {
        // кандидаты ссылок: ui-sref, ui-state, ng-href, обычный href
        By[] bys = new By[]{
                By.cssSelector("a[ui-sref*='organization'][ui-sref-opts], a[ui-sref*='organization']"),
                By.cssSelector("a[ui-state*='organization']"),
                By.cssSelector("a[ng-href*='/#!/organization'], a[ng-href*='organization']"),
                By.cssSelector("a[href*='/#!/organization'], a[href*='organizationView'], a[href*='/organization/']")
        };
        for (By by : bys) {
            WebElement a = firstOrNull(card, by);
            if (a != null) {
                String href = a.getAttribute("href");
                if (href == null || href.isBlank()) href = a.getAttribute("ng-href");
                if (href != null && !href.isBlank()) {
                    return href.startsWith("/") ? BASE_URL + href : href;
                }
            }
        }
        // иногда “Подробнее” ведет на нужную ссылку
        WebElement more = firstOrNull(card, By.xpath(".//a[contains(.,'Подробнее') or contains(.,'Перейти')]"));
        if (more != null) {
            String href = more.getAttribute("href");
            if (href != null && !href.isBlank()) {
                return href.startsWith("/") ? BASE_URL + href : href;
            }
        }
        return "";
    }

    private WebElement firstOrNull(WebElement scope, By by) {
        try {
            List<WebElement> list = scope.findElements(by);
            return list.isEmpty() ? null : list.get(0);
        } catch (Exception e) {
            return null;
        }
    }

    private void parseCompanyDetails(Company company, WebDriver threadDriver) {
        if (company.getProfileUrl() == null || company.getProfileUrl().isEmpty()) {
            notifyLog("❌ Пустая ссылка для компании: " + company.getName());
            return;
        }

        // Проверка отмены в начале
        if (cancelRequested.get()) {
            return;
        }

        WebDriverWait threadWait = new WebDriverWait(threadDriver, Duration.ofSeconds(TIMEOUT_SECONDS));

        try {
            notifyLog("🔄 Переходим на страницу: " + company.getName());

            String originalWindow = threadDriver.getWindowHandle();
            ((JavascriptExecutor) threadDriver).executeScript("window.open(arguments[0], '_blank');", company.getProfileUrl());
            sleep(1000);

            // Проверка отмены после открытия вкладки
            if (cancelRequested.get()) {
                threadDriver.quit();
                return;
            }

            // Переключаемся на новую вкладку
            for (String windowHandle : threadDriver.getWindowHandles()) {
                if (!windowHandle.equals(originalWindow)) {
                    threadDriver.switchTo().window(windowHandle);
                    break;
                }
            }

            threadWait.until(ExpectedConditions.presenceOfElementLocated(By.cssSelector("body")));
            sleep(1500);

            // Проверка отмены после загрузки страницы
            if (cancelRequested.get()) {
                threadDriver.close();
                threadDriver.switchTo().window(originalWindow);
                return;
            }

            clickAdditionalInfoButton(threadDriver, threadWait);
            sleep(1500);

            // Финальная проверка отмены перед парсингом
            if (cancelRequested.get()) {
                threadDriver.close();
                threadDriver.switchTo().window(originalWindow);
                return;
            }

            parseAdditionalInfo(company, threadDriver, threadWait);

            // Закрываем вкладку
            threadDriver.close();
            threadDriver.switchTo().window(originalWindow);

        } catch (Exception e) {
            if (!cancelRequested.get()) {
                notifyLog("❌ Ошибка при парсинге детальной информации для " + company.getName() + ": " + e.getMessage());
            }
        }
    }

    // Обновленные вспомогательные методы с передачей драйвера
    private void clickAdditionalInfoButton(WebDriver driver, WebDriverWait wait) {
        try {
            List<WebElement> additionalInfoButtons = driver.findElements(By.xpath(
                    "//*[contains(text(), 'Дополнительная информация')]"
            ));
            for (WebElement button : additionalInfoButtons) {
                try {
                    if (button.isDisplayed() && button.isEnabled()) {
                        ((JavascriptExecutor) driver).executeScript("arguments[0].scrollIntoView(true);", button);
                        sleep(1000);
                        button.click();
                        return;
                    }
                } catch (Exception ignore) {
                }
            }
            notifyLog("⚠️ Не удалось найти кнопку 'Дополнительная информация', продолжаем парсинг...");
        } catch (Exception e) {
            notifyLog("❌ Ошибка при нажатии кнопки 'Дополнительная информация': " + e.getMessage());
        }
    }

    private void parseReceptionBeforeHours(Company company, WebDriver driver, WebDriverWait wait) {
        try {
            StringBuilder receptionInfo = new StringBuilder();

            // ВАЖНО: используем переданный driver (локальный для потока), а не общий
            List<WebElement> receptionBlocks = driver.findElements(By.cssSelector(
                    "ef-ppa-di-citizen-reception-info ef-ppa-di-block[header-text], ef-ppa-di-citizen-reception-info .info-card__table"
            ));

            if (!receptionBlocks.isEmpty()) {
                // Лицо, осуществляющее прием граждан
                List<WebElement> person = driver.findElements(By.cssSelector(
                        "ef-ppa-di-citizen-reception-info .info-card__val span[ng-bind='$dataItem.employeePosition']"
                ));
                if (!person.isEmpty()) {
                    String val = safeTrim(person.get(0).getText());
                    if (!val.isEmpty()) {
                        receptionInfo.append(val).append("\n");
                    }
                }

                // Адрес места приема граждан
                List<WebElement> addr = driver.findElements(By.cssSelector(
                        "ef-ppa-di-citizen-reception-info .info-card__val span[ng-bind$='| addressFormatter']"
                ));
                if (!addr.isEmpty()) {
                    String val = safeTrim(addr.get(0).getText());
                    if (!val.isEmpty()) {
                        receptionInfo.append(val).append("\n");
                    }
                }

                // Контактные телефоны
                List<WebElement> phoneSpans = driver.findElements(By.cssSelector(
                        "ef-ppa-di-citizen-reception-info ef-ppa-di-phone [ng-repeat='onePhone in data'] span[ng-bind='onePhone.value | phone']"
                ));
                if (!phoneSpans.isEmpty()) {
                    List<String> phones = new ArrayList<>();
                    for (WebElement ps : phoneSpans) {
                        String p = safeTrim(ps.getText());
                        if (!p.isEmpty()) phones.add(p);
                    }
                    if (!phones.isEmpty()) {
                        receptionInfo.append(String.join(", ", phones)).append("\n");
                    }
                }
            }

            if (!receptionInfo.isEmpty()) {
                company.setReceptionInfo(receptionInfo.toString().trim());
            } else {
                notifyLog("❌ Информация о приёме не найдена для " + company.getName());
            }

        } catch (Exception e) {
            notifyLog("❌ Ошибка парсинга информации о приёме для " + company.getName() + ": " + e.getMessage());
        }
    }

    /**
     * Читаем ТОЛЬКО "Часы приема граждан" из ef-ppa-di-citizen-reception-info hours-sheet.
     * Обрабатываем наличие "выходной", перерывы и комментарии (если колонка есть).
     * "Режим работы организации" игнорируем.
     */
    private void parseReceptionInfo(Company company, WebDriver driver, WebDriverWait wait) {
        try {
            WebElement citizenHoursContainer = null;
            // Ищем hours-sheet внутри ef-ppa-di-citizen-reception-info с alt-markup
            List<WebElement> candidates = driver.findElements(By.cssSelector(
                    "ef-ppa-di-citizen-reception-info ef-ppa-di-hours-sheet[alt-markup='true']"
            ));
            if (!candidates.isEmpty()) {
                citizenHoursContainer = candidates.get(0);
            }

            if (citizenHoursContainer == null) {
                notifyLog("🕒 Блок 'Часы приема граждан' не найден");
                return;
            }

            WebElement table = citizenHoursContainer.findElement(By.cssSelector("table.table.table-entity"));
            List<WebElement> rows = table.findElements(By.cssSelector("tbody > tr"));

            StringBuilder receptionHours = new StringBuilder();
            StringBuilder breakTimes = new StringBuilder();
            StringBuilder staffNotes = new StringBuilder();

            for (WebElement row : rows) {
                List<WebElement> tds = row.findElements(By.tagName("td"));
                if (tds.isEmpty()) continue;

                // День недели
                String day = "";
                try {
                    WebElement daySpan = row.findElement(By.cssSelector("td.table-entity_cell_dark span[ng-bind^='days[']"));
                    day = safeTrim(daySpan.getText());
                } catch (NoSuchElementException ignore) {
                    continue; // если нет дня недели — пропускаем строку
                }

                // Интервалы приема
                String begin = getTextOrEmpty(row, By.cssSelector("td:nth-of-type(2) span[ng-bind='openingHours.openHours.beginDate']"));
                String end = getTextOrEmpty(row, By.cssSelector("td:nth-of-type(2) span[ng-bind='openingHours.openHours.endDate']"));
                String workInterval = (!begin.isEmpty() && !end.isEmpty()) ? (begin + "—" + end) : "";

                // Перерыв
                String brBegin = getTextOrEmpty(row, By.cssSelector("td:nth-of-type(3) span[ng-bind='openingHours.breakHours.beginDate']"));
                String brEnd = getTextOrEmpty(row, By.cssSelector("td:nth-of-type(3) span[ng-bind='openingHours.breakHours.endDate']"));
                String breakInterval = (!brBegin.isEmpty() && !brEnd.isEmpty()) ? (brBegin + "—" + brEnd) : "";

                // Комментарий (если колонка включена)
                String comment = getTextOrEmpty(row, By.cssSelector("span[ng-bind='openingHours.comment']"));

                if (!workInterval.isEmpty()) {
                    appendLine(receptionHours, day + ": " + workInterval);
                } else {
                    // если нет интервала — возможно пустая строка, пропускаем
                    continue;
                }
                if (!breakInterval.isEmpty()) {
                    appendLine(breakTimes, day + ": " + breakInterval);
                }
                if (!comment.isEmpty()) {
                    appendLine(staffNotes, comment);
                }
            }

            if (!receptionHours.isEmpty()) {
                company.setReceptionHours(receptionHours.toString().trim());
                company.setBreakTimes(breakTimes.toString().trim());
            }

            if (!staffNotes.isEmpty()) {
                String existing = company.getNotes() != null ? company.getNotes() : "";
                company.setNotes((existing.isEmpty() ? "" : (existing + "\n")) + staffNotes.toString().trim());
            }
        } catch (Exception e) {
            notifyLog("Ошибка парсинга часов приёма: " + e.getMessage());
        }
    }

    private void parseDirectorInfo(Company company, WebDriver driver, WebDriverWait wait) {
        try {
            String fio = "";
            String position = "";

            // Ищем ФИО - исправленный селектор
            List<WebElement> fioElements = driver.findElements(By.cssSelector("div.info-card__val[ng-bind='$dataItem.fio'], div[ng-bind='$dataItem.fio']"));
            if (!fioElements.isEmpty()) {
                fio = safeTrim(fioElements.get(0).getText());
            }

            // Ищем должность - исправленный селектор
            List<WebElement> positionElements = driver.findElements(By.cssSelector("div.info-card__val[ng-bind='$dataItem.position'], div[ng-bind='$dataItem.position']"));
            if (!positionElements.isEmpty()) {
                position = safeTrim(positionElements.get(0).getText());
            }

            // Сборка результата
            if (!fio.isEmpty() || !position.isEmpty()) {
                StringBuilder sb = new StringBuilder();
                if (!fio.isEmpty()) sb.append(fio);
                if (!position.isEmpty()) {
                    if (!sb.isEmpty()) sb.append("\n");
                    sb.append(position);
                }
                company.setDirectorInfo(sb.toString());
            } else {
                notifyLog("❌ Информация о руководителе не найдена");
            }
        } catch (Exception e) {
            company.setDirectorInfo("Ошибка парсинга");
            notifyLog("❌ Ошибка парсинга информации о руководителе: " + e.getMessage());
        }
    }

    private void parseEmailInfo(Company company, WebDriver driver, WebDriverWait wait) {
        try {
            // Ищем элемент с email по селектору из span с ng-bind="data.orgEmail"
            List<WebElement> emailElements = driver.findElements(By.cssSelector("span[ng-bind='data.orgEmail']"));
            if (!emailElements.isEmpty()) {
                String email = safeTrim(emailElements.get(0).getText());
                if (!email.isEmpty()) {
                    company.setEmail(email);
                    notifyLog("✅ Найден email: " + email);
                }
            } else {
                notifyLog("⚠️ Email не найден для " + company.getName());
            }
        } catch (Exception e) {
            notifyLog("❌ Ошибка парсинга email для " + company.getName() + ": " + e.getMessage());
        }
    }

    private String getTextOrEmpty(WebElement scope, By by) {
        try {
            WebElement el = (scope == null) ? driver.findElement(by) : scope.findElement(by);
            String t = el.getText();
            return t == null ? "" : t.trim();
        } catch (Exception e) {
            return "";
        }
    }

    private void appendLine(StringBuilder sb, String line) {
        if (line == null || line.trim().isEmpty()) return;
        if (!sb.isEmpty()) sb.append("\n");
        sb.append(line.trim());
    }

    private String safeTrim(String s) {
        return s == null ? "" : s.trim();
    }

    private void parseAdditionalInfo(Company company, WebDriver driver, WebDriverWait wait) {
        try {
            parseReceptionBeforeHours(company, driver, wait); // Прием граждан: лицо/адрес/телефоны
            parseReceptionInfo(company, driver, wait);        // Часы приема граждан
            parseDirectorInfo(company, driver, wait);         // Руководитель
            parseEmailInfo(company, driver, wait);           // Email
            parseNotesInfo(company, driver, wait);            // Примечания
        } catch (Exception e) {
            notifyLog("❌ Не удалось найти дополнительную информацию для " + company.getName() + ": " + e.getMessage());
        }
    }

    private void parseNotesInfo(Company company, WebDriver driver, WebDriverWait wait) {
        try {
            StringBuilder notes = new StringBuilder();
            notes.append(parseSpecificNote("Примечание", driver));
            notes.append(parseSpecificNote("Дополнительная информация", driver));
            notes.append(parseSpecificNote("Особые условия", driver));
            notes.append(parseSpecificNote("Комментарий", driver));

            if (!notes.isEmpty()) {
                company.setNotes(notes.toString().trim());
            }
        } catch (Exception e) {
            notifyLog("❌ Ошибка парсинга примечаний: " + e.getMessage());
        }
    }

    private String parseSpecificNote(String fieldName, WebDriver driver) {
        try {
            List<WebElement> elements = driver.findElements(By.xpath(
                    "//*[contains(text(), '" + fieldName + "')]"
            ));
            for (WebElement element : elements) {
                try {
                    WebElement valueElement = element.findElement(By.xpath(
                            "./following-sibling::div[contains(@class, 'info-card_val')] | " +
                            "./ancestor::tr[1]//div[contains(@class, 'info-card_val')] | " +
                            "./following::span[1] | ./following::div[1]"
                    ));
                    String value = valueElement.getText().trim();
                    if (!value.isEmpty() && !isJustDayOfWeek(value)) {
                        return value + "\n";
                    }
                } catch (Exception ignore) {
                }
            }
        } catch (Exception ignore) {
        }
        return "";
    }

    private boolean isJustDayOfWeek(String text) {
        if (text == null || text.trim().isEmpty()) return false;
        String cleanedText = text.trim().toLowerCase();
        return cleanedText.matches("^(понедельник|вторник|среда|четверг|пятница|суббота|воскресенье)$");
    }

    private boolean goToNextPage() {
        try {
            WebElement currentPage = driver.findElement(By.cssSelector(".pagination .active"));
            if (currentPage != null) {
                String currentPageText = currentPage.getText();

                int currentPageNum = Integer.parseInt(currentPageText);
                WebElement nextPage = driver.findElement(By.xpath("//a[text()='" + (currentPageNum + 1) + "']"));
                if (nextPage != null && nextPage.isEnabled()) {
                    ((JavascriptExecutor) driver).executeScript("arguments[0].scrollIntoView(true);", nextPage);
                    sleep(1000);
                    nextPage.click();

                    // Ждем загрузки новой страницы
                    wait.until(ExpectedConditions.presenceOfAllElementsLocatedBy(
                            By.cssSelector("ef-poch-ro-row[ng-repeat='organization in organizations'] .register-card")));
                    sleep(2000);

                    notifyLog("➡️ Переход на страницу " + (currentPageNum + 1));
                    return true;
                }
            }
            return false;
        } catch (Exception e) {
            // Просто возвращаем false - страницы закончились
            return false;
        }
    }

    private void createHeaders(Sheet sheet, Workbook workbook) {
        CellStyle headerStyle = createHeaderStyle(workbook);
        Row headerRow = sheet.createRow(0);
        String[] headers = {
                "Наименование", "Вид организации", "Фактический адрес", "Сайт", "Телефон",
                "Email", "Информация о приёме", "Часы приёма", "Перерыв", "Примечание",
                "Руководитель", "Ссылка на карточку"
        };
        for (int i = 0; i < headers.length; i++) {
            Cell cell = headerRow.createCell(i);
            cell.setCellValue(headers[i]);
            cell.setCellStyle(headerStyle);
        }
    }

    private void createCompanyRow(Row row, Company company, CellStyle defaultStyle, CellStyle linkStyle, CreationHelper createHelper) {
        Cell nameCell = row.createCell(0);
        nameCell.setCellValue(company.getName() != null ? company.getName() : "");
        nameCell.setCellStyle(defaultStyle);

        setCellValue(row, 1, company.getType(), defaultStyle);
        setCellValue(row, 2, company.getAddress(), defaultStyle);
        setCellValue(row, 3, company.getWebsite(), defaultStyle);
        setCellValue(row, 4, company.getPhone(), defaultStyle);
        setCellValue(row, 5, company.getEmail(), defaultStyle);
        setCellValue(row, 6, company.getReceptionInfo(), defaultStyle);
        setCellValue(row, 7, company.getReceptionHours(), defaultStyle);
        setCellValue(row, 8, company.getBreakTimes(), defaultStyle);
        setCellValue(row, 9, company.getNotes(), defaultStyle);
        setCellValue(row, 10, company.getDirectorInfo(), defaultStyle);

        Cell linkCell = row.createCell(11);
        if (company.getProfileUrl() != null && !company.getProfileUrl().isEmpty()) {
            linkCell.setCellValue("Открыть карточку");
            Hyperlink link = createHelper.createHyperlink(HyperlinkType.URL);
            link.setAddress(company.getProfileUrl());
            linkCell.setHyperlink(link);
            linkCell.setCellStyle(linkStyle);
        } else {
            linkCell.setCellValue("Нет ссылки");
            linkCell.setCellStyle(defaultStyle);
        }
    }

    private void updateCompanyRow(Row row, Company company, CellStyle defaultStyle, CellStyle linkStyle, CreationHelper createHelper) {
        setCellValue(row, 1, company.getType(), defaultStyle);
        setCellValue(row, 2, company.getAddress(), defaultStyle);
        setCellValue(row, 3, company.getWebsite(), defaultStyle);
        setCellValue(row, 4, company.getPhone(), defaultStyle);
        setCellValue(row, 5, company.getEmail(), defaultStyle);
        setCellValue(row, 6, company.getReceptionInfo(), defaultStyle);
        setCellValue(row, 7, company.getReceptionHours(), defaultStyle);
        setCellValue(row, 8, company.getBreakTimes(), defaultStyle);
        setCellValue(row, 9, company.getNotes(), defaultStyle);
        setCellValue(row, 10, company.getDirectorInfo(), defaultStyle);

        Cell linkCell = row.getCell(11);
        if (linkCell == null) {
            linkCell = row.createCell(11);
        }
        if (company.getProfileUrl() != null && !company.getProfileUrl().isEmpty()) {
            linkCell.setCellValue("Открыть карточку");
            Hyperlink link = createHelper.createHyperlink(HyperlinkType.URL);
            link.setAddress(company.getProfileUrl());
            linkCell.setHyperlink(link);
            linkCell.setCellStyle(linkStyle);
        } else {
            linkCell.setCellValue("Нет ссылки");
            linkCell.setCellStyle(defaultStyle);
        }
    }

    private void setCellValue(Row row, int cellIndex, String value, CellStyle style) {
        Cell cell = row.getCell(cellIndex);
        if (cell == null) {
            cell = row.createCell(cellIndex);
        }
        cell.setCellValue(value != null ? value : "");
        cell.setCellStyle(style);
    }

    private CellStyle createHeaderStyle(Workbook workbook) {
        CellStyle headerStyle = workbook.createCellStyle();
        Font headerFont = workbook.createFont();
        headerFont.setFontName("Times New Roman");
        headerFont.setFontHeightInPoints((short) 12);
        headerFont.setBold(true);
        headerStyle.setFont(headerFont);
        headerStyle.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
        headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        headerStyle.setWrapText(true);
        headerStyle.setVerticalAlignment(VerticalAlignment.TOP);
        return headerStyle;
    }

    private CellStyle createDefaultStyle(Workbook workbook) {
        CellStyle defaultStyle = workbook.createCellStyle();
        Font defaultFont = workbook.createFont();
        defaultFont.setFontName("Times New Roman");
        defaultFont.setFontHeightInPoints((short) 12);
        defaultStyle.setFont(defaultFont);
        defaultStyle.setWrapText(true);
        defaultStyle.setVerticalAlignment(VerticalAlignment.TOP);
        return defaultStyle;
    }

    private CellStyle createLinkStyle(Workbook workbook) {
        CellStyle linkStyle = workbook.createCellStyle();
        Font linkFont = workbook.createFont();
        linkFont.setFontName("Times New Roman");
        linkFont.setFontHeightInPoints((short) 12);
        linkFont.setUnderline(Font.U_SINGLE);
        linkFont.setColor(IndexedColors.BLUE.getIndex());
        linkStyle.setFont(linkFont);
        linkStyle.setWrapText(true);
        linkStyle.setVerticalAlignment(VerticalAlignment.TOP);
        return linkStyle;
    }

    private void saveToExcel() {
        if (companies.isEmpty()) {
            notifyLog("❌ Нет данных для сохранения");
            return;
        }

        boolean fileExists = false;
        String fileName = "Управляющие компании " + region + " " + LocalDate.now().getYear() + ".xlsx";

        if (new File("Управляющие компании " + region + " " + LocalDate.now().minusYears(1).getYear() + ".xlsx").exists()) {
            fileExists = true;
            fileName = "Управляющие компании " + region + " " + LocalDate.now().minusYears(1).getYear() + ".xlsx";
        } else if (new File("Управляющие компании " + region + " " + LocalDate.now().getYear() + ".xlsx").exists()) {
            fileExists = true;
        }

        try {
            Workbook workbook;
            Sheet sheet;

            if (fileExists) {
                try (FileInputStream fis = new FileInputStream(fileName)) {
                    workbook = new XSSFWorkbook(fis);
                }
                sheet = workbook.getSheet("Компании");
                if (sheet == null) {
                    sheet = workbook.createSheet("Компании");
                    createHeaders(sheet, workbook);
                }
            } else {
                workbook = new XSSFWorkbook();
                sheet = workbook.createSheet("Компании");
                createHeaders(sheet, workbook);
            }

            CellStyle defaultStyle = createDefaultStyle(workbook);
            CellStyle linkStyle = createLinkStyle(workbook);

            Map<String, Integer> existingCompanies = new HashMap<>();
            if (fileExists && sheet.getPhysicalNumberOfRows() > 1) {
                for (int i = 1; i <= sheet.getLastRowNum(); i++) {
                    Row row = sheet.getRow(i);
                    if (row != null && row.getCell(0) != null) {
                        String companyName = row.getCell(0).getStringCellValue();
                        if (companyName != null && !companyName.trim().isEmpty()) {
                            existingCompanies.put(companyName.trim(), i);
                        }
                    }
                }
            }

            CreationHelper createHelper = workbook.getCreationHelper();
            int newRowsCount = 0;
            int updatedRowsCount = 0;

            for (Company company : companies) {
                if (company.getName() == null || company.getName().trim().isEmpty()) {
                    continue;
                }

                String companyName = company.getName().trim();
                Integer existingRowIndex = existingCompanies.get(companyName);

                if (existingRowIndex != null) {
                    updateCompanyRow(sheet.getRow(existingRowIndex), company, defaultStyle, linkStyle, createHelper);
                    updatedRowsCount++;
                } else {
                    int newRowIndex = sheet.getLastRowNum() + 1;
                    Row row = sheet.createRow(newRowIndex);
                    createCompanyRow(row, company, defaultStyle, linkStyle, createHelper);
                    newRowsCount++;
                }
            }

            for (int i = 0; i < 11; i++) {
                sheet.autoSizeColumn(i);
                sheet.setColumnWidth(i, sheet.getColumnWidth(i) + 512);
            }

            for (int i = 0; i <= sheet.getLastRowNum(); i++) {
                Row row = sheet.getRow(i);
                if (row != null) {
                    row.setHeight((short) -1);
                }
            }

            sheet.setAutoFilter(new CellRangeAddress(0, sheet.getLastRowNum(), 0, 11));

            try (FileOutputStream fos = new FileOutputStream("Управляющие компании " + region + " " + LocalDate.now().getYear() + ".xlsx")) {
                workbook.write(fos);
            }

            workbook.close();

            notifyLog("💾 Данные " + (fileExists ? "обновлены" : "сохранены") + " в файл: " + fileName);
            if (fileExists) {
                notifyLog("📊 Обновлено: " + updatedRowsCount + " записей, Добавлено: " + newRowsCount + " новых записей");
            }

        } catch (IOException e) {
            notifyLog("❌ Ошибка сохранения в Excel: " + e.getMessage());
        }
    }

    public static void main(String[] args) {
        DomGosuslugiParser parser = new DomGosuslugiParser();
        parser.parseOrganizations();
    }
}