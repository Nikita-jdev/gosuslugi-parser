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
import org.openqa.selenium.OutputType;
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
import java.util.concurrent.atomic.AtomicBoolean;

import static java.lang.Thread.sleep;

public class DomGosuslugiHousesParser {
    private WebDriver driver;
    private WebDriverWait wait;
    private final List<House> houses = new ArrayList<>();

    private static final String TARGET_URL = "https://dom.gosuslugi.ru/#!/houses";
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

    public void parseHouses() {
        try {
            checkSeleniumSetup();

            notifyStatus("Запуск драйвера...");
            initDriver();
            notifyLog("🚀 Запуск парсера объектов жилищного фонда...");

            driver.get(TARGET_URL);
            wait.until(ExpectedConditions.presenceOfElementLocated(By.cssSelector("body")));
            sleep(5000);
            checkCancelled();

            //selectRegionFilter();
            selectSpbFilter();
            checkCancelled();

            clickSearchButton();
            sleep(1000);
            checkCancelled();

            selectItemsPerPage("100");
            sleep(3000);

            parseAllPages(startPage);

            notifyLog("📊 Всего найдено домов: " + houses.size());

            notifyStatus("Сохранение в Excel...");
            saveToExcel();

            notifyFinished(true, "Успешно: сохранено " + houses.size() + " записей");
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
        //options.addArguments("--headless=new");
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
                    if (option.getText().contains(region)) {
                        dropdown.selectByVisibleText(option.getText());
                        regionFound = true;
                        notifyLog("✅ Выбран регион: " + region);
                        break;
                    }
                }

                if (!regionFound) {
                    notifyLog("⚠️ Регион '" + region + "' не найден в списке, используется первый доступный");
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

    private void selectSpbFilter() {
        try {
            List<WebElement> selects = driver.findElements(By.cssSelector("select"));
            if (!selects.isEmpty()) {
                Select dropdown = new Select(selects.get(0));
                for (WebElement option : dropdown.getOptions()) {
                    if (option.getText().contains("Чукотский автономный округ")) {
                        dropdown.selectByVisibleText(option.getText());
                        break;
                    }
                }
            }
            sleep(1000);
        } catch (Exception e) {
            notifyLog("Ошибка выбора фильтра: " + e.getMessage());
        }
    }

    private void clickSearchButton() {
        try {
            // Ищем кнопку поиска по различным селекторам
            WebElement button = null;
            String[] buttonSelectors = {
                    "button[type='submit']",
                    "button.btn-prime",
                    "button[class*='btn-prime']",
                    "button[ng-click*='search']",
                    "button:contains('Найти')"
            };

            for (String selector : buttonSelectors) {
                try {
                    List<WebElement> buttons = driver.findElements(By.cssSelector(selector));
                    if (!buttons.isEmpty()) {
                        button = buttons.get(0);
                        break;
                    }
                } catch (Exception e) {
                    continue;
                }
            }

            if (button != null) {
                ((JavascriptExecutor) driver).executeScript("arguments[0].scrollIntoView(true);", button);
                sleep(1000);
                ((JavascriptExecutor) driver).executeScript("arguments[0].click();", button);
            } else {
                notifyLog("❌ Кнопка поиска не найдена");
            }

            sleep(3000); // Ждем загрузки результатов

        } catch (Exception e) {
            notifyLog("❌ Ошибка при нажатии кнопки 'Найти': " + e.getMessage());
        }
    }

    private void selectItemsPerPage(String countPerPage) {
        try {
            // Ждем появления элемента выбора количества
            wait.until(ExpectedConditions.presenceOfElementLocated(
                    By.cssSelector("select.output-by_select, .output-by select, select[ng-model*='itemsPerPage'], select[ng-model*='pagination']")
            ));

            WebElement selectElement = null;
            String[] selectors = {
                    "select.output-by_select",
                    ".output-by select",
                    "select[ng-model*='itemsPerPage']",
                    "select[ng-model*='pagination']",
                    "select#count"
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
                    notifyLog("✅ Выбрано элементов на странице: " + countPerPage);
                } catch (Exception e) {
                    // Пробуем выбрать по видимому тексту
                    try {
                        dropdown.selectByVisibleText(countPerPage);
                        notifyLog("✅ Выбрано элементов на странице по тексту: " + countPerPage);
                    } catch (Exception e2) {
                        notifyLog("❌ Ошибка выбора количества элементов: " + e2.getMessage());
                    }
                }
                sleep(1500); // Ждем обновления контента
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
                if (cancelRequested.get()) {
                    throw new InterruptedException("Операция отменена пользователем");
                }

                notifyPageProgress(currentPage, totalPages);
                notifyLog("📄 Обработка страницы " + currentPage + " из " + totalPages);

                waitForPageLoad(currentPage);
                parseCurrentPage();

                if (!houses.isEmpty()) {
                    notifyStatus("Сохранение данных страницы " + currentPage + "...");
                    saveIntermediateResults();
                    cleanupMemory();
                }

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
            if (!houses.isEmpty()) {
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
                waitForPageLoad(choicePage);
                waitForModalToDisappear(); // Ждем исчезновения модального окна
                sleep(1000);

                WebElement pageLink = driver.findElement(By.xpath("//a[text()='" + (choicePage + 2) + "']"));
                if (pageLink != null && pageLink.isEnabled()) {
                    ((JavascriptExecutor) driver).executeScript("arguments[0].scrollIntoView(true);", pageLink);
                    sleep(1000);

                    // Кликаем через JavaScript чтобы обойти перехват клика
                    ((JavascriptExecutor) driver).executeScript("arguments[0].click();", pageLink);

                    countingPage -= 2;
                    choicePage += 2;

                    notifyLog("➡️ Переход на страницу " + choicePage);
                }
            }

            if (countingPage == 2) {
                waitForPageLoad(choicePage);
                waitForModalToDisappear(); // Ждем исчезновения модального окна
                sleep(1000);

                WebElement pageLink = driver.findElement(By.xpath("//a[text()='" + (choicePage + 1) + "']"));
                if (pageLink != null && pageLink.isEnabled()) {
                    ((JavascriptExecutor) driver).executeScript("arguments[0].scrollIntoView(true);", pageLink);
                    sleep(1000);

                    // Кликаем через JavaScript
                    ((JavascriptExecutor) driver).executeScript("arguments[0].click();", pageLink);
                    sleep(1000);
                }
            }

            notifyLog("➡️ Переход на страницу " + pageNumber);
            sleep(2000);
        } catch (Exception e) {
            notifyLog("❌ Ошибка перехода на страницу " + pageNumber + ": " + e.getMessage());
        }
    }

    private void waitForModalToDisappear() {
        try {
            // Ждем исчезновения модального окна
            wait.until(ExpectedConditions.invisibilityOfElementLocated(
                    By.cssSelector(".modal-backdrop, .modal, [role='dialog']")));
            sleep(500);
        } catch (Exception e) {
            // Если модального окна нет, просто продолжаем
        }
    }

    private void waitForPageLoad(int expectedPage) {
        try {
            // Сначала ждем исчезновения модального окна
            waitForModalToDisappear();

            // Ждем исчезновения индикатора загрузки если есть
            try {
                wait.until(ExpectedConditions.invisibilityOfElementLocated(
                        By.cssSelector(".loading, .spinner, [data-ng-show='loading']")));
            } catch (Exception e) {
                // Игнорируем, если нет индикатора загрузки
            }

            // Ждем появления карточек домов с таймаутом и проверкой количества
            wait.until((WebDriver d) -> {
                try {
                    // Проверяем, что карточки загрузились и их достаточно
                    List<WebElement> cards = driver.findElements(
                            By.cssSelector(".register-card[ng-repeat*='house in searchResults.items']"));
                    return !cards.isEmpty() && cards.size() >= 50; // Минимум 50 карточек
                } catch (Exception e) {
                    return false;
                }
            });

            // Дополнительная проверка, что данные карточек загружены (не пустые адреса)
            wait.until((WebDriver d) -> {
                try {
                    List<WebElement> cards = driver.findElements(
                            By.cssSelector(".register-card[ng-repeat*='house in searchResults.items']"));
                    if (cards.isEmpty()) return false;

                    // Проверяем первую карточку на наличие адреса
                    WebElement firstCard = cards.get(0);
                    List<WebElement> addressElements = firstCard.findElements(By.cssSelector(
                            ".register-card__header-title .cnt-link-hover.ng-binding"));
                    return !addressElements.isEmpty() &&
                           !addressElements.get(0).getText().trim().isEmpty();
                } catch (Exception e) {
                    return false;
                }
            });

            // Ждем, пока активная страница в пагинации станет ожидаемой
            wait.until((WebDriver d) -> {
                try {
                    int currentPage = getCurrentPageNumber();
                    return currentPage == expectedPage;
                } catch (Exception e) {
                    return false;
                }
            });

            // Финальная задержка для полной стабилизации
            sleep(2000);

        } catch (Exception e) {
            notifyLog("⚠️ Ожидание загрузки страницы " + expectedPage + " завершилось с ошибкой: " + e.getMessage());
            // Пробуем продолжить, возможно страница все же частично загружена
        }
    }

    private int getCurrentPageNumber() {
        try {
            // Ищем активную страницу в пагинации
            List<WebElement> pageLinks = driver.findElements(By.cssSelector(
                    ".pagination a, [ng-repeat*='page'] a, .page-link"
            ));

            for (WebElement page : pageLinks) {
                try {
                    WebElement parent = page.findElement(By.xpath("./.."));
                    if (parent.getAttribute("class").contains("active") ||
                        parent.getAttribute("class").contains("current")) {
                        return Integer.parseInt(page.getText().trim());
                    }
                } catch (Exception e) {
                    // Продолжаем поиск
                }
            }

            // Альтернативный способ
            WebElement activePage = driver.findElement(By.cssSelector(
                    ".pagination .active, .current-page, [aria-current='page']"
            ));
            return Integer.parseInt(activePage.getText().trim());

        } catch (Exception e) {
            notifyLog("⚠️ Не удалось определить текущую страницу");
            return 1;
        }
    }

    private int getTotalPages() {
        try {
            // Способ 1: Ищем элемент с классом pagination-base__static-text (число страниц)
            List<WebElement> totalPagesElements = driver.findElements(By.xpath(
                    "//span[contains(@class, 'pagination-base__static-text') and string-length(normalize-space(text())) > 0]"
            ));

            if (!totalPagesElements.isEmpty()) {
                // Берем последний найденный элемент (на случай если их несколько)
                WebElement lastElement = totalPagesElements.get(totalPagesElements.size() - 1);
                String pageText = lastElement.getText().trim();

                // Проверяем, что текст содержит только цифры
                if (pageText.matches("\\d+")) {
                    int totalPages = Integer.parseInt(pageText);
                    if (totalPages >= 0) {
                        return totalPages;
                    }
                }
            }

            notifyLog("⚠️ Не удалось определить общее количество страниц");
            return 1;

        } catch (Exception e) {
            notifyLog("❌ Ошибка при получении количества страниц: " + e.getMessage());
            return 1;
        }
    }

    private void saveIntermediateResults() {
        if (houses.isEmpty()) {
            return;
        }

        try {
            String fileName = "Объекты жилищного фонда " + region + " " + LocalDate.now().getYear() + ".xlsx";
            boolean fileExists = new File(fileName).exists();

            Workbook workbook;
            Sheet sheet;

            if (fileExists) {
                try (FileInputStream fis = new FileInputStream(fileName)) {
                    workbook = new XSSFWorkbook(fis);
                }
                sheet = workbook.getSheet("Дома");
                if (sheet == null) {
                    sheet = workbook.createSheet("Дома");
                    createHeaders(sheet, workbook);
                }
            } else {
                workbook = new XSSFWorkbook();
                sheet = workbook.createSheet("Дома");
                createHeaders(sheet, workbook);
            }

            CellStyle defaultStyle = createDefaultStyle(workbook);
            CellStyle linkStyle = createLinkStyle(workbook);
            CreationHelper createHelper = workbook.getCreationHelper();

            Map<String, Integer> existingHouses = new HashMap<>();
            if (fileExists && sheet.getPhysicalNumberOfRows() > 1) {
                for (int i = 1; i <= sheet.getLastRowNum(); i++) {
                    Row row = sheet.getRow(i);
                    if (row != null && row.getCell(0) != null) {
                        String houseAddress = row.getCell(0).getStringCellValue();
                        if (houseAddress != null && !houseAddress.trim().isEmpty()) {
                            existingHouses.put(houseAddress.trim(), i);
                        }
                    }
                }
            }

            int newRowsCount = 0;
            int updatedRowsCount = 0;

            for (House house : houses) {
                if (house.getAddress() == null || house.getAddress().trim().isEmpty()) {
                    continue;
                }

                String houseAddress = house.getAddress().trim();
                Integer existingRowIndex = existingHouses.get(houseAddress);

                if (existingRowIndex != null) {
                    updateHouseRow(sheet.getRow(existingRowIndex), house, defaultStyle, linkStyle, createHelper);
                    updatedRowsCount++;
                } else {
                    int newRowIndex = sheet.getLastRowNum() + 1;
                    Row row = sheet.createRow(newRowIndex);
                    createHouseRow(row, house, defaultStyle, linkStyle, createHelper);
                    newRowsCount++;
                    existingHouses.put(houseAddress, newRowIndex);
                }
            }

            // Авто-размер колонок
            for (int i = 0; i < 6; i++) {
                sheet.autoSizeColumn(i);
                sheet.setColumnWidth(i, sheet.getColumnWidth(i) + 512);
            }

            sheet.setAutoFilter(new CellRangeAddress(0, sheet.getLastRowNum(), 0, 5));

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

    private void cleanupMemory() {
        houses.clear();
        System.gc();
        notifyLog("🧹 Память очищена");
    }

    private void parseCurrentPage() throws InterruptedException {
        try {
            if (cancelRequested.get()) {
                throw new InterruptedException("Операция отменена пользователем");
            }

            // Дополнительная проверка, что страница полностью загружена
            if (!isPageFullyLoaded(currentPage)) {
                notifyLog("⚠️ Страница " + currentPage + " не полностью загружена, повторная попытка...");
                waitForPageLoad(currentPage); // Повторная попытка
            }

            // Ждем появления карточек домов с правильным селектором
            wait.until(ExpectedConditions.presenceOfAllElementsLocatedBy(
                    By.cssSelector(".register-card[ng-repeat*='house in searchResults.items']")));

            sleep(2000);

            List<House> pageHouses = new ArrayList<>();

            int cardCount = driver.findElements(
                    By.cssSelector(".register-card[ng-repeat*='house in searchResults.items']")).size();
            notifyLog("Найдено карточек домов на странице: " + cardCount);

            for (int i = 0; i < cardCount; i++) {
                if (cancelRequested.get()) {
                    notifyLog("⏹️ Отмена запрошена, прерываем парсинг карточек");
                    break;
                }

                try {
                    List<WebElement> currentCards = driver.findElements(
                            By.cssSelector(".register-card[ng-repeat*='house in searchResults.items']"));

                    if (currentCards.isEmpty()) {
                        currentCards = driver.findElements(By.cssSelector(".register-card"));
                    }

                    if (i < currentCards.size()) {
                        WebElement card = currentCards.get(i);
                        House house = parseHouseCard(card);
                        if (house != null && house.getAddress() != null && !house.getAddress().isEmpty()) {
                            pageHouses.add(house);
                            notifyLog("✅ Обработана карточка: " + house.getAddress());
                        }
                    }
                } catch (Exception e) {
                    if (e.getMessage().contains("stale element reference")) {
                        notifyLog("❌ STALE ЭЛЕМЕНТ при парсинге карточки " + (i + 1) + ", пропускаем");
                    } else {
                        notifyLog("❌ Ошибка при парсинге карточки " + (i + 1) + ": " + e.getMessage());
                    }
                }
            }

            if (pageHouses.isEmpty()) {
                notifyLog("⚠️ На странице не найдено домов для парсинга");
                return;
            }

            houses.addAll(pageHouses);
            notifyLog("🎯 Парсинг страницы " + currentPage + " завершен, собрано: " + pageHouses.size() + " домов");

        } catch (InterruptedException ie) {
            throw ie;
        } catch (TimeoutException te) {
            notifyLog("⚠️ Карточки домов не появились: " + te.getMessage());
            // Попробуем сделать скриншот для отладки
            try {
                File screenshot = ((ChromeDriver) driver).getScreenshotAs(OutputType.FILE);
                notifyLog("📸 Сделан скриншот для отладки");
            } catch (Exception e) {
                notifyLog("❌ Не удалось сделать скриншот: " + e.getMessage());
            }
        } catch (Exception e) {
            notifyLog("❌ Ошибка парсинга страницы: " + e.getMessage());
        }
    }

    private House parseHouseCard(WebElement card) {
        try {
            House house = new House();

            // 1) Адрес - правильный селектор из HTML
            String address = "";
            List<WebElement> addressElements = card.findElements(By.cssSelector(
                    ".register-card__header-title .cnt-link-hover.ng-binding"
            ));
            if (!addressElements.isEmpty()) {
                address = safeTrim(addressElements.get(0).getText());
            }

            if (address.isEmpty()) {
                // Альтернативный поиск адреса
                addressElements = card.findElements(By.cssSelector(".register-card__header-title .ng-binding"));
                if (!addressElements.isEmpty()) {
                    address = safeTrim(addressElements.get(0).getText());
                }
            }

            house.setAddress(address);

            // 2) Ссылка на карточку - из кнопки "Сведения об объекте жилищного фонда"
            String url = findCardUrl(card);
            house.setProfileUrl(url);

            // 3) Парсим таблицы с характеристиками
            List<WebElement> tables = card.findElements(By.cssSelector(".register-card__table"));
            for (WebElement table : tables) {
                for (WebElement tr : table.findElements(By.tagName("tr"))) {
                    List<WebElement> tds = tr.findElements(By.tagName("td"));
                    if (tds.size() < 2) continue;

                    String labelText = safeTrim(tds.get(0).getText());
                    String valueText = safeTrim(tds.get(1).getText());

                    // Нормализуем текст метки (убираем переносы строк)
                    String normalizedLabel = labelText.replaceAll("\\s+", " ").trim();

                    // Только нужные поля
                    switch (normalizedLabel) {
                        case "Год ввода в эксплуатацию:":
                            house.setCommissioningYear(valueText);
                            break;
                        case "Количество этажей:":
                            house.setFloorsCount(valueText);
                            break;
                        case "Управляющая организация:":
                            house.setManagementOrganization(valueText);
                            break;
                        case "Количество помещений (жилых/нежилых):":
                            house.setPremisesCount(valueText);
                            break;
                    }
                }
            }

            // минимальная валидация
            if ((house.getAddress() == null || house.getAddress().isBlank())) {
                notifyLog("⚠️ Карточка дома без адреса пропущена");
                return null;
            }

            return house;

        } catch (Exception e) {
            notifyLog("❌ Ошибка парсинга карточки дома: " + e.getMessage());
            return null;
        }
    }

    private String findCardUrl(WebElement card) {
        try {
            WebElement houseLink = card.findElement(By.cssSelector("a[ng-click*='viewHouse']"));
            String ngClick = houseLink.getAttribute("ng-click");

            if (ngClick != null && ngClick.contains("viewHouse")) {
                // Получаем house данные через Angular scope
                String script =
                        "var card = arguments[0]; " +
                        "var link = card.querySelector('[ng-click*=\"viewHouse\"]'); " +
                        "var scope = angular.element(link).scope(); " +
                        "if (scope && scope.house) { " +
                        "    return { " +
                        "        guid: scope.house.guid, " +
                        "        typeCode: scope.house.houseType ? scope.house.houseType.code : '1' " +
                        "    }; " +
                        "} " +
                        "return null;";

                @SuppressWarnings("unchecked")
                Map<String, Object> houseData = (Map<String, Object>) ((JavascriptExecutor) driver).executeScript(script, card);

                if (houseData != null) {
                    String guid = (String) houseData.get("guid");
                    Object typeCodeObj = houseData.get("typeCode");
                    String typeCode = typeCodeObj != null ? typeCodeObj.toString() : "1";

                    if (guid != null && !guid.isEmpty()) {
                        return BASE_URL + "/#!/house-view?guid=" + guid + "&typeCode=" + typeCode;
                    }
                }
            }
        } catch (Exception e) {
            System.out.println("Ошибка при извлечении house URL: " + e.getMessage());
        }

        return "";
    }

    private String safeTrim(String s) {
        return s == null ? "" : s.trim();
    }

    private boolean goToNextPage() {
        try {
            int currentPageNum = getCurrentPageNumber();
            WebElement nextPage = driver.findElement(By.xpath("//a[text()='" + (currentPageNum + 1) + "']"));

            if (nextPage != null && nextPage.isEnabled()) {
                ((JavascriptExecutor) driver).executeScript("arguments[0].scrollIntoView(true);", nextPage);
                sleep(1000);
                nextPage.click();

                // Ждем загрузки новой страницы с улучшенной проверкой
                waitForPageLoad(currentPageNum + 1);

                notifyLog("➡️ Переход на страницу " + (currentPageNum + 1));
                return true;
            }
            return false;
        } catch (Exception e) {
            notifyLog("❌ Ошибка перехода на следующую страницу: " + e.getMessage());
            return false;
        }
    }

    private boolean isPageFullyLoaded(int expectedPage) {
        try {
            // Проверяем наличие карточек
            List<WebElement> cards = driver.findElements(
                    By.cssSelector(".register-card[ng-repeat*='house in searchResults.items']"));

            if (cards.isEmpty()) {
                notifyLog("⚠️ Карточки не найдены на странице " + expectedPage);
                return false;
            }

            // Проверяем, что текущая страница правильная
            int actualPage = getCurrentPageNumber();
            if (actualPage != expectedPage) {
                notifyLog("⚠️ Несоответствие страниц: ожидалась " + expectedPage + ", получена " + actualPage);
                return false;
            }

            // Проверяем, что карточки содержат данные
            WebElement firstCard = cards.get(0);
            List<WebElement> addressElements = firstCard.findElements(By.cssSelector(
                    ".register-card__header-title .cnt-link-hover.ng-binding"));

            boolean dataLoaded = !addressElements.isEmpty() &&
                                 !addressElements.get(0).getText().trim().isEmpty();

            if (!dataLoaded) {
                notifyLog("⚠️ Данные в карточках не загружены на странице " + expectedPage);
            }

            return dataLoaded;

        } catch (Exception e) {
            notifyLog("❌ Ошибка проверки загрузки страницы " + expectedPage + ": " + e.getMessage());
            return false;
        }
    }

    private void createHeaders(Sheet sheet, Workbook workbook) {
        CellStyle headerStyle = createHeaderStyle(workbook);
        Row headerRow = sheet.createRow(0);
        String[] headers = {
                "Адрес", "Год ввода в эксплуатацию", "Количество этажей",
                "Управляющая организация", "Количество помещений\n(жилых/нежилых)", "Ссылка на карточку"
        };
        for (int i = 0; i < headers.length; i++) {
            Cell cell = headerRow.createCell(i);
            cell.setCellValue(headers[i]);
            cell.setCellStyle(headerStyle);
        }
    }

    private void createHouseRow(Row row, House house, CellStyle defaultStyle, CellStyle linkStyle, CreationHelper createHelper) {
        setCellValue(row, 0, house.getAddress(), defaultStyle);
        setCellValue(row, 0, house.getAddress(), defaultStyle);
        setCellValue(row, 1, house.getCommissioningYear(), defaultStyle);
        setCellValue(row, 2, house.getFloorsCount(), defaultStyle);
        setCellValue(row, 3, house.getManagementOrganization(), defaultStyle);
        setCellValue(row, 4, house.getPremisesCount(), defaultStyle);

        Cell linkCell = row.createCell(5);
        if (house.getProfileUrl() != null && !house.getProfileUrl().isEmpty()) {
            linkCell.setCellValue("Открыть карточку");
            Hyperlink link = createHelper.createHyperlink(HyperlinkType.URL);
            link.setAddress(house.getProfileUrl());
            linkCell.setHyperlink(link);
            linkCell.setCellStyle(linkStyle);
        } else {
            linkCell.setCellValue("Нет ссылки");
            linkCell.setCellStyle(defaultStyle);
        }
    }

    private void updateHouseRow(Row row, House house, CellStyle defaultStyle, CellStyle linkStyle, CreationHelper createHelper) {
        setCellValue(row, 1, house.getCommissioningYear(), defaultStyle);
        setCellValue(row, 2, house.getFloorsCount(), defaultStyle);
        setCellValue(row, 3, house.getManagementOrganization(), defaultStyle);
        setCellValue(row, 4, house.getPremisesCount(), defaultStyle);

        Cell linkCell = row.getCell(5);
        if (linkCell == null) {
            linkCell = row.createCell(5);
        }
        if (house.getProfileUrl() != null && !house.getProfileUrl().isEmpty()) {
            linkCell.setCellValue("Открыть карточку");
            Hyperlink link = createHelper.createHyperlink(HyperlinkType.URL);
            link.setAddress(house.getProfileUrl());
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
        if (houses.isEmpty()) {
            notifyLog("❌ Нет данных для сохранения");
            return;
        }

        String fileName = "Объекты жилищного фонда " + region + " " + LocalDate.now().getYear() + ".xlsx";
        boolean fileExists = new File(fileName).exists();

        if (new File("Объекты жилищного фонда " + region + " " + LocalDate.now().minusYears(1).getYear() + ".xlsx").exists()) {
            fileExists = true;
            fileName = "Объекты жилищного фонда " + region + " " + LocalDate.now().minusYears(1).getYear() + ".xlsx";
        } else if (new File("Объекты жилищного фонда " + region + " " + LocalDate.now().getYear() + ".xlsx").exists()) {
            fileExists = true;
        }

        try {
            Workbook workbook;
            Sheet sheet;

            if (fileExists) {
                try (FileInputStream fis = new FileInputStream(fileName)) {
                    workbook = new XSSFWorkbook(fis);
                }
                sheet = workbook.getSheet("Дома");
                if (sheet == null) {
                    sheet = workbook.createSheet("Дома");
                    createHeaders(sheet, workbook);
                }
            } else {
                workbook = new XSSFWorkbook();
                sheet = workbook.createSheet("Дома");
                createHeaders(sheet, workbook);
            }

            CellStyle defaultStyle = createDefaultStyle(workbook);
            CellStyle linkStyle = createLinkStyle(workbook);

            Map<String, Integer> existingHouses = new HashMap<>();
            if (fileExists && sheet.getPhysicalNumberOfRows() > 1) {
                for (int i = 1; i <= sheet.getLastRowNum(); i++) {
                    Row row = sheet.getRow(i);
                    if (row != null && row.getCell(0) != null) {
                        String houseAddress = row.getCell(0).getStringCellValue();
                        if (houseAddress != null && !houseAddress.trim().isEmpty()) {
                            existingHouses.put(houseAddress.trim(), i);
                        }
                    }
                }
            }

            CreationHelper createHelper = workbook.getCreationHelper();
            int newRowsCount = 0;
            int updatedRowsCount = 0;

            for (House house : houses) {
                if (house.getAddress() == null || house.getAddress().trim().isEmpty()) {
                    continue;
                }

                String houseAddress = house.getAddress().trim();
                Integer existingRowIndex = existingHouses.get(houseAddress);

                if (existingRowIndex != null) {
                    updateHouseRow(sheet.getRow(existingRowIndex), house, defaultStyle, linkStyle, createHelper);
                    updatedRowsCount++;
                } else {
                    int newRowIndex = sheet.getLastRowNum() + 1;
                    Row row = sheet.createRow(newRowIndex);
                    createHouseRow(row, house, defaultStyle, linkStyle, createHelper);
                    newRowsCount++;
                }
            }

            for (int i = 0; i < 6; i++) {
                sheet.autoSizeColumn(i);
                sheet.setColumnWidth(i, sheet.getColumnWidth(i) + 512);
            }

            for (int i = 0; i <= sheet.getLastRowNum(); i++) {
                Row row = sheet.getRow(i);
                if (row != null) {
                    row.setHeight((short) -1);
                }
            }

            sheet.setAutoFilter(new CellRangeAddress(0, sheet.getLastRowNum(), 0, 5));

            try (FileOutputStream fos = new FileOutputStream("Объекты жилищного фонда " + region + " " + LocalDate.now().getYear() + ".xlsx")) {
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
        DomGosuslugiHousesParser parser = new DomGosuslugiHousesParser();
        parser.parseHouses();
    }
}