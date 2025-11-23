package com.parser;

import javax.swing.*;
import javax.swing.text.DefaultCaret;
import java.awt.*;
import java.util.List;
import java.util.concurrent.atomic.AtomicBoolean;

public class ParserGUI extends JFrame implements ProgressListener {
    private final JLabel statusLabel = new JLabel("Готово");
    private final JProgressBar pageProgress = new JProgressBar();
    private final JTextArea logArea = new JTextArea();
    private final JButton startButton = new JButton("Старт");
    private final JButton stopButton = new JButton("Стоп");
    private final JSpinner startPageSpinner = new JSpinner(new SpinnerNumberModel(1, 1, 1000, 1));
    private final JComboBox<String> parserComboBox = new JComboBox<>();
    private final AtomicBoolean cancelRequested = new AtomicBoolean(false);

    private Thread workerThread;
    private String selectedRegion; // Выбранный пользователем регион

    public ParserGUI() {
        super("Парсер управляющих компаний (dom.gosuslugi.ru)");
        setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        setSize(900, 600);
        setLocationRelativeTo(null);

        // Верхняя панель: статус и настройки
        JPanel top = new JPanel(new BorderLayout(8, 8));
        top.setBorder(BorderFactory.createEmptyBorder(8, 8, 8, 8));

        // Панель с настройками
        JPanel settingsPanel = new JPanel(new FlowLayout(FlowLayout.LEFT));

        // Выбор парсера
        settingsPanel.add(new JLabel("Тип парсера:"));
        parserComboBox.addItem("Реестр объектов жилищного фонда");
        parserComboBox.addItem("Реестры поставщиков информации");
        parserComboBox.setToolTipText("Выберите тип данных для парсинга");
        settingsPanel.add(parserComboBox);

        // Стартовая страница
        settingsPanel.add(new JLabel("Начать со страницы:"));
        startPageSpinner.setToolTipText("Номер страницы для начала парсинга (по умолчанию: 1)");
        startPageSpinner.setPreferredSize(new Dimension(80, 25));
        settingsPanel.add(startPageSpinner);

        top.add(settingsPanel, BorderLayout.NORTH);

        JPanel statusPanel = new JPanel(new FlowLayout(FlowLayout.LEFT));
        statusPanel.add(new JLabel("Статус:"));
        statusPanel.add(statusLabel);
        top.add(statusPanel, BorderLayout.SOUTH);

        // Прогресс-бар: настройка для отображения прогресса по страницам
        pageProgress.setStringPainted(true);
        pageProgress.setIndeterminate(false);
        pageProgress.setMinimum(0);
        pageProgress.setMaximum(100); // Проценты по умолчанию
        pageProgress.setValue(0);
        pageProgress.setString("Ожидание начала...");
        pageProgress.setToolTipText("Прогресс парсинга страниц");

        // Логи
        logArea.setEditable(false);
        logArea.setFont(new Font(Font.MONOSPACED, Font.PLAIN, 12));
        JScrollPane scroll = new JScrollPane(logArea);
        // Автопрокрутка вниз
        DefaultCaret caret = (DefaultCaret) logArea.getCaret();
        caret.setUpdatePolicy(DefaultCaret.ALWAYS_UPDATE);

        // Кнопки
        startButton.setToolTipText("Запустить парсинг");
        stopButton.setToolTipText("Остановить парсинг");
        JPanel buttons = new JPanel(new FlowLayout(FlowLayout.RIGHT));
        buttons.add(startButton);
        buttons.add(stopButton);
        stopButton.setEnabled(false);

        JPanel center = new JPanel(new BorderLayout(8, 8));
        center.setBorder(BorderFactory.createEmptyBorder(0, 8, 8, 8));
        center.add(pageProgress, BorderLayout.NORTH);
        center.add(scroll, BorderLayout.CENTER);

        setLayout(new BorderLayout());
        add(top, BorderLayout.NORTH);
        add(center, BorderLayout.CENTER);
        add(buttons, BorderLayout.SOUTH);

        // Действия кнопок
        startButton.addActionListener(e -> startParsing());
        stopButton.addActionListener(e -> requestCancel());
    }

    private void startParsing() {
        cleanupSystem();

        startButton.setEnabled(false);
        stopButton.setEnabled(true);
        cancelRequested.set(false);
        selectedRegion = null; // Сбрасываем выбранный регион

        // Получаем выбранные настройки
        int startPage = (Integer) startPageSpinner.getValue();
        String selectedParser = (String) parserComboBox.getSelectedItem();

        // Сброс прогресса
        SwingUtilities.invokeLater(() -> {
            pageProgress.setIndeterminate(false);
            pageProgress.setMinimum(0);
            pageProgress.setMaximum(100);
            pageProgress.setValue(0);
            pageProgress.setString("Подготовка к парсингу...");
            statusLabel.setText("Подготовка к запуску...");
            logArea.setText(""); // Очищаем логи при новом запуске
        });

        workerThread = new Thread(() -> {
            boolean ok = true;
            String msg = "Готово";
            try {
                if ("Реестр объектов жилищного фонда".equals(selectedParser)) {
                    DomGosuslugiHousesParser parser = new DomGosuslugiHousesParser();
                    parser.setListener(this);
                    parser.setCancellationFlag(cancelRequested);
                    parser.setStartPage(startPage);
                    parser.parseHouses();
                    msg = "Парсинг объектов жилищного фонда завершён";
                } else {
                    DomGosuslugiParser parser = new DomGosuslugiParser();
                    parser.setListener(this);
                    parser.setCancellationFlag(cancelRequested);
                    parser.setStartPage(startPage);
                    parser.parseOrganizations();
                    msg = "Парсинг поставщиков информации завершён";
                }
            } catch (Throwable t) {
                ok = false;
                msg = "Ошибка: " + t.getMessage();
                log("Исключение: " + t.toString());
            } finally {
                cleanupSystem();
                onFinished(ok, msg);
            }
        }, "parser-thread");
        workerThread.start();
    }

    // Новый метод для отображения диалога выбора региона
    @Override
    public String showRegionSelectionDialog(List<String> regions) {
        try {
            // Создаем диалоговое окно для выбора региона
            final String[] result = {null};

            SwingUtilities.invokeAndWait(() -> {
                JDialog dialog = new JDialog(this, "Выбор региона", true);
                dialog.setLayout(new BorderLayout());
                dialog.setSize(400, 500);
                dialog.setLocationRelativeTo(this);

                JPanel contentPanel = new JPanel(new BorderLayout(10, 10));
                contentPanel.setBorder(BorderFactory.createEmptyBorder(10, 10, 10, 10));

                // Заголовок
                JLabel titleLabel = new JLabel("Выберите регион для парсинга:");
                titleLabel.setFont(new Font(Font.SANS_SERIF, Font.BOLD, 14));
                contentPanel.add(titleLabel, BorderLayout.NORTH);

                // Список регионов
                JList<String> regionList = new JList<>(regions.toArray(new String[0]));
                regionList.setSelectionMode(ListSelectionModel.SINGLE_SELECTION);
                regionList.setFont(new Font(Font.SANS_SERIF, Font.PLAIN, 12));

                JScrollPane scrollPane = new JScrollPane(regionList);
                scrollPane.setPreferredSize(new Dimension(350, 350));
                contentPanel.add(scrollPane, BorderLayout.CENTER);

                // Панель кнопок
                JPanel buttonPanel = new JPanel(new FlowLayout());
                JButton okButton = new JButton("OK");
                JButton cancelButton = new JButton("Отмена");

                okButton.addActionListener(e -> {
                    String selected = regionList.getSelectedValue();
                    if (selected != null) {
                        result[0] = selected;
                        dialog.dispose();
                    } else {
                        JOptionPane.showMessageDialog(dialog, "Пожалуйста, выберите регион", "Внимание", JOptionPane.WARNING_MESSAGE);
                    }
                });

                cancelButton.addActionListener(e -> {
                    result[0] = null;
                    dialog.dispose();
                });

                buttonPanel.add(okButton);
                buttonPanel.add(cancelButton);
                contentPanel.add(buttonPanel, BorderLayout.SOUTH);

                dialog.add(contentPanel);
                dialog.setVisible(true);
            });

            selectedRegion = result[0];
            return result[0];

        } catch (Exception e) {
            log("❌ Ошибка при выборе региона: " + e.getMessage());
            return null;
        }
    }

    private void requestCancel() {
        stopButton.setEnabled(false);
        cancelRequested.set(true);
        onStatus("Остановка по запросу...");
        log("⏹️ Пользователь запросил остановку. Завершаем текущие операции...");

        // Принудительно прерываем рабочий поток
        if (workerThread != null && workerThread.isAlive()) {
            workerThread.interrupt();
            log("⚠️ Отправлен сигнал прерывания потока");
        }
    }

    private void cleanupSystem() {
        // Принудительный вызов сборщика мусора
        System.gc();
        System.runFinalization();

        log("🧹 Системная очистка памяти выполнена");
    }

    // ProgressListener implementation
    @Override
    public void onStatus(String text) {
        SwingUtilities.invokeLater(() -> statusLabel.setText(text));
    }

    @Override
    public void onPageProgress(int current, int total) {
        SwingUtilities.invokeLater(() -> {
            if (total <= 0) {
                // Если общее количество страниц неизвестно
                pageProgress.setIndeterminate(true);
                pageProgress.setString("Страница " + current + " (всего: определяется...)");
            } else {
                // Режим с известным общим количеством страниц
                pageProgress.setIndeterminate(false);
                pageProgress.setMinimum(0);
                pageProgress.setMaximum(total);
                pageProgress.setValue(current);

                // Вычисляем процент выполнения
                int percent = (int) Math.round((double) current / total * 100);
                pageProgress.setString(String.format("Страница %d из %d (%d%%)", current, total, percent));
            }
        });
    }

    @Override
    public void log(String line) {
        SwingUtilities.invokeLater(() -> {
            logArea.append(line + System.lineSeparator());
        });
    }

    @Override
    public void onFinished(boolean success, String message) {
        SwingUtilities.invokeLater(() -> {
            startButton.setEnabled(true);
            stopButton.setEnabled(false);

            // Финализируем прогресс-бар
            pageProgress.setIndeterminate(false);
            if (success) {
                pageProgress.setValue(pageProgress.getMaximum());
                pageProgress.setString("Завершено - " + message);
            } else {
                pageProgress.setString("Прервано - " + message);
            }

            onStatus(message + (success ? "" : " (см. лог)"));

            if (!success) {
                JOptionPane.showMessageDialog(this, message, "Ошибка", JOptionPane.ERROR_MESSAGE);
            } else {
                JOptionPane.showMessageDialog(this, message, "Готово", JOptionPane.INFORMATION_MESSAGE);
            }
        });
    }

    public static void main(String[] args) {
        SwingUtilities.invokeLater(() -> {
            ParserGUI gui = new ParserGUI();
            gui.setVisible(true);
        });
    }
}