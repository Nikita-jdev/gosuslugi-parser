package com.parser;

import javax.swing.*;
import javax.swing.text.DefaultCaret;
import java.awt.*;
import java.util.concurrent.atomic.AtomicBoolean;

public class ParserGUI extends JFrame implements ProgressListener {
    private final JLabel statusLabel = new JLabel("Готово");
    private final JProgressBar pageProgress = new JProgressBar();
    private final JTextArea logArea = new JTextArea();
    private final JButton startButton = new JButton("Старт");
    private final JButton stopButton = new JButton("Стоп");
    private final AtomicBoolean cancelRequested = new AtomicBoolean(false);

    private Thread workerThread;

    public ParserGUI() {
        super("Парсер управляющих компаний (dom.gosuslugi.ru)");
        setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        setSize(900, 600);
        setLocationRelativeTo(null);

        // Верхняя панель: статус
        JPanel top = new JPanel(new BorderLayout(8, 8));
        top.setBorder(BorderFactory.createEmptyBorder(8, 8, 8, 8));
        top.add(new JLabel("Статус:"), BorderLayout.WEST);
        top.add(statusLabel, BorderLayout.CENTER);

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
        startButton.setEnabled(false);
        stopButton.setEnabled(true);
        cancelRequested.set(false);

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
                DomGosuslugiParser parser = new DomGosuslugiParser();
                parser.setListener(this);
                parser.setCancellationFlag(cancelRequested);
                parser.parseOrganizations();
                msg = "Парсинг завершён";
            } catch (Throwable t) {
                ok = false;
                msg = "Ошибка: " + t.getMessage();
                log("Исключение: " + t.toString());
            } finally {
                onFinished(ok, msg);
            }
        }, "parser-thread");
        workerThread.start();
    }

    private void requestCancel() {
        stopButton.setEnabled(false);
        cancelRequested.set(true);
        onStatus("Остановка по запросу...");
        log("Пользователь запросил остановку.");
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