package com.parser;

public interface ProgressListener {
    // Краткий статус этапа
    void onStatus(String text);

    // Прогресс по страницам; total < 0 означает неизвестное количество и indeterminate
    void onPageProgress(int current, int total);

    // Строка лога
    void log(String line);

    // Сигнал об окончании (успех/ошибка)
    void onFinished(boolean success, String message);
}